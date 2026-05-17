"""Gemini review for low-confidence EN/ZH paragraph alignments."""

from __future__ import annotations

import json
import os
import re
import time
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

DEFAULT_MODEL = "gemini-2.0-flash"
MAX_CHARS_PER_SIDE = 6000
REVIEW_SCHEMA = {
    "type": "object",
    "properties": {
        "verdict": {
            "type": "string",
            "enum": ["match", "partial", "mismatch"],
            "description": "Whether the English and Chinese paragraphs refer to the same rule/definition.",
        },
        "explanation": {
            "type": "string",
            "description": "Brief reason for the verdict (1-2 sentences, English).",
        },
    },
    "required": ["verdict", "explanation"],
}

REVIEW_PROMPT = """You are reviewing aligned paragraphs from the HKEX listing rules (English) and its Traditional Chinese translation.

The paragraphs were paired automatically for translation QA. Your job is to decide whether they refer to the same rule, definition, or substantive content.

Guidelines:
- "match": same term/definition/rule; minor wording or layout differences are fine.
- "partial": overlapping topic but clearly not the same paragraph (e.g. wrong pairing, mixed definitions).
- "mismatch": unrelated content on one or both sides.

Chinese may embed English terms in parentheses, e.g. "會計及財務匯報局"(Accounting and Financial Reporting Council)或(AFRC).

English paragraph:
{english}

Chinese paragraph:
{chinese}

Respond with JSON only."""


@dataclass(frozen=True)
class LlmReviewResult:
    verdict: str
    explanation: str


def load_env(project_root: Path | None = None) -> None:
    root = project_root or Path(__file__).resolve().parent
    load_dotenv(root / ".env")


def get_api_key() -> str | None:
    load_env()
    return os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")


def llm_available() -> bool:
    return bool(get_api_key())


def _truncate(text: str, limit: int = MAX_CHARS_PER_SIDE) -> str:
    text = text.strip()
    if len(text) <= limit:
        return text
    return text[: limit - 3] + "..."


def _parse_review_response(raw: str) -> LlmReviewResult:
    text = raw.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    data = json.loads(text)
    verdict = str(data.get("verdict", "partial")).lower()
    if verdict not in ("match", "partial", "mismatch"):
        verdict = "partial"
    explanation = str(data.get("explanation", "")).strip() or "No explanation provided."
    return LlmReviewResult(verdict=verdict, explanation=explanation)


def review_pair(
    english: str,
    chinese: str,
    *,
    client=None,
    model: str = DEFAULT_MODEL,
    max_retries: int = 3,
) -> LlmReviewResult:
    """Ask Gemini whether an EN/ZH paragraph pair is a valid translation match."""
    if client is None:
        from google import genai
        from google.genai import types

        api_key = get_api_key()
        if not api_key:
            raise RuntimeError(
                "GEMINI_API_KEY or GOOGLE_API_KEY not set. Add one to the project .env file."
            )
        client = genai.Client(api_key=api_key)
    else:
        from google.genai import types

    prompt = REVIEW_PROMPT.format(
        english=_truncate(english),
        chinese=_truncate(chinese),
    )
    config = types.GenerateContentConfig(
        temperature=0.1,
        response_mime_type="application/json",
        response_schema=REVIEW_SCHEMA,
    )

    last_error: Exception | None = None
    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model=model,
                contents=prompt,
                config=config,
            )
            return _parse_review_response(response.text or "")
        except Exception as exc:
            last_error = exc
            if attempt < max_retries - 1:
                time.sleep(1.5 * (attempt + 1))
    raise RuntimeError(f"Gemini review failed after {max_retries} attempts: {last_error}")


def review_low_confidence_rows(
    rows: list[tuple[str, str, float | None, str]],
    *,
    threshold: float = 0.55,
    model: str = DEFAULT_MODEL,
    on_progress: Callable[[int, int, int], None] | None = None,
) -> list[tuple[str, str, float | None, str, str, str]]:
    """
    Add llm_verdict and llm_explanation to rows with match_score < threshold (both sides present).
    Returns rows as (en, zh, score, note, llm_verdict, llm_explanation).
    """
    from google import genai

    api_key = get_api_key()
    if not api_key:
        raise RuntimeError(
            "GEMINI_API_KEY or GOOGLE_API_KEY not set. Add one to the project .env file."
        )
    client = genai.Client(api_key=api_key)

    enriched: list[tuple[str, str, float | None, str, str, str]] = []
    review_indices = {
        i
        for i, (en, zh, score, _) in enumerate(rows)
        if en and zh and score is not None and score < threshold
    }
    total = len(review_indices)
    reviewed = 0

    for out_idx, (en, zh, score, note) in enumerate(rows):
        if out_idx in review_indices:
            reviewed += 1
            if on_progress:
                on_progress(reviewed, total, out_idx + 1)
            try:
                result = review_pair(en, zh, client=client, model=model)
                llm_verdict = result.verdict
                llm_explanation = result.explanation
            except Exception as exc:
                llm_verdict = "error"
                llm_explanation = str(exc)[:500]
            enriched.append((en, zh, score, note, llm_verdict, llm_explanation))
        else:
            enriched.append((en, zh, score, note, "", ""))

    return enriched
