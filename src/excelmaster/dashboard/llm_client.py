"""Provider-agnostic LLM client with JSON repair and retry."""
from __future__ import annotations
import json
import re
import time
from typing import Any

from openai import OpenAI

from ..config import get_settings


MINIMAX_BASE_URL = "https://api.minimax.chat/v1/"


class LLMClient:
    def __init__(self):
        cfg = get_settings()
        self.provider = cfg.llm_provider
        self.model = cfg.active_model
        self.max_tokens = cfg.max_tokens
        self.temperature = cfg.temperature
        self.max_retries = cfg.max_retries

        if self.provider == "minimax":
            self.client = OpenAI(
                api_key=cfg.minimax_api_token,
                base_url=MINIMAX_BASE_URL,
            )
        else:
            self.client = OpenAI(api_key=cfg.openai_api_key)

    def generate_json(self, system_prompt: str, user_prompt: str,
                       context: str = "",
                       max_tokens_override: int | None = None) -> dict[str, Any]:
        """Call LLM and return parsed JSON. Retries on failure."""
        full_user = f"{context}\n\n{user_prompt}" if context else user_prompt
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": full_user},
        ]
        tokens = max_tokens_override or self.max_tokens
        last_err = None
        for attempt in range(self.max_retries):
            try:
                resp = self.client.chat.completions.create(
                    model=self.model,
                    messages=messages,
                    max_tokens=tokens,
                    temperature=self.temperature,
                )
                raw = resp.choices[0].message.content or ""
                return self._parse_json(raw)
            except Exception as e:
                last_err = e
                if attempt < self.max_retries - 1:
                    time.sleep(2 ** attempt)
        raise RuntimeError(f"LLM failed after {self.max_retries} attempts: {last_err}")

    def generate_chat_json(
        self,
        messages: list[dict[str, str]],
        max_tokens_override: int | None = None,
    ) -> dict[str, Any]:
        """Call LLM with a full messages list and return parsed JSON.

        Used by the chat engine where conversation history is maintained
        as a list of system/user/assistant messages.
        """
        tokens = max_tokens_override or self.max_tokens
        last_err = None
        for attempt in range(self.max_retries):
            try:
                resp = self.client.chat.completions.create(
                    model=self.model,
                    messages=messages,
                    max_tokens=tokens,
                    temperature=self.temperature,
                )
                raw = resp.choices[0].message.content or ""
                return self._parse_json_or_object(raw)
            except Exception as e:
                last_err = e
                if attempt < self.max_retries - 1:
                    time.sleep(2 ** attempt)
        raise RuntimeError(f"LLM failed after {self.max_retries} attempts: {last_err}")

    def _parse_json_or_object(self, raw: str) -> dict[str, Any]:
        """Parse JSON that may be a top-level object or array."""
        # Remove <think>...</think> blocks
        raw = re.sub(r"<think>.*?</think>", "", raw, flags=re.DOTALL).strip()

        # Extract JSON from markdown code blocks
        for pattern in [r"```json\s*(.*?)```", r"```\s*(.*?)```"]:
            m = re.search(pattern, raw, re.DOTALL)
            if m:
                raw = m.group(1).strip()
                break

        # Detect top-level arrays vs objects
        stripped = raw.strip()
        if stripped.startswith("["):
            start = stripped.find("[")
            end = stripped.rfind("]")
            if start != -1 and end != -1:
                raw = stripped[start:end + 1]
            try:
                arr = json.loads(raw)
                return {"message": "", "actions": arr}
            except json.JSONDecodeError:
                raw = self._repair_json(raw)
                try:
                    arr = json.loads(raw)
                    return {"message": "", "actions": arr}
                except json.JSONDecodeError:
                    pass

        # Fall back to normal object parsing
        return self._parse_json(raw)

    def generate_text(self, system_prompt: str, user_prompt: str) -> str:
        """Call LLM and return plain text."""
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ]
        resp = self.client.chat.completions.create(
            model=self.model,
            messages=messages,
            max_tokens=self.max_tokens,
            temperature=self.temperature,
        )
        return resp.choices[0].message.content or ""

    # ── JSON Repair ────────────────────────────────────────────────────────────

    def _parse_json(self, raw: str) -> dict[str, Any]:
        """Extract and repair JSON from LLM response."""
        # Remove <think>...</think> blocks
        raw = re.sub(r"<think>.*?</think>", "", raw, flags=re.DOTALL).strip()

        # Extract JSON from markdown code blocks
        for pattern in [r"```json\s*(.*?)```", r"```\s*(.*?)```"]:
            m = re.search(pattern, raw, re.DOTALL)
            if m:
                raw = m.group(1).strip()
                break

        # Find the outermost JSON object
        start = raw.find("{")
        end = raw.rfind("}")
        if start != -1 and end != -1:
            raw = raw[start:end + 1]

        # Try direct parse
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            pass

        # Repair common issues
        raw = self._repair_json(raw)
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            # Last resort: partial recovery
            return self._recover_partial(raw)

    def _repair_json(self, s: str) -> str:
        """Fix common JSON formatting errors."""
        # Remove trailing commas before } or ]
        s = re.sub(r",\s*([}\]])", r"\1", s)
        # Replace single quotes with double quotes for string values/keys
        # (careful not to break apostrophes in values)
        s = re.sub(r"(?<![a-zA-Z])'([^']*)'(?![a-zA-Z])", r'"\1"', s)
        # Fix Python True/False/None → JSON true/false/null
        s = re.sub(r"\bTrue\b", "true", s)
        s = re.sub(r"\bFalse\b", "false", s)
        s = re.sub(r"\bNone\b", "null", s)
        # Remove comments
        s = re.sub(r"//.*?\n", "\n", s)
        s = re.sub(r"/\*.*?\*/", "", s, flags=re.DOTALL)
        return s

    def _recover_partial(self, s: str) -> dict[str, Any]:
        """Try to recover from partial/truncated JSON."""
        # Add missing closing brackets
        open_braces = s.count("{") - s.count("}")
        open_brackets = s.count("[") - s.count("]")
        s = s.rstrip().rstrip(",")
        s += "]" * max(open_brackets, 0)
        s += "}" * max(open_braces, 0)
        try:
            return json.loads(s)
        except json.JSONDecodeError as e:
            raise ValueError(f"Cannot parse JSON response: {e}\nRaw (first 500): {s[:500]}")
