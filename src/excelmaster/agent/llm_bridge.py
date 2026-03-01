"""ToolCallingBridge — wraps OpenAI tool-calling API for the agent."""
from __future__ import annotations

import json
import time
from typing import Any

from openai import OpenAI

from ..config import get_settings
from .tools import get_tool_schemas


SYSTEM_PREAMBLE = """You are an AI assistant that builds and modifies Excel dashboards.
You have access to tools for adding charts, tables, KPIs, writing to cells,
formatting, conditional formatting, managing sheets, and more.

RULES:
- Use EXACT column names from the data schema. Never invent column names.
- Use query_workbook to discover object IDs before modifying or removing.
- Keep your text responses concise (1-2 sentences).
- You may call multiple tools in one turn if the user's request requires it.
- For charts, pick appropriate types for the data (bar for categories, line for time, pie for composition).
"""


class ToolCallingBridge:
    """Wraps OpenAI chat completions with tool calling."""

    def __init__(self) -> None:
        cfg = get_settings()
        self.model = cfg.active_model
        self.max_tokens = cfg.max_tokens
        self.temperature = cfg.temperature
        self.max_retries = cfg.max_retries
        self.client = OpenAI(api_key=cfg.openai_api_key)
        self._tools = get_tool_schemas()

    def call_with_tools(
        self,
        messages: list[dict],
    ) -> tuple[str, list[dict]]:
        """Send messages to LLM with tool definitions.

        Returns:
            (assistant_text, tool_calls) where tool_calls is a list of
            {"id": str, "name": str, "arguments": dict}
        """
        last_err = None
        for attempt in range(self.max_retries):
            try:
                resp = self.client.chat.completions.create(
                    model=self.model,
                    messages=messages,
                    tools=self._tools,
                    tool_choice="auto",
                    max_tokens=self.max_tokens,
                    temperature=self.temperature,
                )
                msg = resp.choices[0].message
                text = msg.content or ""
                tool_calls = []
                if msg.tool_calls:
                    for tc in msg.tool_calls:
                        try:
                            args = json.loads(tc.function.arguments)
                        except json.JSONDecodeError:
                            args = {}
                        tool_calls.append({
                            "id": tc.id,
                            "name": tc.function.name,
                            "arguments": args,
                        })
                return text, tool_calls
            except Exception as e:
                last_err = e
                if attempt < self.max_retries - 1:
                    time.sleep(2 ** attempt)

        raise RuntimeError(f"LLM tool-calling failed after {self.max_retries} attempts: {last_err}")

    def send_tool_results(
        self,
        messages: list[dict],
        tool_results: list[dict],
    ) -> tuple[str, list[dict]]:
        """Send tool results back to LLM for follow-up.

        Args:
            messages: Current conversation (including the assistant message with tool_calls)
            tool_results: List of {"tool_call_id": str, "content": str}

        Returns:
            (assistant_text, more_tool_calls)
        """
        # Append tool result messages
        extended = list(messages)
        for tr in tool_results:
            extended.append({
                "role": "tool",
                "tool_call_id": tr["tool_call_id"],
                "content": tr["content"],
            })

        return self.call_with_tools(extended)

    def build_assistant_tool_call_message(
        self,
        text: str,
        tool_calls: list[dict],
    ) -> dict:
        """Build an assistant message with tool_calls for conversation history."""
        tc_list = []
        for tc in tool_calls:
            tc_list.append({
                "id": tc["id"],
                "type": "function",
                "function": {
                    "name": tc["name"],
                    "arguments": json.dumps(tc["arguments"]),
                },
            })
        msg: dict[str, Any] = {"role": "assistant", "content": text}
        if tc_list:
            msg["tool_calls"] = tc_list
        return msg
