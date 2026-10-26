"""
llm_curl_client.py  (nis_booking)
==================================
Shared helper that calls the Azure OpenAI Chat Completions API using
``curl`` via subprocess — mirroring exactly what ``llm_client.bat`` does.

This is needed because Python's ``requests`` library cannot reach the
internal Azure OpenAI endpoint on this VM (SSL/proxy issues), but
``curl`` works correctly.

Usage (from invoice_extractor or debit_credit_pdf_extractor):

    from llm_curl_client import call_llm_curl

    reply_text = call_llm_curl(messages, max_tokens=100)
"""

import json
import os
import subprocess
import tempfile

import config


def call_llm_curl(messages: list, max_tokens: int = 200, timeout: int = 90) -> str:
    """Send *messages* to Azure OpenAI via ``curl`` and return the raw reply text.

    Mirrors the llm_client.bat logic:
      1. Writes the request payload to a temp JSON file.
      2. Calls: curl -s -X POST <endpoint> -H ... -d @<tempfile>
      3. Parses the JSON response and returns choices[0].message.content.

    Raises RuntimeError on curl failure or unexpected response shape.
    """
    endpoint = str(config.LLM_ENDPOINT).rstrip("/")
    deployment = str(config.LLM_DEPLOYMENT)
    api_version = str(config.LLM_API_VERSION)
    api_key = str(config.LLM_API_KEY)

    url = (
        f"{endpoint}/openai/deployments/{deployment}"
        f"/chat/completions?api-version={api_version}"
    )

    payload = {
        "messages": messages,
        "temperature": 0,
        "max_tokens": max_tokens,
    }

    # Write payload to a temp file (same pattern as llm_client.bat's @%PROMPT_FILE%)
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".json", delete=False, encoding="utf-8"
    ) as tmp:
        json.dump(payload, tmp, ensure_ascii=False)
        tmp_path = tmp.name

    try:
        result = subprocess.run(
            [
                "curl",
                "-s",                          # silent (no progress bar)
                "-k",                          # ignore SSL cert errors (same as VM env)
                "-X", "POST", url,
                "-H", "Content-Type: application/json",
                "-H", f"api-key: {api_key}",
                "-d", f"@{tmp_path}",
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",       # curl returns UTF-8; don't use the Windows
            errors="replace",       # cp1252 default which crashes on 0x81 bytes
            timeout=timeout,
        )
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass

    if result.returncode != 0:
        raise RuntimeError(
            f"curl exited with code {result.returncode}: {(result.stderr or '').strip()}"
        )

    raw = (result.stdout or "").strip()
    if not raw:
        raise RuntimeError("curl returned empty response from LLM endpoint.")

    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"LLM response is not valid JSON: {exc}\nRaw: {raw[:500]}")

    # Surface API-level errors (e.g. auth failure, quota)
    if "error" in data:
        err = data["error"]
        raise RuntimeError(
            f"Azure OpenAI error {err.get('code')}: {err.get('message')}"
        )

    try:
        content = data["choices"][0]["message"]["content"] or ""
        return content.strip()
    except (KeyError, IndexError) as exc:
        raise RuntimeError(
            f"Unexpected LLM response shape ({exc}): {raw[:500]}"
        )
