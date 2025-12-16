# backend/llm_client.py

import os
from typing import List, Dict, Optional

from openai import OpenAI

# Gemini SDK (google-genai)
try:
    from google import genai  # type: ignore
    from google.genai import types as genai_types  # type: ignore
except Exception:  # ImportError or others
    genai = None
    genai_types = None

_openai_client: Optional[OpenAI] = None
_perplexity_client: Optional[OpenAI] = None
_gemini_client = None

# You can override these via environment variables:
#   LLM_MODEL=gpt-4.1-mini
#   LLM_PROVIDER=openai | gemini | perplexity
#   GEMINI_MODEL=gemini-2.5-flash
#   PERPLEXITY_MODEL=sonar-small-online
DEFAULT_MODEL = os.getenv("LLM_MODEL", "gpt-4.1-mini")
DEFAULT_PROVIDER = os.getenv("LLM_PROVIDER", "openai").lower()

# Gemini: use a current, supported text model for the Gemini API
# (this matches Google's own Python quickstart).
DEFAULT_GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")

# Perplexity: the official API supports these Sonar models:
#   sonar, sonar-pro, sonar-deep-research, sonar-reasoning, sonar-reasoning-pro
# We'll default to the simple "sonar" model.
DEFAULT_PERPLEXITY_MODEL = os.getenv("PERPLEXITY_MODEL", "sonar")




def _get_openai_client() -> OpenAI:
    """
    Lazily create a standard OpenAI client using OPENAI_API_KEY.
    """
    global _openai_client
    if _openai_client is None:
        api_key = os.getenv("OPENAI_API_KEY")
        if api_key:
            _openai_client = OpenAI(api_key=api_key)
        else:
            # Will fall back to env var that OpenAI client itself reads.
            _openai_client = OpenAI()
    return _openai_client


def _get_perplexity_client() -> OpenAI:
    """
    Lazily create an OpenAI-compatible client pointed at Perplexity's Sonar API.
    Uses PERPLEXITY_API_KEY.
    """
    global _perplexity_client
    if _perplexity_client is None:
        api_key = os.getenv("PERPLEXITY_API_KEY")
        if not api_key:
            raise RuntimeError(
                "PERPLEXITY_API_KEY is not set, but provider='perplexity' was requested."
            )
        _perplexity_client = OpenAI(api_key=api_key, base_url="https://api.perplexity.ai")
    return _perplexity_client



def _get_gemini_client():
    """
    Lazily create a Google GenAI client for Gemini models.

    We explicitly read GEMINI_API_KEY or GOOGLE_API_KEY so that
    errors are obvious if you're not configured.
    """
    global _gemini_client
    if _gemini_client is not None:
        return _gemini_client

    if genai is None:
        raise RuntimeError(
            "google-genai is not installed in this virtualenv, "
            "but provider='gemini' was requested. "
            "Run: pip install google-genai"
        )

    api_key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
    if not api_key:
        raise RuntimeError(
            "GEMINI_API_KEY (or GOOGLE_API_KEY) is not set, "
            "but provider='gemini' was requested. "
            "Set the env var to your Gemini API key."
        )

    _gemini_client = genai.Client(api_key=api_key)
    return _gemini_client


def _call_openai(
    messages: List[Dict[str, str]],
    model: Optional[str],
    temperature: float,
    json_mode: bool,
) -> str:
    client = _get_openai_client()
    used_model = model or DEFAULT_MODEL

    # Some newer models (like gpt-5.x) require default temperature;
    # we only pass temperature for older ones.
    model_requires_default_temp = used_model.startswith("gpt-5.1")

    kwargs: Dict[str, object] = {
        "model": used_model,
        "messages": messages,
    }

    if not model_requires_default_temp:
        kwargs["temperature"] = temperature

    if json_mode:
        # Strongly encourage valid JSON so our parsing doesn't crash.
        kwargs["response_format"] = {"type": "json_object"}

    resp = client.chat.completions.create(**kwargs)
    content = resp.choices[0].message.content or ""
    return content


def _combine_messages_for_gemini(messages: List[Dict[str, str]]) -> str:
    """
    Flatten OpenAI-style chat messages into a single text prompt for Gemini.
    We keep system messages at the top as instructions.
    """
    system_parts = []
    convo_parts = []

    for m in messages:
        role = m.get("role", "")
        content = m.get("content", "")
        if not content:
            continue
        if role == "system":
            system_parts.append(content)
        else:
            # Keep rough role markers so context isn't totally lost
            convo_parts.append(f"{role.upper()}: {content}")

    system_text = "\n\n".join(system_parts).strip()
    convo_text = "\n\n".join(convo_parts).strip()

    if system_text and convo_text:
        return f"SYSTEM INSTRUCTIONS:\n{system_text}\n\nCONVERSATION:\n{convo_text}"
    elif system_text:
        return f"SYSTEM INSTRUCTIONS:\n{system_text}"
    else:
        return convo_text


def _call_gemini(
    messages: List[Dict[str, str]],
    model: Optional[str],
    temperature: float,
    json_mode: bool,
) -> str:
    client = _get_gemini_client()
    used_model = model or DEFAULT_GEMINI_MODEL

    prompt = _combine_messages_for_gemini(messages)

    # Prepare optional config
    config = None
    if genai_types is not None:
        cfg_kwargs: Dict[str, object] = {}
        # Respect temperature if caller set it explicitly
        cfg_kwargs["temperature"] = float(temperature)
        if json_mode:
            # Ask Gemini to treat the output as JSON text
            cfg_kwargs["response_mime_type"] = "application/json"
        config = genai_types.GenerateContentConfig(**cfg_kwargs)

    if config is not None:
        response = client.models.generate_content(
            model=used_model,
            contents=prompt,
            config=config,
        )
    else:
        # Minimal call if types module isn't available for some reason
        response = client.models.generate_content(
            model=used_model,
            contents=prompt,
        )

    # For text-only responses, .text is the primary field
    text = getattr(response, "text", None)
    if isinstance(text, str) and text:
        return text

    # Fallback: try to assemble from parts
    try:
        parts = []
        for cand in getattr(response, "candidates", []) or []:
            content = getattr(cand, "content", None)
            if not content:
                continue
            for part in getattr(content, "parts", []) or []:
                t = getattr(part, "text", None)
                if isinstance(t, str):
                    parts.append(t)
        return "\n".join(parts)
    except Exception:
        # Last resort: string repr
        return str(response)


def _call_perplexity(
    messages: List[Dict[str, str]],
    model: Optional[str],
    temperature: float,
    json_mode: bool,
) -> str:
    """
    Call Perplexity's Sonar API using the OpenAI client with custom base_url.
    """
    client = _get_perplexity_client()
    used_model = model or DEFAULT_PERPLEXITY_MODEL

    kwargs: Dict[str, object] = {
        "model": used_model,
        "messages": messages,
        "temperature": float(temperature),
    }

    # Do NOT set response_format here; Perplexity's docs don't list it yet.
    # We still tell the model to return JSON via the system prompt when json_mode=True.
    resp = client.chat.completions.create(**kwargs)
    content = resp.choices[0].message.content or ""
    return content


def get_chat_completion(
    messages: List[Dict[str, str]],
    model: Optional[str] = None,
    temperature: float = 0.0,
    json_mode: bool = False,
    provider: Optional[str] = None,
) -> str:
    """
    Unified chat completion helper used by the rest of the backend.

    - messages: list of {"role": "system"|"user"|"assistant", "content": str}
    - model: optional explicit model name
    - temperature: sampling temperature (where applicable)
    - json_mode: if True, we strongly encourage the model to output valid JSON
    - provider: "openai" | "gemini" | "perplexity" (defaults to env LLM_PROVIDER)
    """
    used_provider = (provider or DEFAULT_PROVIDER or "openai").lower()

    if used_provider == "openai":
        return _call_openai(messages, model, temperature, json_mode)
    elif used_provider == "gemini":
        return _call_gemini(messages, model, temperature, json_mode)
    elif used_provider == "perplexity":
        return _call_perplexity(messages, model, temperature, json_mode)
    else:
        # Fallback: if something weird is passed, default back to OpenAI
        return _call_openai(messages, model, temperature, json_mode)
