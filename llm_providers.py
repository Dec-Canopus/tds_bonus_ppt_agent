import os, json, re
import requests

class ProviderError(Exception):
    pass

# Centralized safe defaults per provider (used when model is blank)
PROVIDER_DEFAULTS = {
    "openai":       "gpt-4o-mini",
    "gemini":       "gemini-2.5-pro",
    "google":       "gemini-2.5-pro",
    "google-gemini":"gemini-2.5-pro",
    "aipipe":       "gpt-4o-mini",
}

SYSTEM_PROMPT = """You are a slide planner. Convert the provided text into a JSON slide plan.
Follow these rules:
- Respect the 'guidance' string for tone/structure if provided.
- Choose a reasonable number of slides (min 4, max 30) based on content.
- Each slide must include: title (string), bullets (array of short strings).
- If include_notes=true, add an optional notes field (string) for speaker notes.
- Use concise, scannable bullets. Avoid paragraphs.
- Do not include any images or graphics; the app will reuse template images itself.
- Output strictly valid JSON only, matching this schema:

{
  "slides": [
    {
      "title": "string",
      "bullets": ["string", "string", "..."],
      "layout_hint": "title_and_content|title_only|section_header|two_content|quote|comparison|timeline|process|overview|summary",
      "notes": "optional string"
    }
  ]
}
"""

USER_TEMPLATE = """GUIDANCE (optional): {guidance}

SOURCE TEXT:
{input_text}
"""

# ---------- Provider POST helpers ----------

def _post_openai(api_key, model, system, user):
    url = "https://api.openai.com/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": model or PROVIDER_DEFAULTS["openai"],
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user}
        ],
        "temperature": 0.2,
        "response_format": {"type": "json_object"}
    }
    r = requests.post(url, headers=headers, json=payload, timeout=60)
    if r.status_code >= 400:
        raise ProviderError(f"OpenAI API error {r.status_code}: {r.text[:200]}")
    data = r.json()
    try:
        return data["choices"][0]["message"]["content"]
    except Exception:
        raise ProviderError("OpenAI response missing content")

def _post_aipipe(api_key, model, system, user):
    """
    AI-Pipe proxy is OpenAI-compatible but some routes 400 if response_format is used.
    We'll try once WITH response_format, and if 400, retry WITHOUT it.
    """
    url = "https://aipipe.org/openai/v1/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}

    payload = {
        "model": model or PROVIDER_DEFAULTS["aipipe"],
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user}
        ],
        "temperature": 0.2,
        "response_format": {"type": "json_object"},
    }
    r = requests.post(url, headers=headers, json=payload, timeout=60)
    if r.status_code == 400:
        # Retry without response_format
        payload.pop("response_format", None)
        r = requests.post(url, headers=headers, json=payload, timeout=60)

    if r.status_code >= 400:
        raise ProviderError(f"AI-Pipe API error {r.status_code}: {r.text[:200]}")
    data = r.json()
    try:
        return data["choices"][0]["message"]["content"]
    except Exception:
        raise ProviderError("AI-Pipe response missing content")

def _post_gemini(api_key, model, system, user):
    model_name = model or PROVIDER_DEFAULTS["gemini"]
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
    headers = {"content-type": "application/json"}
    # Concatenate system+user into a single user message; Gemini doesn't use system separately in v1beta.
    prompt = f"{system}\n\nUser Input:\n{user}\n\nReturn STRICT JSON only."
    payload = {
        "contents": [{"role": "user", "parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.2}
    }
    r = requests.post(url, headers=headers, json=payload, timeout=60)
    if r.status_code >= 400:
        raise ProviderError(f"Gemini API error {r.status_code}: {r.text[:200]}")
    data = r.json()
    try:
        return data["candidates"][0]["content"]["parts"][0]["text"]
    except Exception:
        raise ProviderError("Gemini response missing text content")

# ---------- JSON coercion / safety ----------

_FENCE_RE = re.compile(r"```(?:json)?\s*(.*?)\s*```", re.S)
_THINK_RE = re.compile(r"<think>.*?</think>", re.S)

def _coerce_json(text: str):
    """
    Accepts raw model text, extracts fenced JSON if present, strips Perplexity <think> blocks,
    and parses JSON. Also trims accidental prefix/suffix around a JSON object.
    """
    if not isinstance(text, str):
        raise ProviderError("Provider returned non-text output")

    s = text.strip()
    # Remove Perplexity <think>...</think> if present
    s = _THINK_RE.sub("", s)

    # Prefer fenced ```json ... ``` or ``` ... ``` blocks
    m = _FENCE_RE.search(s)
    if m:
        s = m.group(1).strip()

    # If the string doesn't start with '{', try to locate the first JSON object
    if not s.startswith("{"):
        m2 = re.search(r"\{.*\}", s, flags=re.S)
        if m2:
            s = m2.group(0).strip()

    try:
        return json.loads(s)
    except json.JSONDecodeError as e:
        raise ProviderError(f"Provider returned non-JSON output: {e}")

# ---------- Public entrypoint ----------

def _validate_api_key_like(s: str):
    bad_substrings = [" ", "http", "Bearer ", "provider.lower()", "elif "]
    if not s or any(x in s for x in bad_substrings) or len(s) < 20:
        raise ProviderError("API key looks invalid. Paste only your provider token (no quotes/Bearer/spaces).")

def plan_slides_via_llm(provider, model, api_key, input_text, guidance, include_notes):
    _validate_api_key_like(api_key)
    p = (provider or "").lower()
    user = USER_TEMPLATE.format(guidance=guidance or "(none)", input_text=input_text[:15000])

    if p == "openai":
        raw = _post_openai(api_key, model, SYSTEM_PROMPT, user)
    elif p in ("google", "gemini", "google-gemini"):
        raw = _post_gemini(api_key, model, SYSTEM_PROMPT, user)
    elif p == "aipipe":
        raw = _post_aipipe(api_key, model, SYSTEM_PROMPT, user)
    else:
        raise ProviderError(f"Unsupported provider: {provider}")

    data = _coerce_json(raw)
    slides = data.get("slides") or []
    if not isinstance(slides, list) or len(slides) == 0:
        raise ProviderError("No slides returned by provider.")

    # Normalize
    slides = slides[:30]
    if include_notes:
        for s in slides:
            if "notes" not in s:
                s["notes"] = ""
    for s in slides:
        if not s.get("layout_hint"):
            s["layout_hint"] = "title_and_content"
    return slides
