from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
import io
import os
from datetime import datetime
from ppt_builder import build_presentation
from llm_providers import plan_slides_via_llm, ProviderError

# Flask app
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024  # 20 MB upload limit
app.config['UPLOAD_EXTENSIONS'] = ['.pptx', '.potx']

# --- Key format validator ---
def _key_looks_like(provider: str, api_key: str) -> bool:
    p = (provider or "").lower()
    k = api_key or ""
    if p == "openai":
        return k.startswith("sk-")
    if p == "anthropic":
        return k.startswith("sk-ant-")
    if p in ("google", "gemini", "google-gemini"):
        return k.startswith("AIza") or len(k) > 25
    if p == "perplexity":
        return k.startswith("pplx-")
    if p == "aipipe":
        return k.startswith("eyJ") or k.startswith("ap_") or len(k) > 25
    return True

# --- Same-provider fallback models ---
def _fallback_models(provider: str, current: str | None):
    p = (provider or "").lower()
    c = (current or "").strip()
    if p == "openai":
        return [x for x in ["gpt-4o-mini", "gpt-4o", "gpt-3.5-turbo"] if x != c]
    if p == "anthropic":
        return [x for x in ["claude-3-5-sonnet-20240620", "claude-3-5-haiku"] if x != c]
    if p in ("google", "gemini", "google-gemini"):
        return [x for x in ["gemini-1.5-pro", "gemini-1.5-flash"] if x != c]
    if p == "perplexity":
        return [x for x in ["sonar-small-chat", "sonar-medium-chat"] if x != c]
    if p == "aipipe":
        return [x for x in ["gpt-4o-mini", "claude-3-5-sonnet", "gemini-1.5-pro"] if x != c]
    return []


# Security headers
@app.after_request
def add_security_headers(resp):
    resp.headers['X-Content-Type-Options'] = 'nosniff'
    resp.headers['X-Frame-Options'] = 'DENY'
    resp.headers['Referrer-Policy'] = 'no-referrer'
    return resp


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


def _plan_with_fallback(provider, model, api_key, input_text, guidance, include_notes):
    """Try main model, then safe same-provider fallbacks."""
    if not _key_looks_like(provider, api_key):
        raise ProviderError(f"API key does not look like a {provider} key. Please check.")

    def attempt(m):
        return plan_slides_via_llm(
            provider=provider,
            model=m or None,
            api_key=api_key,
            input_text=input_text,
            guidance=guidance,
            include_notes=include_notes
        )

    try:
        return attempt(model)
    except ProviderError as e:
        msg = str(e)
        candidates = _fallback_models(provider, model)
        # Special case: Perplexity invalid → force chat models
        if ("invalid model" in msg.lower() or "not found" in msg.lower()) and provider.lower() == "perplexity":
            candidates = ["sonar-small-chat", "sonar-medium-chat"]

        last_err = msg
        for alt in candidates:
            try:
                return attempt(alt)
            except ProviderError as e2:
                last_err = str(e2)
        raise ProviderError(last_err)


@app.route("/generate", methods=["POST"])
def generate():
    try:
        input_text = request.form.get("inputText", "").strip()
        guidance = request.form.get("guidance", "").strip()
        provider = request.form.get("provider", "openai").strip()
        model = request.form.get("model", "").strip()
        api_key = request.form.get("apiKey", "").strip()
        include_notes = request.form.get("includeNotes", "off") == "on"

        if not input_text:
            return jsonify({"ok": False, "error": "Input text is required."}), 400
        if not api_key:
            return jsonify({"ok": False, "error": "API key is required for the selected provider."}), 400

        # Validate uploaded file
        f = request.files.get("templateFile", None)
        if f is None or f.filename == "":
            return jsonify({"ok": False, "error": "Please upload a .pptx or .potx template/presentation."}), 400
        filename = secure_filename(f.filename)
        ext = os.path.splitext(filename)[1].lower()
        if ext not in app.config['UPLOAD_EXTENSIONS']:
            return jsonify({"ok": False, "error": "Only .pptx or .potx files are supported."}), 400
        template_bytes = f.read()

        # Step 1: Ask LLM
        try:
            slide_plan = _plan_with_fallback(provider, model, api_key, input_text, guidance, include_notes)
        except ProviderError as e:
            return jsonify({"ok": False, "error": f"LLM provider error ({provider}): {e}"}), 400

        # Step 2: Build PPTX
        try:
            out_pptx = build_presentation(template_bytes, slide_plan)
        except Exception as e:
            return jsonify({"ok": False, "error": f"Failed to build PPTX: {e}"}), 500

        # Step 3: Return file
        stamp = datetime.utcnow().strftime("%Y%m%d-%H%M%S")
        out_name = f"text-to-pptx-{stamp}.pptx"
        return send_file(
            io.BytesIO(out_pptx),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=out_name
        )

    except Exception as e:
        return jsonify({"ok": False, "error": f"Unexpected error: {e}"}), 500


@app.route("/preview", methods=["POST"])
def preview():
    try:
        input_text = request.form.get("inputText", "").strip()
        guidance = request.form.get("guidance", "").strip()
        provider = request.form.get("provider", "openai").strip()
        model = request.form.get("model", "").strip()
        api_key = request.form.get("apiKey", "").strip()
        include_notes = request.form.get("includeNotes", "off") == "on"

        if not input_text:
            return jsonify({"ok": False, "error": "Input text is required."}), 400
        if not api_key:
            return jsonify({"ok": False, "error": "API key is required for the selected provider."}), 400

        slide_plan = _plan_with_fallback(provider, model, api_key, input_text, guidance, include_notes)
        return jsonify({"ok": True, "slides": slide_plan})

    except ProviderError as e:
        return jsonify({"ok": False, "error": f"LLM provider error ({provider}): {e}"}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": f"Unexpected error: {e}"}), 500


if __name__ == "__main__":
    port = int(os.getenv("PORT", 8000))
    app.run(host="0.0.0.0", port=port, debug=False)
