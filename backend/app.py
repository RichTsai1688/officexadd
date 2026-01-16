from flask import Flask, request, jsonify
from flask_cors import CORS
import json
import openai
import os
from urllib import request as url_request
from urllib import error as url_error
from urllib.parse import urlparse
from dotenv import load_dotenv

load_dotenv()  # Load environment variables from .env file

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

# Initialize OpenAI client

# Helper configuration for provider selection
from openai import OpenAI
LEGACY_BASE_URL = os.getenv("BASE_URL") or ""
LEGACY_API_KEY = os.getenv("API_KEY") or ""
OLLAMA_BASE_URL = os.getenv("AI_BASE_URL") or LEGACY_BASE_URL
OLLAMA_API_KEY = os.getenv("AI_API_KEY") or LEGACY_API_KEY
OLLAMA_WEB_SEARCH_API_KEY = os.getenv("OLLAMA_WEB_SEARCH_API_KEY") or os.getenv("ollama_web_search_api_key") or ""
OLLAMA_WEB_SEARCH_URL = os.getenv("OLLAMA_WEB_SEARCH_URL") or ""
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY") or LEGACY_API_KEY
MODEL_NAME = os.getenv("MODEL_NAME") or ""
DEFAULT_MODELS = {
    "openai": "gpt-4o-mini",
    "ollama": "llama3.1",
}

def build_client(base_url: str | None = None, api_key: str | None = None):
    kwargs = {}
    if base_url:
        kwargs["base_url"] = base_url
    if api_key:
        kwargs["api_key"] = api_key
    return OpenAI(**kwargs)

def extract_response_text(response):
    if isinstance(response, dict):
        message = response.get("message")
        if isinstance(message, dict) and message.get("content"):
            return message["content"]
        if response.get("response"):
            return response["response"]
        choices = response.get("choices")
        if isinstance(choices, list) and choices:
            content = choices[0].get("message", {}).get("content")
            if content:
                return content
    if hasattr(response, "output_text"):
        return response.output_text
    if hasattr(response, "choices"):
        return response.choices[0].message.content
    output = getattr(response, "output", None)
    if output:
        texts = []
        for item in output:
            content = getattr(item, "content", None)
            if not content:
                continue
            for block in content:
                text = getattr(block, "text", None)
                if text:
                    texts.append(text)
        if texts:
            return "\n".join(texts)
    return ""

def run_openai_web_search(client, model_name, messages):
    response = client.responses.create(
        model=model_name,
        tools=[{"type": "web_search"}],
        input=messages
    )
    return extract_response_text(response)

def parse_tool_arguments(arguments):
    if not arguments:
        return {}
    try:
        return json.loads(arguments)
    except json.JSONDecodeError:
        return {}

def build_tool_call_dicts(tool_calls):
    tool_call_dicts = []
    for call in tool_calls:
        if isinstance(call, dict):
            call_id = call.get("id") or ""
            func = call.get("function") or {}
            tool_call_dicts.append({
                "id": call_id,
                "type": call.get("type") or "function",
                "function": {
                    "name": func.get("name") or "",
                    "arguments": func.get("arguments") or "",
                },
            })
        else:
            tool_call_dicts.append({
                "id": call.id,
                "type": "function",
                "function": {
                    "name": call.function.name,
                    "arguments": call.function.arguments,
                },
            })
    return tool_call_dicts

def run_ollama_web_search_function(query):
    search_api_key = OLLAMA_WEB_SEARCH_API_KEY or OLLAMA_API_KEY
    if not search_api_key:
        raise RuntimeError("OLLAMA_WEB_SEARCH_API_KEY or AI_API_KEY must be configured for web search.")
    urls = []
    if OLLAMA_WEB_SEARCH_URL:
        urls.append(OLLAMA_WEB_SEARCH_URL)
    else:
        urls.extend([
            "https://ollama.com/api/web_search",
            "https://api.ollama.com/api/web_search",
            "https://api.ollama.com/v1/web-search",
            "https://api.ollama.com/v1/web/search",
        ])
        if OLLAMA_BASE_URL:
            base_root = OLLAMA_BASE_URL.rstrip("/")
            if base_root.endswith("/v1"):
                base_root = base_root[:-3].rstrip("/")
            urls.extend([
                f"{base_root}/api/web_search",
                f"{base_root}/api/web-search",
                f"{base_root}/api/web/search",
                f"{base_root}/v1/web_search",
                f"{base_root}/v1/web-search",
                f"{base_root}/v1/web/search",
            ])

    payload = {"query": query}
    data = json.dumps(payload).encode("utf-8")
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {search_api_key}",
    }

    last_error = None
    for url in urls:
        parsed = urlparse(url)
        if not parsed.scheme or not parsed.netloc:
            last_error = RuntimeError(f"Ollama web search URL is invalid: {url}")
            continue
        req = url_request.Request(url, data=data, headers=headers)
        try:
            with url_request.urlopen(req, timeout=90) as resp:
                return resp.read().decode("utf-8")
        except url_error.HTTPError as e:
            body = e.read().decode("utf-8", "ignore")
            last_error = RuntimeError(f"Ollama web search failed: {e.code} {body} (url: {url})")
        except url_error.URLError as e:
            last_error = RuntimeError(f"Ollama web search failed: {e.reason} (url: {url})")

    if last_error:
        raise last_error
    raise RuntimeError("Ollama web search URL is not configured.")

def run_ollama_web_search_tool_flow(client, model_name, messages):
    tools = [{
        "type": "function",
        "function": {
            "name": "web_search",
            "description": "Search the web for relevant, recent information.",
            "parameters": {
                "type": "object",
                "properties": {
                    "query": {"type": "string", "description": "Search query"}
                },
                "required": ["query"],
            },
        },
    }]
    response = client.chat.completions.create(
        model=model_name,
        messages=messages,
        tools=tools
    )
    message = response.choices[0].message
    tool_calls = getattr(message, "tool_calls", None) or message.get("tool_calls") or []
    if not tool_calls:
        return extract_response_text(response)

    tool_call_dicts = build_tool_call_dicts(tool_calls)
    assistant_message = {
        "role": "assistant",
        "content": message.get("content") if isinstance(message, dict) else (message.content or ""),
        "tool_calls": tool_call_dicts,
    }
    tool_messages = []
    for call in tool_calls:
        if isinstance(call, dict):
            function_name = (call.get("function") or {}).get("name")
            arguments = (call.get("function") or {}).get("arguments")
            call_id = call.get("id") or ""
        else:
            function_name = call.function.name
            arguments = call.function.arguments
            call_id = call.id
        if function_name != "web_search":
            continue
        args = parse_tool_arguments(arguments)
        query = args.get("query") or args.get("search_query") or args.get("q") or ""
        if not query:
            tool_result = json.dumps({"error": "Missing search query."})
        else:
            tool_result = run_ollama_web_search_function(query)
        tool_messages.append({
            "role": "tool",
            "tool_call_id": call_id,
            "content": tool_result,
        })

    if not tool_messages:
        return message.content or ""

    followup_messages = messages + [assistant_message] + tool_messages
    final_response = client.chat.completions.create(
        model=model_name,
        messages=followup_messages
    )
    return extract_response_text(final_response)

def run_with_web_search(client, model_name, messages, provider):
    if provider == "ollama":
        return run_ollama_web_search_tool_flow(client, model_name, messages)
    return run_openai_web_search(client, model_name, messages)


def resolve_model_name(provider: str, requested: str | None) -> tuple[str, str | None]:
    """
    Pick a model name appropriate for the provider and fall back if the value
    looks incompatible (e.g., Ollama-style model name used with OpenAI).
    Returns (model_name, warning_note).
    """
    fallback = DEFAULT_MODELS.get(provider, "gpt-4o-mini")
    model = (requested or "").strip()
    if not model:
        return fallback, None

    if provider == "openai" and ":" in model:
        # A colon is common in Ollama model IDs; OpenAI would return model_not_found.
        return fallback, f"Incompatible model '{model}' for provider openai. Fell back to '{fallback}'."

    return model, None


def classify_api_error(error: Exception) -> tuple[str, int]:
    """
    Translate provider errors into client-friendly HTTP status codes.
    """
    message = str(error)
    lowered = message.lower()
    if "model_not_found" in lowered or "does not exist" in lowered:
        return message, 400
    if "rate limit" in lowered or "too many requests" in lowered:
        return message, 429
    if "invalid api key" in lowered or "authentication" in lowered:
        return message, 401
    return message, 500

# client = openai.OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
@app.route('/rewrite', methods=['POST'])
def rewrite_text():
    try:
        data = request.get_json()
        original_text = data.get('text', '')
        instruction = data.get('instruction', '')
        provider = (data.get('provider') or 'openai').strip().lower()
        requested_model = data.get('model') or MODEL_NAME or ''
        model_name, model_warning = resolve_model_name(provider, requested_model)
        use_web_search = bool(data.get('use_web_search'))
        context_mode = (data.get('context_mode') or '').strip().lower()
        context_text = data.get('context_text') or ''
        context_note = data.get('context_note') or ''

        original_text = original_text or ""
        instruction = instruction or ""
        if not instruction.strip() and original_text.strip():
            instruction = 'Rewrite this text in a formal academic tone'
        if not original_text.strip() and not instruction.strip():
            return jsonify({'error': 'No instruction provided'}), 400

        if provider == 'ollama':
            base_url = OLLAMA_BASE_URL
            api_key = OLLAMA_API_KEY
            if not base_url or not api_key:
                return jsonify({'error': 'Ollama configuration is missing (AI_BASE_URL/AI_API_KEY).'}), 500
        else:
            base_url = ''
            api_key = OPENAI_API_KEY
            if not api_key:
                return jsonify({'error': 'OpenAI API key is not configured.'}), 500

        client = build_client(base_url=base_url or None, api_key=api_key or None)

        system_prompt = (
            "Rewrite the user's text according to the instruction and produce HTML fragments "
            "(for example, <p>, <strong>, <em>, <ul>, <li>). Return only the rewritten content without "
            "introductions, explanations, AI commentary, and do not emit <html> or <body> tags."
        )
        if not original_text.strip():
            system_prompt += (
                " If the input text is empty, generate new content that satisfies the instruction and "
                "fits the provided context. Avoid repeating nearby context."
            )
        if context_text:
            system_prompt += (
                " Use the provided document context to keep continuity and avoid repeating content. "
                "The context may contain markers like [[EDIT_START]], [[EDIT_END]], or [[CURSOR]] to show "
                "the rewrite location; never include these markers in the output."
            )
        if use_web_search:
            system_prompt += (
                " Verify factual accuracy using web search. After the rewrite, include a short "
                "'Sources' section with clickable links (HTML list is fine). Do not add any extra commentary."
            )
        else:
            system_prompt += " Do not include citations or source lists."

        user_message = f"Instruction: {instruction}"
        if original_text.strip():
            user_message += f"\n\nText: {original_text}"
        else:
            user_message += "\n\nText: (none)"
        if context_text:
            mode_label = context_mode or "custom"
            user_message += f"\n\nContext ({mode_label}):\n{context_text}"
            if context_note:
                user_message += f"\n\nContext note: {context_note}"
        if model_warning:
            user_message += f"\n\nModel note: {model_warning}"

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message}
        ]

        if use_web_search:
            rewritten_text = run_with_web_search(client, model_name, messages, provider)
            if not rewritten_text or not rewritten_text.strip():
                response = client.chat.completions.create(
                    model=model_name,
                    messages=messages
                )
                rewritten_text = extract_response_text(response)
        else:
            response = client.chat.completions.create(
                model=model_name,
                messages=messages
            )
            rewritten_text = extract_response_text(response)
        if not rewritten_text or not rewritten_text.strip():
            return jsonify({'error': 'Model returned empty output.'}), 502
        response_body = {'rewritten_text': rewritten_text}
        if model_warning:
            response_body['model_note'] = model_warning
        return jsonify(response_body)

    except Exception as e:
        message, status = classify_api_error(e)
        print(f"Error: {message}")  # Log error to console
        return jsonify({'error': message}), status


@app.route('/models', methods=['GET'])
def list_models():
    provider = (request.args.get('provider') or 'openai').strip().lower()

    if provider == 'ollama':
        base_url = OLLAMA_BASE_URL
        api_key = OLLAMA_API_KEY
        if not base_url or not api_key:
            return jsonify({'error': 'Ollama configuration is missing (AI_BASE_URL/AI_API_KEY).'}), 500
    else:
        base_url = ''
        api_key = OPENAI_API_KEY
        if not api_key:
            return jsonify({'error': 'OpenAI API key is not configured.'}), 500

    try:
        client = build_client(base_url=base_url or None, api_key=api_key or None)
        response = client.models.list()
        raw_models = getattr(response, 'data', []) or []
        models = []
        for entry in raw_models:
            model_id = None
            if isinstance(entry, str):
                model_id = entry
            elif hasattr(entry, 'get'):
                model_id = entry.get('id')
            elif hasattr(entry, 'id'):
                model_id = getattr(entry, 'id')
            if model_id:
                models.append(model_id)

        return jsonify({'provider': provider, 'models': models})

    except Exception as e:
        print(f"Error listing models: {str(e)}")
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5010, debug=True)
