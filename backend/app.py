from flask import Flask, request, jsonify
from flask_cors import CORS
import json
import openai
import os
from urllib import request as url_request
from urllib import error as url_error
from urllib.parse import quote, urlparse
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
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY") or ""
GOOGLE_IMAGE_MODEL = os.getenv("GOOGLE_IMAGE_MODEL") or "gemini-3.1-flash-image-preview"
GOOGLE_IMAGE_ASPECT_RATIO = os.getenv("GOOGLE_IMAGE_ASPECT_RATIO") or "1:1"
GOOGLE_IMAGE_SIZE = os.getenv("GOOGLE_IMAGE_SIZE") or ""
GOOGLE_API_BASE_URL = os.getenv("GOOGLE_API_BASE_URL") or "https://generativelanguage.googleapis.com/v1beta"
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


def list_model_ids(client):
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
    return models

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


def extract_google_error_message(payload):
    if not isinstance(payload, dict):
        return ""
    error_obj = payload.get("error")
    if isinstance(error_obj, dict):
        return str(error_obj.get("message") or "").strip()
    if isinstance(error_obj, str):
        return error_obj.strip()
    return ""


def extract_google_text_parts(payload):
    if not isinstance(payload, dict):
        return []
    texts = []
    for candidate in payload.get("candidates", []) or []:
        content = candidate.get("content") or {}
        for part in content.get("parts", []) or []:
            text_value = part.get("text")
            if text_value:
                texts.append(text_value)
    return texts


def normalize_base64_payload(image_data):
    if not isinstance(image_data, str):
        return ""
    value = image_data.strip()
    marker = ";base64,"
    if value.startswith("data:") and marker in value:
        return value.split(marker, 1)[1]
    return value


def extract_google_image_part(payload):
    if not isinstance(payload, dict):
        return None

    for candidate in payload.get("candidates", []) or []:
        content = candidate.get("content") or {}
        for part in content.get("parts", []) or []:
            inline = part.get("inlineData") or part.get("inline_data") or {}
            image_data = normalize_base64_payload(inline.get("data") or "")
            mime_type = inline.get("mimeType") or inline.get("mime_type") or "image/png"
            if image_data:
                return {
                    "image_base64": image_data,
                    "mime_type": mime_type,
                }

    # Fallback for Imagen-style response shape.
    for entry in payload.get("generatedImages", []) or []:
        image = entry.get("image") if isinstance(entry, dict) else None
        if not isinstance(image, dict):
            continue
        image_data = normalize_base64_payload(image.get("imageBytes") or image.get("bytesBase64Encoded") or "")
        mime_type = image.get("mimeType") or "image/png"
        if image_data:
            return {
                "image_base64": image_data,
                "mime_type": mime_type,
            }
    return None


def run_google_image_generation(prompt, requested_model="", aspect_ratio="", image_size=""):
    if not GOOGLE_API_KEY:
        raise RuntimeError("GOOGLE_API_KEY is not configured.")

    model_name = (requested_model or "").strip() or GOOGLE_IMAGE_MODEL
    ratio = (aspect_ratio or "").strip() or GOOGLE_IMAGE_ASPECT_RATIO or "1:1"
    size = (image_size or "").strip() or GOOGLE_IMAGE_SIZE
    base_url = GOOGLE_API_BASE_URL.rstrip("/")
    endpoint = f"{base_url}/models/{quote(model_name, safe='')}:generateContent"

    image_config = {"aspectRatio": ratio}
    if size:
        image_config["imageSize"] = size

    payload = {
        "contents": [
            {
                "parts": [
                    {"text": prompt}
                ]
            }
        ],
        "generationConfig": {
            "imageConfig": image_config
        }
    }
    body = json.dumps(payload).encode("utf-8")
    headers = {
        "Content-Type": "application/json",
        "x-goog-api-key": GOOGLE_API_KEY,
    }
    req = url_request.Request(endpoint, data=body, headers=headers, method="POST")

    try:
        with url_request.urlopen(req, timeout=120) as resp:
            raw_body = resp.read().decode("utf-8")
    except url_error.HTTPError as e:
        error_body = e.read().decode("utf-8", "ignore")
        message = f"Google image generation failed: HTTP {e.code}"
        try:
            parsed = json.loads(error_body)
            detailed = extract_google_error_message(parsed)
            if detailed:
                message = f"{message} - {detailed}"
        except Exception:
            if error_body:
                message = f"{message} - {error_body}"
        raise RuntimeError(message)
    except url_error.URLError as e:
        raise RuntimeError(f"Google image generation failed: {e.reason}")

    try:
        response_payload = json.loads(raw_body) if raw_body else {}
    except json.JSONDecodeError:
        raise RuntimeError("Google image generation returned invalid JSON.")

    possible_error = extract_google_error_message(response_payload)
    if possible_error:
        raise RuntimeError(f"Google image generation failed: {possible_error}")

    image_part = extract_google_image_part(response_payload)
    if not image_part:
        text_parts = extract_google_text_parts(response_payload)
        if text_parts:
            raise RuntimeError(f"Google image generation did not return image data. Model message: {' '.join(text_parts)}")
        raise RuntimeError("Google image generation did not return image data.")

    result = {
        "model": model_name,
        "image_base64": image_part["image_base64"],
        "mime_type": image_part["mime_type"],
    }
    text_parts = extract_google_text_parts(response_payload)
    if text_parts:
        result["message"] = "\n".join(text_parts)
    return result


def resolve_model_name(provider: str, requested: str | None, client=None) -> tuple[str, str | None]:
    """
    Pick a model name appropriate for the provider and fall back if the value
    looks incompatible (e.g., Ollama-style model name used with OpenAI).
    Returns (model_name, warning_note).
    """
    fallback = DEFAULT_MODELS.get(provider, "gpt-4o-mini")
    model = (requested or "").strip()
    if not model:
        if provider == "ollama" and client is not None:
            available_models = list_model_ids(client)
            if available_models:
                selected = available_models[0]
                return selected, f"No Ollama model was specified. Fell back to '{selected}'."
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
    if "model_not_found" in lowered or "does not exist" in lowered or "not_found_error" in lowered or ('model "' in lowered and 'not found' in lowered):
        return message, 400
    if "invalid argument" in lowered or "invalid value" in lowered:
        return message, 400
    if "rate limit" in lowered or "too many requests" in lowered:
        return message, 429
    if "invalid api key" in lowered or "api key not valid" in lowered or "authentication" in lowered:
        return message, 401
    if "permission denied" in lowered or "forbidden" in lowered:
        return message, 403
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
        model_name, model_warning = resolve_model_name(provider, requested_model, client)

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


@app.route('/generate-image', methods=['POST'])
def generate_image():
    try:
        data = request.get_json() or {}
        prompt = (data.get('prompt') or '').strip()
        requested_model = (data.get('model') or '').strip()
        aspect_ratio = (data.get('aspect_ratio') or '').strip()
        image_size = (data.get('image_size') or '').strip()

        if not prompt:
            return jsonify({'error': 'No prompt provided'}), 400

        result = run_google_image_generation(
            prompt=prompt,
            requested_model=requested_model,
            aspect_ratio=aspect_ratio,
            image_size=image_size,
        )

        response_body = {
            'image_base64': result['image_base64'],
            'mime_type': result['mime_type'],
            'model': result['model'],
        }
        if result.get('message'):
            response_body['model_message'] = result['message']
        return jsonify(response_body)
    except Exception as e:
        message, status = classify_api_error(e)
        print(f"Image generation error: {message}")
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
        models = list_model_ids(client)
        return jsonify({'provider': provider, 'models': models})

    except Exception as e:
        print(f"Error listing models: {str(e)}")
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5010, debug=True)
