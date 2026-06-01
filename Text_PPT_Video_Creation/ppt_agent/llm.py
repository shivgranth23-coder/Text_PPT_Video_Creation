"""
llm.py - LLM content generator for slide outlines (Gemini, NVIDIA Gemma, or DeepSeek).
"""

import json
import re
import time
import gradio as gr
import requests
from google import genai
from google.genai import types
from google.genai import errors

from .normalize import normalize_slides

SYSTEM_PROMPT = (
    "You are an expert PowerPoint presentation designer and educator. "
    "You produce highly engaging, professional, and explainable slide content. "
    "Your content should be detailed and provide deep insights, not just surface-level facts. "
    "Always respond with valid JSON only — no markdown fences, no extra text."
)

LAYOUT_DESCRIPTIONS = {
    "title_slide":     "Opening slide — big title + subtitle",
    "section_header":  "Section divider — section title + brief teaser line",
    "bullet_points":   "Title + 4-6 concise, insightful bullet points",
    "two_column":      "Two equal text columns side by side",
    "quote_slide":     "Large impactful pull-quote with attribution",
    "data_summary":    "Title + 3 key statistics or KPIs with labels",
    "closing_slide":   "Thank-you / call-to-action closing slide",
}


def get_gemini_client(api_key: str) -> genai.Client:
    return genai.Client(api_key=api_key)


def get_client(api_key: str, provider: str = "gemini") -> genai.Client:
    """Legacy wrapper for backward compatibility."""
    if provider == "gemini":
        return get_gemini_client(api_key)
    # For NVIDIA, we don't need a pre-initialized client
    return None


def _generate_with_gemini(client, prompt: str) -> str:
    """Helper to generate content with Gemini using standard models list."""
    models_to_try = [
        "gemini-2.5-flash",
        "gemini-2.0-flash",
        "gemini-2.0-flash-lite",
        "gemini-flash-latest",
        "gemini-flash-lite-latest",
        "gemini-2.5-pro",
    ]
    last_error = None
    for model_name in models_to_try:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config=types.GenerateContentConfig(
                    system_instruction=SYSTEM_PROMPT,
                    response_mime_type="application/json",
                    temperature=0.75,
                    max_output_tokens=8192,
                ),
            )
            return response.text.strip()
        except Exception as e:
            last_error = e
            continue
    raise last_error or Exception("Failed to generate content with Gemini")


def _generate_with_nvidia(prompt: str, api_key: str, log_cb=None) -> str:
    """
    Generate content using NVIDIA API (Gemma 4 31B).
    Returns JSON text response.
    """
    invoke_url = "https://integrate.api.nvidia.com/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Accept": "application/json"
    }
    
    payload = {
        "model": "google/gemma-4-31b-it",
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": 4096,
        "temperature": 0.75,
        "top_p": 0.95,
        "stream": False,
    }
    
    response = requests.post(invoke_url, headers=headers, json=payload, timeout=120)
    response.raise_for_status()
    
    data = response.json()
    if "choices" in data and len(data["choices"]) > 0:
        return data["choices"][0]["message"]["content"]
    raise ValueError("Unexpected NVIDIA API response format")


def _generate_with_deepseek(prompt: str, api_key: str, log_cb=None) -> str:
    """
    Generate content using DeepSeek v4 Pro via NVIDIA API (OpenAI-compatible).
    Returns JSON text response.
    """
    from openai import OpenAI
    
    try:
        if log_cb:
            log_cb("Connecting to DeepSeek API...")
        
        client = OpenAI(
            base_url="https://integrate.api.nvidia.com/v1",
            api_key=api_key,
            timeout=30.0,  # 30 second timeout per request
        )
        
        if log_cb:
            log_cb("Sending request to DeepSeek v4 Pro...")
        
        # Collect streaming response with timeout protection
        response_text = ""
        chunk_count = 0
        start_time = time.time()
        max_wait = 120  # 2 minute absolute max
        
        completion = client.chat.completions.create(
            model="deepseek-ai/deepseek-v4-pro",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.75,
            top_p=0.95,
            max_tokens=4096,
            stream=True,
        )
        
        for chunk in completion:
            # Check timeout
            elapsed = time.time() - start_time
            if elapsed > max_wait:
                raise TimeoutError(f"DeepSeek response took too long ({elapsed:.0f}s)")
            
            if not getattr(chunk, "choices", None):
                continue
            if chunk.choices and chunk.choices[0].delta.content is not None:
                response_text += chunk.choices[0].delta.content
                chunk_count += 1
                if chunk_count % 20 == 0:  # Log every 20 chunks
                    if log_cb:
                        log_cb(f"Receiving... ({len(response_text)} chars)")
        
        if not response_text.strip():
            raise ValueError("Empty response from DeepSeek API")
        
        if log_cb:
            log_cb(f"Received {len(response_text)} characters")
        
        return response_text
        
    except Exception as e:
        error_msg = str(e)
        if log_cb:
            log_cb(f"DeepSeek error: {error_msg}")
        raise Exception(f"DeepSeek API error: {error_msg}")



def _generate_with_ollama(prompt: str, model: str, log_cb=None) -> str:
    """
    Generate content using local Ollama API.
    """
    if log_cb:
        log_cb(f"Connecting to Ollama ({model})...")
    
    url = "http://localhost:11434/api/generate"
    payload = {
        "model": model,
        "prompt": f"{SYSTEM_PROMPT}\n\n{prompt}",
        "stream": False,
        "format": "json"
    }
    
    try:
        response = requests.post(url, json=payload, timeout=180)
        response.raise_for_status()
        return response.json().get("response", "")
    except Exception as e:
        if log_cb:
            log_cb(f"Ollama error: {e}")
        raise Exception(f"Ollama failed: {e}")


def generate_outline(
    client,
    topic: str,
    n_slides: int,
    allowed_layouts: list[str],
    theme_name: str,
    provider: str = "gemini",
    ollama_model: str = "gemma4:31b-cloud",
    log_cb=None,
) -> list[dict]:
    """
    Ask LLM (Gemini, NVIDIA Gemma, or DeepSeek) to produce a full slide-deck outline as a JSON array.
    Returns a list of slide dicts.
    """
    available = {k: LAYOUT_DESCRIPTIONS[k] for k in allowed_layouts if k in LAYOUT_DESCRIPTIONS}

    prompt = f"""Create a professional PowerPoint presentation on: "{topic}"

Rules:
- Exactly {n_slides} slides total.
- First slide MUST be "title_slide". Last slide MUST be "closing_slide".
- Use ONLY these layouts: {json.dumps(list(available.keys()))}
- Vary layouts — do not repeat the same layout consecutively more than once.
- Content must be highly educational, factual, deeply engaging, and self-explanatory.
- Provide comprehensive details in bullets and descriptions.
- Visual theme context: {theme_name}

Return a JSON array. Each element must follow this exact schema:
{{
  "slide_number": <int>,
  "layout": "<layout_name>",
  "title": "<slide title>",
  "subtitle": "<subtitle text or null>",
  "content": {{ ... }},
  "speaker_notes": "<2-4 sentences of narration in natural, conversational speech>"
}}

CRITICAL — populate the correct content fields for EACH layout (never leave body fields empty):
- title_slide: set "title" and "subtitle". content may be {{}}.
- section_header: set "title" and "subtitle" (teaser line, not null).
- bullet_points: set "title" and content.bullets with 4-6 non-empty strings. Do NOT use left_column/right_column/stats here.
- two_column: set "title", content.left_column (3-5 sentences), content.right_column (3-5 sentences). Do NOT put column text only in bullets.
- quote_slide: set content.quote (full quote) and content.attribution.
- data_summary: set "title" and content.stats with EXACTLY 3 objects, each with non-empty label, value, and description.
- closing_slide: set "title" and content.cta (call-to-action). subtitle optional.

Use exact layout names: title_slide, section_header, bullet_points, two_column, quote_slide, data_summary, closing_slide.

Return ONLY the JSON array. No markdown. No extra text."""

    # Provider: Ollama
    if provider == "ollama":
        if not ollama_model:
            raise gr.Error("Ollama model name required")
        max_retries = 2
        for attempt in range(max_retries):
            try:
                if log_cb:
                    log_cb(f"Attempting Ollama {ollama_model} (Attempt {attempt+1})")
                text = _generate_with_ollama(prompt, ollama_model, log_cb=log_cb)
                text = text.strip()
                text = re.sub(r"^```(?:json)?\s*", "", text)
                text = re.sub(r"\s*```$", "", text)
                text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
                slides = json.loads(text)
                if not isinstance(slides, list):
                    raise ValueError("Expected a JSON array of slides")
                slides = normalize_slides(slides)
                if log_cb:
                    log_cb(f"Outline normalized — {len(slides)} slides with layout-specific content")
                return slides
            except Exception as e:
                if attempt < max_retries - 1:
                    time.sleep(5)
                else:
                    raise gr.Error(f"Ollama error: {e}")

    # Provider: Gemini (most reliable)
    if provider == "gemini":
        max_retries = 3
        for attempt in range(max_retries):
            try:
                if log_cb:
                    log_cb(f"Attempting Gemini (Attempt {attempt+1})")
                text = _generate_with_gemini(client, prompt)
                text = text.strip()
                text = re.sub(r"^```(?:json)?\s*", "", text)
                text = re.sub(r"\s*```$", "", text)
                text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
                slides = json.loads(text)
                if not isinstance(slides, list):
                    raise ValueError("Expected a JSON array of slides")
                slides = normalize_slides(slides)
                if log_cb:
                    log_cb(f"Outline normalized — {len(slides)} slides with layout-specific content")
                return slides
            except Exception as e:
                if attempt < max_retries - 1:
                    time.sleep(5)
                else:
                    raise gr.Error(f"Gemini error: {e}")
    
    # Provider: NVIDIA Gemma-4
    if provider == "nvidia":
        if not nvidia_api_key:
            raise gr.Error("NVIDIA API key required")
        max_retries = 3
        for attempt in range(max_retries):
            try:
                if log_cb:
                    log_cb(f"Attempting NVIDIA Gemma-4 (Attempt {attempt+1})")
                print(f"Attempting NVIDIA Gemma-4 (Attempt {attempt+1})")
                text = _generate_with_nvidia(prompt, nvidia_api_key, log_cb=log_cb)
                text = text.strip()
                text = re.sub(r"^```(?:json)?\s*", "", text)
                text = re.sub(r"\s*```$", "", text)
                text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
                slides = json.loads(text)
                if not isinstance(slides, list):
                    raise ValueError("Expected a JSON array of slides")
                slides = normalize_slides(slides)
                if log_cb:
                    log_cb(f"Outline normalized — {len(slides)} slides with layout-specific content")
                return slides
            except Exception as e:
                err_str = str(e)
                is_rate_limit = "429" in err_str or "rate" in err_str.lower()
                if is_rate_limit and attempt < max_retries - 1:
                    wait_sec = 60 + (attempt * 30)
                    msg = f"Rate limited on NVIDIA. Waiting {wait_sec}s before retry..."
                    if log_cb: log_cb(msg)
                    print(msg)
                    time.sleep(wait_sec)
                else:
                    raise gr.Error(f"NVIDIA API failed: {e}")
        raise gr.Error(f"NVIDIA API exhausted after {max_retries} retries")
    
    # Provider: Gemini (default)
    # Ordered from latest/fastest to most capable fallback
    models_to_try = [
        "gemini-2.5-flash",
        "gemini-2.0-flash",
        "gemini-2.0-flash-lite",
        "gemini-flash-latest",
        "gemini-flash-lite-latest",
        "gemini-2.5-pro",
    ]

    last_error = None
    for model_name in models_to_try:
        max_retries = 3
        for attempt in range(max_retries):
            try:
                if log_cb:
                    log_cb(f"Attempting with model: {model_name} (Attempt {attempt+1})")
                print(f"Attempting with model: {model_name} (Attempt {attempt+1})")
                response = client.models.generate_content(
                    model=model_name,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        system_instruction=SYSTEM_PROMPT,
                        response_mime_type="application/json",
                        temperature=0.75,
                        max_output_tokens=8192,
                    ),
                )
                # If successful, return immediately
                text = response.text.strip()
                text = re.sub(r"^```(?:json)?\s*", "", text)
                text = re.sub(r"\s*```$", "", text)
                # Strip invisible control characters that corrupt PPTX XML
                text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
                slides = json.loads(text)
                if not isinstance(slides, list):
                    raise ValueError("Expected a JSON array of slides")
                slides = normalize_slides(slides)
                if log_cb:
                    log_cb(f"Outline normalized — {len(slides)} slides with layout-specific content")
                return slides

            except (errors.ClientError, errors.ServerError, Exception) as e:
                last_error = e
                err_str = str(e)
                is_rate_limit = "429" in err_str
                is_server_error = "503" in err_str or "UNAVAILABLE" in err_str or isinstance(e, errors.ServerError)

                if is_rate_limit or is_server_error:
                    if attempt < max_retries - 1:
                        # Longer wait for 503 server overload vs 429 rate limit
                        wait_sec = (60 if is_server_error else 35) + (attempt * 30)
                        reason = "Server overloaded" if is_server_error else "Rate limited"
                        msg = f"{reason} on {model_name}. Waiting {wait_sec}s before retry..."
                        if log_cb: log_cb(msg)
                        print(msg)
                        time.sleep(wait_sec)
                    else:
                        msg = f"Quota/capacity exhausted for {model_name}, trying next model..."
                        if log_cb: log_cb(msg)
                        print(msg)
                        break  # Try next model
                else:
                    # Non-retryable error (e.g. 404 model not found, 400 bad schema, etc.)
                    msg = f"Model {model_name} failed with non-retryable error: {err_str}. Trying next model..."
                    if log_cb: log_cb(msg)
                    print(msg)
                    break  # Break out of attempts for this model, try next model

    # If we get here, all models failed
    raise gr.Error(f"All Gemini models exhausted or failed. Last error: {last_error}")
