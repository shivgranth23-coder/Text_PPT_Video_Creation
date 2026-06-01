---
title: AI Video Generator — Google Veo
emoji: 🎬
colorFrom: violet
colorTo: pink
sdk: gradio
sdk_version: "4.44.0"
app_file: app.py
pinned: false
license: mit
---

# 🎬 AI Video Generator — Google Veo + Gradio

Generate high-quality AI videos from text prompts using **Google Veo 3** (latest) or **Veo 2**, wrapped in a clean **Gradio** UI and deployed for free on **Hugging Face Spaces**.

## 🚀 Features
- Text-to-video generation with Google Veo 3.1 / 3.0 / 2.0
- Configurable aspect ratio (16:9, 9:16, 1:1)
- Adjustable duration (4 – 8 seconds)
- Negative prompt support
- AI-powered prompt enhancement
- Download generated MP4s

## 🔑 Setup

### Local
```bash
git clone <your-repo>
cd ai-video-generator
pip install -r requirements.txt
cp .env.example .env          # then fill in your key
python app.py
```

### Hugging Face Spaces
1. Fork / push this repo to a new Space (SDK = Gradio).
2. Add your `GEMINI_API_KEY` in **Settings → Secrets**.
3. The Space will install deps and launch automatically.

## 📋 Requirements
- Python 3.10+
- A **Google AI Studio** API key with Veo access (Paid Tier required for video generation)
  → [Get your key](https://aistudio.google.com/app/apikey)

## 🛠️ Project Structure
```
ai-video-generator/
├── app.py              # Main Gradio application
├── requirements.txt    # Python dependencies
├── README.md           # This file (also the HF Space config)
├── .env.example        # Environment variable template
└── outputs/            # Generated videos saved here (local)
```
