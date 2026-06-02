# Text PPT Video Creation

This folder contains the public-ready source code for the AI-powered PPT generation app.
It includes only the scripts needed to build the presentation generator, not any generated output files.

## What is included
- `ppt_agent/` source files for the PPT agent
- `requirements.txt` with the Python dependencies
- `README.md` with usage guidance
- `.gitignore` to exclude private files and generated outputs
- `.env.example` with placeholder environment variable names

## Important
This folder does not contain any real API keys or private credentials.

## About the Text → PPT → Video Agent
This app converts a simple text topic into a complete presentation and optionally into a video.

The process includes:
- **AI slide generation**: The agent uses an LLM to create titles, bullet points, section headers, quotes, and closing slides from your prompt.
- **Slide formatting**: The app converts the AI output into a polished PowerPoint presentation with consistent layouts and theme styling.
- **Narration audio**: Optionally, speaker notes are turned into voice narration and embedded into each slide.
- **Video export**: If Microsoft PowerPoint is installed on Windows, the app can export the final PPTX as an MP4 video with synced audio.

This makes it easy to go from a single idea to a fully designed presentation and shareable video.

## Setup
1. Copy `.env.example` to `.env`.
2. Add your own keys to `.env`.
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Run the app:
   ```bash
   python -m ppt_agent.agent
   ```

## If the port is busy
- The app tries to use port `7861` by default.
- If that port is already in use, it will automatically try the next available port.
- To force a port, set:
  ```bash
  set GRADIO_SERVER_PORT=7870
  python -m ppt_agent.agent
  ```

## Required environment variables
- `GEMINI_API_KEY`
- `NVIDIA_NIM_API_KEY`

## Notes
- Do not commit `.env` to GitHub.
- Keep your real API keys private.
- This folder is safe to share publicly because it does not include any `.env` or generated output files.
