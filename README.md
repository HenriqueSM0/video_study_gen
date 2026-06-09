# 🎬 video_study_gen
 
A CLI tool that takes a video or audio URL and uses AI to automatically generate **study quizzes** and **PowerPoint presentations** from the content.
 
## How It Works
 
1. **Download** — the video/audio is fetched from the provided URL via `yt-dlp`
2. **Transcribe** — the audio is split into 1-minute chunks and transcribed using Groq's Whisper large-v3 model
3. **Generate** — an AI agent (powered by Groq) processes the transcript to produce either:
   - A quiz in Markdown format (`quiz.md` prompt)
   - A structured `.pptx` presentation with extracted video frames as slide images
## Features
 
- Supports any URL compatible with `yt-dlp` (YouTube, and many other platforms)
- Handles long videos by chunking audio before transcription
- Generates slide decks with auto-extracted frames synced to each slide
- Outputs quizzes directly to the terminal in Markdown
- Graceful error handling for API auth failures, rate limits, and download issues
## Requirements
 
- Python 3.10+
- `ffmpeg` installed and available in your PATH
- `yt-dlp` installed and available in your PATH
- A **Groq API key**
## Installation
 
```bash
git clone https://github.com/HenriqueSM0/video_study_gen.git
cd video_study_gen
pip install -r requirements.txt
```
 
Copy the example environment file and fill in your API key:
 
```bash
cp .env_ex .env
```
 
Edit `.env` and add:
 
```
GROQ_API_KEY=your_groq_api_key_here
```
 
## Usage
 
```bash
python transcripter.py
```
 
You will be prompted for:
 
1. **URL** — the video or audio URL to process
2. **Mode** — choose what to generate:
   - `0` → Quiz (printed to terminal as Markdown)
   - `1` → Slide deck (saved as `presentation.pptx`)
## Project Structure
 
```
video_study_gen/
├── transcripter.py     # Main script
├── quiz.md             # Prompt template for quiz generation
├── slide.md            # Prompt template for slide generation
├── frames/             # Extracted video frames (auto-populated)
├── requirements.txt
├── .env_ex             # Example environment file
└── .gitignore
```
 
## Tech Stack
 
| Library | Purpose |
|---|---|
| `groq` | Audio transcription (Whisper large-v3) |
| `agno` | AI agent framework |
| `yt-dlp` | Video/audio downloading |
| `pydub` | Audio chunking |
| `ffmpeg-python` | Frame extraction |
| `python-pptx` | PowerPoint generation |
| `Pillow` | Image processing |
| `python-dotenv` | Environment variable management |
 
## License
 
This project is open source. Feel free to use and adapt it.
