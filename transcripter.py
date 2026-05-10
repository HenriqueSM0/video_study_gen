import os, json
from agno.agent import Agent
from agno.models.groq import Groq as Agno_Groq
from dotenv import load_dotenv
from groq import Groq
from pydub import AudioSegment
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import subprocess

load_dotenv()
client = Groq()

class Audio_transcription:
    output = "video.mp4"

    def download_from_url(self, url):
        try: subprocess.run(["yt-dlp", url, "-o", self.output], check=True)
        except subprocess.CalledProcessError as e:
            raise RuntimeError(f"Download failed (exit code {e.returncode}). Check the URL and try again.")

    def split_audio(self):
        chunk_duration_ms = 60000
        webm_path = self.output + ".webm"
        if not os.path.exists(webm_path):
            raise RuntimeError(f"Audio file not found: {webm_path}")
        audio = AudioSegment.from_file(webm_path)
        chunks = []
        for i in range(0, len(audio), chunk_duration_ms):
            chunk = audio[i:i + chunk_duration_ms]
            chunk_path = f"/tmp/chunk_{i}.mp3"
            chunk.export(chunk_path, format="mp3", bitrate="64k")
            chunks.append(chunk_path)
        return chunks

    def transcript_chunks(self, chunks):
        full_transcript = ""
        for chunk_path in chunks:
            try:
                with open(chunk_path, "rb") as f:
                    result = client.audio.transcriptions.create(
                        model="whisper-large-v3",
                        file=f,
                    )
                full_transcript += result.text + " "
                os.remove(chunk_path)
            except Groq.AuthenticationError:
                raise RuntimeError("Invalid Groq API key — check your credentials.")
            except Groq.RateLimitError:
                raise RuntimeError("Groq rate limit or token quota exceeded.")
            except Groq.APIConnectionError as e:
                raise RuntimeError(f"Could not reach Groq API: {e}")
            except Groq.GroqError as e:
                raise RuntimeError(f"Groq API error on chunk {chunk_path}: {e}")
        return full_transcript.strip()

    def transcribe_large_audio(self, url):
        self.download_from_url(url)  
        chunks = self.split_audio()  
        return self.transcript_chunks(chunks)


class Agent_tools:
    agent = Agent(model=Agno_Groq(id="openai/gpt-oss-120b"))
    quiz_prompt = "quiz.md"
    slide_prompt = "slide.md"

    def file_to_prompt(self, filename):
        return open(filename, "r", encoding="utf-8").read()

    def build_pptx(self, slides):
        filename = "presentation.pptx"
        prs = Presentation()
        layout = prs.slide_layouts[1]
        for s in slides:
            slide = prs.slides.add_slide(layout)
            slide.shapes.title.text = s["title"]
            slide.shapes.title.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(0x1E, 0x27, 0x61)
            tf = slide.placeholders[1].text_frame
            tf.clear()
            for bullet in s["bullets"]:
                p = tf.add_paragraph()
                p.text = bullet
                p.level = 0
            slide.notes_slide.notes_text_frame.text = s["notes"]
        prs.save(filename)

    def gen_questions(self, url):
        prompt = self.file_to_prompt(self.quiz_prompt)
        transcript = Audio_transcription().transcribe_large_audio(url)
        self.agent.print_response(prompt + "\n" + transcript)

    def gen_slides(self, url):
        prompt = self.file_to_prompt(self.slide_prompt)
        transcript = Audio_transcription().transcribe_large_audio(url)
        raw = self.agent.run(prompt + "\n" + transcript).content
        self.build_pptx(json.loads(raw.strip().strip("```json").strip("```")))
        print("presentation.pptx saved!")

print("---------WELCOME!----------")
print("URL: ")
url = input()
print("Digit:")
print("0 - Generate video quiz")
print("1 - Generate slide")
mode = input()

try:
    AT = Agent_tools()
    if mode == "0":
        AT.gen_questions(url)
    elif mode == "1":
        AT.gen_slides(url)
except RuntimeError as e:
    print(f"\n❌ Error: {e}")