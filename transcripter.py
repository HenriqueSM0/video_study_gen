import os, json
from agno.agent import Agent
from agno.models.groq import Groq as Agno_Groq
from dotenv import load_dotenv
from groq import Groq
from pydub import AudioSegment
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE
from PIL import Image as PILImage
import ffmpeg
import subprocess

load_dotenv()
client = Groq()

class Video_decompose:
    output = "video.mp4"
    frames_dir = 'frames'

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
    
    def extract_frames(self, n_frames) :
        webm_path = self.output + ".webm"
        audio = AudioSegment.from_file(webm_path)
        print(n_frames)
        fps = 1000 * n_frames / len(audio)
        print(fps)
        try :
            (
                ffmpeg
                .input(webm_path)
                .filter("fps", fps)
                .output(f"{self.frames_dir}/frame_%04d.jpg")
                .run(quiet=True)
            )
        except ffmpeg.Error as e:
            print(e.stderr.decode())

    def transcribe_large_audio(self, url):
        self.download_from_url(url)  
        chunks = self.split_audio()  
        return self.transcript_chunks(chunks)

class Agent_tools:
    agent = Agent(model=Agno_Groq(id="openai/gpt-oss-120b"))
    quiz_prompt = "quiz.md"
    slide_prompt = "slide.md"
    frames_folder = Video_decompose.frames_dir

    def file_to_prompt(self, filename):
        return open(filename, "r", encoding="utf-8").read()
    
    def build_pptx(self, slides):
        filename = "presentation.pptx"
        prs = Presentation()
        
        slide_width = prs.slide_width
        slide_height = prs.slide_height
        half_width = slide_width // 2

        # Use a blank layout to avoid default placeholders
        blank_layout = prs.slide_layouts[6]  # index 6 = blank

        for i in range(len(slides)):
            slide = prs.slides.add_slide(blank_layout)
            bg = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,  
                left=Emu(0),
                top=Emu(0),
                width=half_width,
                height=slide_height,
            )
            bg.fill.solid()
            bg.fill.fore_color.rgb = RGBColor(0x00, 0x00, 0x00)
            bg.line.fill.background()  # no border

            # Title text box (white)
            title_top = slide_height * 15 // 100        # 15% from top
            title_height = slide_height * 15 // 100
            title_box = slide.shapes.add_textbox(
                left=Emu(360000),                        # ~0.4" left margin
                top=title_top,
                width=half_width - Emu(720000),          # 0.4" margins each side
                height=title_height,
            )
            tf = title_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = slides[i]["title"]
            run = p.runs[0]
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.font.bold = True
            run.font.size = Pt(28)

            # Bullets text box (white)
            bullets_top = title_top + title_height + Emu(200000)
            bullets_box = slide.shapes.add_textbox(
                left=Emu(360000),
                top=bullets_top,
                width=half_width - Emu(720000),
                height=slide_height - bullets_top - Emu(360000),
            )
            btf = bullets_box.text_frame
            btf.word_wrap = True
            btf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

            for j, bullet in enumerate(slides[i]["bullets"]):
                p = btf.paragraphs[0] if j == 0 else btf.add_paragraph()
                p.text = bullet
                p.level = 0
                p.space_after = Pt(6)
                run = p.runs[0]
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.size = Pt(16)

            # ── RIGHT HALF: image at original aspect ratio, centered ─────
            image_files = sorted(os.listdir(self.frames_folder))
            image_path = os.path.join(self.frames_folder, image_files[i])

            # Measure original image size to preserve aspect ratio
            from PIL import Image as PILImage
            with PILImage.open(image_path) as img:
                img_w, img_h = img.size

            # Scale to fit within the right half, preserving aspect ratio
            available_w = half_width
            available_h = slide_height
            scale = min(available_w / img_w, available_h / img_h)
            pic_w = int(img_w * scale)
            pic_h = int(img_h * scale)

            # Center within the right half
            pic_left = half_width + (available_w - pic_w) // 2
            pic_top = (available_h - pic_h) // 2

            slide.shapes.add_picture(
                image_path,
                left=pic_left,
                top=pic_top,
                width=pic_w,
                height=pic_h,
            )

            # Notes
            slide.notes_slide.notes_text_frame.text = slides[i]["notes"]
        prs.save(filename)

    def gen_questions(self, url):
        prompt = self.file_to_prompt(self.quiz_prompt)
        transcript = Video_decompose().transcribe_large_audio(url)
        self.agent.print_response(prompt + "\n" + transcript)

    def gen_slides(self, url):
        prompt = self.file_to_prompt(self.slide_prompt)
        transcript = Video_decompose().transcribe_large_audio(url)
        raw = self.agent.run(prompt + "\n" + transcript).content
        slides = json.loads(raw.strip().strip("```json").strip("```"))
        Video_decompose().extract_frames(len(slides))
        self.build_pptx(slides)
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