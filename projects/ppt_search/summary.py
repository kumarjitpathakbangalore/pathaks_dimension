import os
from pptx import Presentation
import openai
import matplotlib.pyplot as plt
from tqdm import tqdm



class SlideSummarizer:
    def __init__(self, folder_path, model="gpt-4", provider="openai", save_path=None):
        self.folder_path = folder_path
        self.model = model
        self.provider = provider.lower()
        self.save_path = save_path

        if self.provider == "openai":
            import openai
            openai.api_key = os.environ.get("OPENAI_API_KEY")
            self.client = openai
        elif self.provider == "groq":
            from groq import Groq
            self.client = Groq(api_key=os.environ.get("GROQ_API_KEY"))
        elif self.provider == "gemini":
            import google.generativeai as genai
            genai.configure(api_key=os.environ.get("GOOGLE_API_KEY"))
            self.client = genai
        else:
            raise ValueError("Unsupported provider. Choose from: 'openai', 'groq', 'gemini'")


    def extract_pptx_text(self):
        extracted = {}
        for file in os.listdir(self.folder_path):
            if file.endswith(".pptx"):
                prs = Presentation(os.path.join(self.folder_path, file))
                slides_text = []
                for slide in prs.slides:
                    text = " ".join(
                        shape.text.strip()
                        for shape in slide.shapes
                        if hasattr(shape, "text") and shape.text.strip()
                    )
                    slides_text.append(text)
                extracted[file] = slides_text
        return extracted

    def summarize_slide(self, slide_text):
        if not slide_text.strip():
            return "[No content]"

        if len(slide_text) > 3000:
            slide_text = slide_text[:3000] + "..."

        prompt = f"Summarize the following slide content in 2 lines, do not repeat the phrase 2-line summary in your response. also always give an output to the point and if you feel there is no content in the slide try to understand why the slide was put here. put special reference to what ever questions that may be asked by the user:\n\n{slide_text}"

        try:
            if self.provider == "openai":
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.3,
                )
                return response.choices[0].message.content

            elif self.provider == "groq":
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.3,
                )
                return response.choices[0].message.content

            elif self.provider == "gemini":
                model = self.client.GenerativeModel(self.model)
                response = model.generate_content(prompt)
                return response.text

        except Exception as e:
            return f"[Error summarizing slide: {e}]"


    def summarize_all(self):
        extracted_texts = self.extract_pptx_text()
        summaries = {}

        for filename, slides in extracted_texts.items():
            print(f"Summarizing {filename} with {len(slides)} slides...")
            summaries[filename] = [
                self.summarize_slide(slide) for slide in tqdm(slides, desc=filename)
            ]

        if self.save_path:
            self.save_summaries(summaries)

        return summaries


    def display_image(self, img):
        plt.imshow(img)
        plt.axis('off')
        plt.show()
