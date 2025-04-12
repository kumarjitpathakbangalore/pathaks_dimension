import comtypes.client
import fitz  # PyMuPDF
import io
from PIL import Image
import IPython.display as display
from sentence_transformers import SentenceTransformer
import faiss
import os
import matplotlib.pyplot as plt

class PPTToFile:
    def __init__(self, input_pptx, output_pdf):
        self.input_pptx = input_pptx
        self.output_pdf = output_pdf

    def convert_pptx_to_pdf(self):
        try:
            
            # Kill existing PowerPoint processes
            os.system("taskkill /f /im POWERPNT.EXE")
            powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
            powerpoint.Visible = 1
            presentation = powerpoint.Presentations.Open(self.input_pptx, WithWindow=False)
            presentation.SaveAs(self.output_pdf, 32)
            presentation.Close()
            powerpoint.Quit()
            print(f"Saved PDF to {self.output_pdf}")
        except Exception as e:
            print(f"Conversion error: {e}")

    def display(self, pdf_path, page_number):
        try:
            doc = fitz.open(pdf_path)
            if page_number < 1 or page_number > len(doc):
                print("Invalid page number!")
                return None
            page = doc.load_page(page_number - 1)
            pix = page.get_pixmap()
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            return img
        except Exception as e:
            print(f"Error displaying PDF page: {e}")
            return None
    # def search_with_rag_pipeline(self, data, file, query):
    #     model = SentenceTransformer('all-MiniLM-L6-v2')  # Efficient and accurate

    #     # Generate embeddings
    #     keys = list(data.keys())
    #     texts = list(data.values())
    #     embeddings = model.encode(texts, convert_to_tensor=True).detach().cpu().numpy()

    #     # Create FAISS index
    #     d = embeddings.shape[1]  # Vector dimension
    #     index = faiss.IndexFlatL2(d)
    #     index.add(embeddings)

        
    #     query_embedding = model.encode([query]).astype('float32')

    #     # Search for closest match
    #     _, closest_index = index.search(query_embedding, 3)

    #     # Display results
    #     print("Top Matches:")
    #     images = []
    #     for idx in closest_index[0]:
    #         img = self.display(file, int(keys[idx]))
    #         if img:
    #             images.append(img)
    #     self.save_images(images=images)  # wrap single image in list
    #     for img in images:
    #         self.display_image(img)

    def search_with_rag_pipeline(self, all_slide_map, top_k=3, query="What am I looking for?"):
        model = SentenceTransformer('all-MiniLM-L6-v2')
        keys = list(all_slide_map.keys())  # keys are (filename, slide_number)
        texts = list(all_slide_map.values())  # summarized slide text

        embeddings = model.encode(texts, convert_to_tensor=True).detach().cpu().numpy()
        d = embeddings.shape[1]

        index = faiss.IndexFlatL2(d)
        index.add(embeddings)

        query_embedding = model.encode([query]).astype('float32')
        _, closest_indices = index.search(query_embedding, top_k)

        print("🔍 Top Matches:")
        for idx in closest_indices[0]:
            filename, slide_num = keys[idx]
            pdf_path = os.path.join("pdf_output", filename.replace(".pptx", ".pdf"))
            print(f"> {filename} - Slide {slide_num + 1}")
            self.show_pdf_page(pdf_path, slide_num + 1)

    def save_images(self, images, output_folder="output_images"):
        os.makedirs(output_folder, exist_ok=True)
        for i, img in enumerate(images):
            filename = f"image_{i+1}.png"
            filepath = os.path.join(output_folder, filename)
            img.save(filepath)
            print(f"Saved: {filepath}")
    def display_image(self, img):
        plt.imshow(img)
        plt.axis('off')
        plt.show()
            
