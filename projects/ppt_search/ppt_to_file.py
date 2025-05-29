import comtypes.client
import fitz  # PyMuPDF for PDF manipulation
import io
from PIL import Image # Pillow for image processing
import os # For interacting with the operating system (e.g., file paths, directory creation)
import matplotlib.pyplot as plt # For displaying images
from sentence_transformers import SentenceTransformer # For generating text embeddings
import faiss # For efficient similarity search of embeddings

class PPTConverterAndSearch:
    """
    Manages the conversion of PowerPoint presentations to PDF,
    displaying PDF pages, and performing semantic search on slide content
    using a RAG pipeline.
    """
    def __init__(self):
        """
        Initializes the PPTConverterAndSearch class.
        Sets up directories for PDF and image outputs.
        """
        # Ensure output directories exist
        os.makedirs("pdf_output", exist_ok=True)
        os.makedirs("output_images", exist_ok=True)
        print("📁 'pdf_output' and 'output_images' directories ensured.")

    def convert_pptx_to_pdf(self, input_pptx_path, output_pdf_path):
        """
        Converts a PowerPoint presentation (.pptx) to a PDF file.
        It attempts to kill any running PowerPoint processes to prevent conflicts,
        unhides all slides, and saves the presentation as a PDF.

        Args:
            input_pptx_path (str): The full path to the input PowerPoint file.
            output_pdf_path (str): The desired full path for the output PDF file.
        """
        powerpoint = None # Initialize powerpoint object to None
        try:
            print(f"🔄 Attempting to convert '{input_pptx_path}' to PDF...")
            # Attempt to kill existing PowerPoint processes to ensure a clean start
            # This is a brute-force method and might close other open PPT files.
            os.system("taskkill /f /im POWERPNT.EXE 2>NUL") # 2>NUL suppresses error messages

            # Create a PowerPoint application object
            powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
            powerpoint.Visible = 1

            # Open the presentation. WithWindow=False prevents a visible window from popping up.
            presentation = powerpoint.Presentations.Open(input_pptx_path, WithWindow=False)

            # Unhide all slides to ensure they are included in the PDF export
            for i in range(1, presentation.Slides.Count + 1):
                presentation.Slides(i).SlideShowTransition.Hidden = False

            # Export the presentation to PDF format (FileFormat=32)
            presentation.SaveAs(output_pdf_path, FileFormat=32)
            
            # Close the presentation and quit the PowerPoint application
            presentation.Close()
            powerpoint.Quit()
            
            print(f"✅ PDF successfully saved with all slides: {output_pdf_path}")

        except Exception as e:
            print(f"❌ Conversion error for '{input_pptx_path}': {e}")
        finally:
            # Ensure PowerPoint application is closed even if an error occurs
            if powerpoint:
                try:
                    powerpoint.Quit()
                except Exception as e:
                    print(f"Warning: Could not gracefully quit PowerPoint application: {e}")

    def display_pdf_page(self, pdf_path, page_number):
        """
        Displays a specific page from a PDF file as an image.

        Args:
            pdf_path (str): The full path to the PDF file.
            page_number (int): The 1-based index of the page to display.

        Returns:
            PIL.Image.Image or None: The PIL Image object of the page if successful,
                                     otherwise None.
        """
        try:
            # Open the PDF document
            doc = fitz.open(pdf_path)
            
            # Validate the page number
            if not (1 <= page_number <= len(doc)):
                print(f"⚠️ Invalid page number: {page_number}. PDF has {len(doc)} pages.")
                return None
            
            # Load the specified page (fitz uses 0-based indexing)
            page = doc.load_page(page_number - 1)
            
            # Get a pixmap (pixel map) of the page
            pix = page.get_pixmap()
            
            # Convert the pixmap to a PIL Image object
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            
            # Display the image using matplotlib
            self._display_image(img)
            return img # Return the image object if needed elsewhere
        except fitz.FileDataError:
            print(f"❌ Error: The file '{pdf_path}' is not a valid PDF or is corrupted.")
            return None
        except Exception as e:
            print(f"❌ Error displaying PDF page {page_number} from '{pdf_path}': {e}")
            return None

    def search_with_rag_pipeline(self, all_slide_map, query="What am I looking for?", top_k=3):
        """
        Performs a semantic search across summarized slide content using SentenceTransformers
        and FAISS. It finds the top_k most relevant slides based on the query.

        Args:
            all_slide_map (dict): A dictionary where keys are (filename, slide_number) tuples
                                  and values are the summarized text content of each slide.
            query (str): The search query provided by the user.
            top_k (int): The number of top matching slides to retrieve and display.
        """
        print(f"\n🔍 Performing semantic search for query: '{query}'")
        
        # Load a pre-trained SentenceTransformer model for generating embeddings
        # 'all-MiniLM-L6-v2' is a good balance of efficiency and accuracy.
        model = SentenceTransformer('all-MiniLM-L6-v2')

        # Prepare keys and texts for embedding and indexing
        # Keys are tuples: (original_pptx_filename, 0-based_slide_index)
        keys = list(all_slide_map.keys())
        # Texts are the summarized content of each slide
        texts = list(all_slide_map.values())

        if not texts:
            print("⚠️ No slide content available for search. Please ensure `all_slide_map` is populated.")
            return

        # Generate embeddings for all slide texts
        # convert_to_tensor=True for GPU acceleration if available, then move to CPU and convert to NumPy
        embeddings = model.encode(texts, convert_to_tensor=True).detach().cpu().numpy()

        # Get the dimensionality of the embeddings (vector dimension)
        d = embeddings.shape[1]

        # Create a FAISS index for efficient similarity search
        # IndexFlatL2 uses Euclidean distance (L2 norm) for similarity.
        index = faiss.IndexFlatL2(d)
        
        # Add the slide embeddings to the FAISS index
        index.add(embeddings)

        # Encode the query into an embedding
        query_embedding = model.encode([query]).astype('float32')

        # Search the FAISS index for the top_k closest matches to the query embedding
        # D is distances, I is indices
        _, closest_indices = index.search(query_embedding, top_k)

        print("\n✨ Top Matches Found:")
        # Iterate through the indices of the closest matches
        for idx in closest_indices[0]:
            # Retrieve the original filename and 0-based slide number using the index
            original_pptx_filename, slide_num_zero_based = keys[idx]
            
            # Construct the PDF path for the matched slide's presentation
            # Assuming PDFs are named the same as PPTXs but with .pdf extension,
            # and stored in 'pdf_output'
            pdf_path = os.path.join("pdf_output", original_pptx_filename.replace(".pptx", ".pdf"))
            
            # Print the matching slide information
            print(f"> Found in: '{original_pptx_filename}' - Slide {slide_num_zero_based + 1}")
            
            # Display the actual PDF page
            self.display_pdf_page(pdf_path, slide_num_zero_based + 1)
            print("-" * 30) # Separator for readability

    def save_images(self, images, output_folder="output_images"):
        """
        Saves a list of PIL Image objects to the specified output folder.

        Args:
            images (list): A list of PIL.Image.Image objects to save.
            output_folder (str): The directory where images will be saved.
        """
        os.makedirs(output_folder, exist_ok=True) # Ensure the output directory exists
        print(f"💾 Saving images to: {output_folder}")
        for i, img in enumerate(images):
            filename = f"image_{i+1}.png"
            filepath = os.path.join(output_folder, filename)
            try:
                img.save(filepath)
                print(f"Saved: {filepath}")
            except Exception as e:
                print(f"❌ Error saving image {i+1} to '{filepath}': {e}")

    def _display_image(self, img):
        """
        Internal helper method to display a PIL Image using Matplotlib.
        This is a private method as it's an internal utility for the class.

        Args:
            img (PIL.Image.Image): The image to display.
        """
        plt.figure(figsize=(10, 7)) # Optional: Set figure size for better display
        plt.imshow(img)
        plt.axis('off') # Hide axes for cleaner image display
        plt.show()