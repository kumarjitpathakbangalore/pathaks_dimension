# PPT Slide Summarizer and Semantic Search

This project provides a comprehensive pipeline to extract text from PowerPoint presentations, summarize each slide using a Large Language Model (LLM), convert PPTX files to searchable PDFs, and then enable semantic search across all slide summaries. This allows users to quickly find relevant slides across multiple presentations using natural language queries.

## ✨ Features

* **PPTX to PDF Conversion:** Converts all `.pptx` files in a specified input folder to `.pdf` format.
* **Text Extraction:** Extracts all textual content from each slide of your PowerPoint presentations.
* **LLM-powered Summarization:** Utilizes OpenAI, Groq, or Gemini models to generate concise, 2-line summaries for every slide.
* **Semantic Search (RAG Pipeline):** Employs Sentence Transformers for embeddings and FAISS for efficient similarity search, allowing you to find slides relevant to your query.
* **Interactive Display:** Shows the actual PDF page of the top matching slides directly.
* **Robust Error Handling:** Includes checks for missing files, invalid API keys, and common conversion issues.

## 🚀 Getting Started

Follow these steps to set up and run the project.

### 1. Prerequisites

Before you begin, ensure you have the following installed on your system:

* **Python 3.8+**
* **Microsoft PowerPoint (Windows Only):** The PPTX to PDF conversion relies on `comtypes.client`, which requires a working installation of Microsoft PowerPoint on a Windows operating system.
* **API Keys for LLM Providers:** You will need API keys for at least one of the supported LLM providers (OpenAI, Groq, or Google Gemini).

### 2. Set up your Project Environment

1.  **Clone the Repository (if applicable) or create your project folder:**
    ```bash
    mkdir ppt_rag_search
    cd ppt_rag_search
    ```

2.  **Create a Virtual Environment (Recommended):**
    This helps manage project dependencies in isolation.
    ```bash
    python -m venv venv
    ```

3.  **Activate the Virtual Environment:**
    * **Windows:**
        ```bash
        .\venv\Scripts\activate
        ```
    * **macOS / Linux:**
        ```bash
        source venv/bin/activate
        ```
    (You'll see `(venv)` prefix in your terminal after activation)

4.  **Install Required Python Libraries:**
    Create a `requirements.txt` file in your project root with the following content:

    ```
    python-pptx
    openai
    groq
    google-generativeai
    comtypes
    PyMuPDF # fitz
    Pillow # PIL
    matplotlib
    sentence-transformers
    faiss-cpu # or faiss-gpu if you have a compatible GPU
    tqdm
    ```

    Then, install them using pip:
    ```bash
    pip install -r requirements.txt
    ```
    * **Note on `faiss`:** If you have an NVIDIA GPU and CUDA installed, you can use `faiss-gpu` instead of `faiss-cpu` for faster performance (`pip install faiss-gpu`). Otherwise, `faiss-cpu` is sufficient.

### 3. Configure API Keys

The project uses environment variables to securely access your LLM API keys.

* **For OpenAI:**
    ```bash
    export OPENAI_API_KEY='your_openai_api_key_here'
    ```
* **For Groq:**
    ```bash
    export GROQ_API_KEY='your_groq_api_key_here'
    ```
* **For Google Gemini:**
    ```bash
    export GOOGLE_API_KEY='your_gemini_api_key_here'
    ```
    **Windows (Command Prompt):**
    ```cmd
    set OPENAI_API_KEY=your_openai_api_key_here
    rem or
    set GROQ_API_KEY=your_groq_api_key_here
    rem or
    set GOOGLE_API_KEY=your_gemini_api_key_here
    ```
    **Windows (PowerShell):**
    ```powershell
    $env:OPENAI_API_KEY='your_openai_api_key_here'
    # or
    $env:GROQ_API_KEY='your_groq_api_key_here'
    # or
    $env:GOOGLE_API_KEY='your_gemini_api_key_here'
    ```
    * **Important:** Replace `'your_api_key_here'` with your actual API key. These keys are sensitive and should not be committed to version control.

### 4. Project Structure and Data Preparation

Organize your project files as follows:

your_project_root/
├── main.py
├── summary.py            # Contains SlideSummarizer class
├── read_json.py          # Contains JSONManager class (previously ReadJSON)
├── ppt_to_file.py        # Contains PPTConverterAndSearch class (previously PPTToFile)
├── requirements.txt
├── pptbase/              # Create this folder
│   └── your_presentation_1.pptx
│   └── your_presentation_2.pptx
│   └── ...
├── pdf_output/           # Will be created by the script (for converted PDFs)
├── text_output/          # Will be created by the script (for JSON summaries)
├── output_images/        # Will be created by the script (for displayed slide images)
└── venv/                 # Your virtual environment


* **`pptbase/`**: Place all the PowerPoint (`.pptx`) files you want to process in this directory.

## 5. Run the Code

Once you have activated your virtual environment, installed dependencies, configured API keys, and placed your `.pptx` files in the `pptbase/` folder, you can run the main script.

```bash
python main.py
What the Script Does:
Converts PPTX to PDF: Iterates through pptbase/, converts each .pptx to .pdf and saves them in pdf_output/.
Summarizes Slides: Extracts text from each slide of the .pptx files, sends the text to your chosen LLM (OpenAI, Groq, or Gemini) for a 2-line summary, and saves all summaries to text_output/slide_summaries.json.
Prepares for Search: Loads the summaries into a format suitable for FAISS.
Prompts for Query: Asks you to "Enter your query to find relevant slides:".
Performs Search: Uses semantic search to find the top 3 most relevant slides based on your query.
Displays Results: For each matching slide, it will display the actual PDF page using matplotlib.
troubleshooting
Conversion error: ... Hiding the application window is not allowed.: This means your PowerPoint installation doesn't allow invisible automation. Fix: In ppt_to_file.py, change powerpoint.Visible = 0 to powerpoint.Visible = 1 within the convert_pptx_to_pdf method. You will see PowerPoint windows pop up during conversion.
Conversion error: ... The system cannot find the path specified.: This indicates PowerPoint can't locate the PPTX file. Fix: In main.py, ensure you are passing absolute paths for input_pptx_path and output_pdf_path to the convert_pptx_to_pdf method:
Python

input_pptx_path = os.path.abspath(os.path.join(PPTX_SOURCE_FOLDER, file))
output_pdf_path = os.path.abspath(os.path.join(PDF_OUTPUT_FOLDER, pdf_name))
KeyError: 'OPENAI_API_KEY environment variable not set.' (or similar for Groq/Gemini): You haven't set your API key environment variable correctly. Fix: Go back to "3. Configure API Keys" and ensure the export or set command was executed in your current terminal session before running main.py.
No .pptx files found in 'pptbase'.: Your pptbase folder is empty or doesn't exist. Fix: Create the pptbase folder and place your .pptx files inside it.
ModuleNotFoundError: No module named 'comtypes' (or other libraries): You haven't installed all necessary libraries. Fix: Make sure you have activated your virtual environment and run pip install -r requirements.txt.