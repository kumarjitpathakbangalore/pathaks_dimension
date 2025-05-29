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
    * **Linux:**
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
- main.py
- summary.py             # Contains SlideSummarizer class
- read_json.py           # Contains JSONManager class (previously ReadJSON)
- ppt_to_file.py         # Contains PPTConverterAndSearch class (previously PPTToFile)
- requirements.txt
- pptbase/               # Create this folder
  - your_presentation_1.pptx
  - your_presentation_2.pptx
  - ...
- pdf_output/            # Will be created by the script (for converted PDFs)
- text_output/           # Will be created by the script (for JSON summaries)
- output_images/         # Will be created by the script (for displayed slide images)
- venv/                  # Your virtual environment


* **`pptbase/`**: Place all the PowerPoint (`.pptx`) files you want to process in this directory.

## 5. Run the Code

Once you have activated your virtual environment, installed dependencies, configured API keys, and placed your `.pptx` files in the `pptbase/` folder, you can run the main script.

```bash
python main.py
```
---

### What the Script Does:

Here's a breakdown of the steps the `main.py` script performs when you run it:

* **Converts PPTX to PDF:** The script iterates through your `pptbase/` folder. For each PowerPoint (`.pptx`) file it finds, it converts it into a PDF (`.pdf`) document, saving the result in the `pdf_output/` directory.
* **Summarizes Slides:** It then extracts all text from the `.pptx` files. Each slide's content is sent to your chosen Large Language Model (LLM)—be it **OpenAI**, **Groq**, or **Gemini**—to generate a concise, two-line summary. These summaries are then saved into a `slide_summaries.json` file located in the `text_output/` folder.
* **Prepares for Search:** The script loads the generated summaries and organizes them into a format optimized for the **FAISS** library, which is used for efficient similarity search.
* **Prompts for Query:** You'll be prompted to "Enter your query to find relevant slides:". This is where you input your natural language search term.
* **Performs Search:** Using semantic search, the script queries the indexed summaries to find the **top 3 most relevant slides** that match your input query.
* **Displays Results:** For each identified matching slide, the script will automatically **display the actual PDF page** using `matplotlib`, providing a visual confirmation of the retrieved content.

---

### 🪛 Troubleshooting

If you encounter any issues while running the script, refer to the common problems and their solutions below:

---

#### `Conversion error: ... Hiding the application window is not allowed.`

* **Reason:** Your Microsoft PowerPoint installation might have security settings or a version that prevents programmatic control when attempting to run PowerPoint in a completely invisible mode.
* **Fix:**
    1.  Open `ppt_to_file.py`.
    2.  Locate the line `powerpoint.Visible = 0` within the `convert_pptx_to_pdf` method.
    3.  Change it to `powerpoint.Visible = 1`.
    * **Note:** After this change, you will see PowerPoint windows briefly pop up on your screen as each conversion takes place.

---

#### `Conversion error: ... The system cannot find the path specified.`

* **Reason:** PowerPoint is unable to locate the `.pptx` file even though the Python script might be passing what looks like a correct path. This often happens if relative paths are used, and PowerPoint's internal working directory differs from your script's.
* **Fix:** In `main.py`, ensure that you are providing **absolute paths** for both the input PPTX file and the output PDF file when calling the `convert_pptx_to_pdf` method. Update the relevant section as follows:

    ```python
    input_pptx_path = os.path.abspath(os.path.join(PPTX_SOURCE_FOLDER, file))
    output_pdf_path = os.path.abspath(os.path.join(PDF_OUTPUT_FOLDER, pdf_name))
    ```

---

#### `KeyError: 'OPENAI_API_KEY environment variable not set.'` (or similar for Groq/Gemini)

* **Reason:** The script relies on environment variables to access your LLM API keys securely, and the required variable hasn't been set in your current terminal session.
* **Fix:**
    1.  Go back to the **"3. Configure API Keys"** section in the `README.md`.
    2.  Make sure you have executed the appropriate `export` (for macOS/Linux) or `set` (`PowerShell`/`CMD` for Windows) command for your chosen LLM provider **in the same terminal window** *before* running `python main.py`.

---

#### `No .pptx files found in 'pptbase'.`

* **Reason:** The `pptbase` folder either does not exist or is empty.
* **Fix:**
    1.  **Create** a folder named `pptbase` in your project's root directory.
    2.  **Place** all the `.pptx` files you wish to process inside this newly created `pptbase` folder.

---

#### `ModuleNotFoundError: No module named 'comtypes'` (or other libraries)

* **Reason:** One or more of the necessary Python libraries are not installed in your active virtual environment.
* **Fix:**
    1.  Ensure your virtual environment is **activated**.
    2.  Navigate to your project's root directory (where `requirements.txt` is located).
    3.  Run the command: `pip install -r requirements.txt` to install all required dependencies.