import os
from summary import SlideSummarizer
from read_json import JSONManager
from ppt_to_file import PPTConverterAndSearch

def ensure_base_directories(directories):
    """
    Ensures that a list of specified directories exist. If a directory
    does not exist, it will be created.
    """
    for directory in directories:
        os.makedirs(directory, exist_ok=True)
        print(f"📁 Ensured directory exists: {directory}")

def main():
    """
    Main function to orchestrate the PowerPoint processing workflow.
    """
    # --- Configuration ---
    PPTX_SOURCE_FOLDER = "pptbase"
    PDF_OUTPUT_FOLDER = "pdf_output"
    SUMMARIES_JSON_FILE = "slide_summaries.json"
    
    # Ensure the root folder for PPTX files exists
    if not os.path.exists(PPTX_SOURCE_FOLDER):
        print(f"Error: The source folder '{PPTX_SOURCE_FOLDER}' does not exist.")
        print("Please create this folder and place your .pptx files inside.")
        return

    # Ensure all necessary output directories exist
    ensure_base_directories([PDF_OUTPUT_FOLDER, "text_output"])
    
    # --- Step 1: Convert PPTX files to PDF ---
    print("\n--- Starting PowerPoint to PDF Conversion ---")
    converter = PPTConverterAndSearch() 
    
    pptx_files_found = False
    for file in os.listdir(PPTX_SOURCE_FOLDER):
        if file.endswith(".pptx"):
            pptx_files_found = True
            pdf_name = file.replace(".pptx", ".pdf")
            
            # Convert both input and output paths to absolute paths
            input_pptx_path = os.path.abspath(os.path.join(PPTX_SOURCE_FOLDER, file))
            output_pdf_path = os.path.abspath(os.path.join(PDF_OUTPUT_FOLDER, pdf_name))

            print(f"Converting '{file}' to PDF...")
            converter.convert_pptx_to_pdf(input_pptx_path, output_pdf_path)
    
    if not pptx_files_found:
        print(f"No .pptx files found in '{PPTX_SOURCE_FOLDER}'. Skipping PDF conversion.")
        print("Ensure your PowerPoint files are in the specified folder.")
        return
    
    # --- Step 2: Summarize PPTX slide content using an LLM ---
    print("\n--- Starting Slide Summarization ---")
    summarizer = SlideSummarizer(
        folder_path=PPTX_SOURCE_FOLDER,
        model="llama-3.1-8b-instant",
        provider="groq",
        save_path=os.path.join("text_output", SUMMARIES_JSON_FILE)
    )
    all_slide_summaries = summarizer.summarize_all()

    if not all_slide_summaries:
        print("No summaries generated. Exiting search phase.")
        return
        
    # --- Step 3: Prepare data for semantic search and execute query ---
    print("\n--- Preparing for Semantic Search ---")
    try:
        json_manager = JSONManager(filename=SUMMARIES_JSON_FILE)
        loaded_summaries = json_manager.get_data()

        all_slide_map_for_search = {}
        for pptx_filename, slides_list in loaded_summaries.items():
            for i, summary_text in enumerate(slides_list):
                all_slide_map_for_search[(pptx_filename, i)] = summary_text

        if not all_slide_map_for_search:
            print("No slide summaries available for search. Please check the summary generation process.")
            return

        print("\n🔍 Ready for your search query!")
        user_query = input("Enter your query to find relevant slides: ")
        
        # Reuse the existing converter object for searching
        converter.search_with_rag_pipeline(
            all_slide_map=all_slide_map_for_search, 
            query=user_query, 
            top_k=3
        )

    except Exception as e:
        print(f"❌ An error occurred during the search phase: {e}")

if __name__ == "__main__":
    main()