import os
from summary import SlideSummarizer
from read_json import ReadJSON
from ppt_to_file import PPTToFile

def ensure_output_dirs():
    os.makedirs("pdf_output", exist_ok=True)
    os.makedirs("text_output", exist_ok=True)

def main():
    pptx_folder = "pptbase"
    pdf_folder = "pdf_output"
    json_file = "slide_summaries.json"

    ensure_output_dirs()

    for file in os.listdir(pptx_folder):
        if file.endswith(".pptx"):
            pdf_name = file.replace(".pptx", ".pdf")
            pptx_path = os.path.abspath(os.path.join(pptx_folder, file))
            pdf_path = os.path.abspath(os.path.join(pdf_folder, pdf_name))


            print(f"Converting {file} to PDF...")
            converter = PPTToFile(pptx_path, pdf_path)
            converter.convert_pptx_to_pdf()

    summarizer = SlideSummarizer(
        folder_path=pptx_folder,
        model="llama-3.1-8b-instant",
        provider="groq"
    )
    summaries = summarizer.summarize_all()

    json_writer = ReadJSON(file=json_file)
    json_writer.write_file(summaries)

    print("\n🔍 Ready for search!")
    try:
        reader = ReadJSON(file=json_file)
        all_summaries = reader.data

        all_slide_map = {}
        for pptx_name, slides in all_summaries.items():
            for i, summary in enumerate(slides):
                all_slide_map[(pptx_name, i)] = summary

        query = input("Enter your query to find a relevant slide: ")
        viewer = PPTToFile(None, None)
        viewer.search_with_rag_pipeline(all_slide_map, top_k=3, query=query)

    except Exception as e:
        print(f"Search failed: {e}")

if __name__ == "__main__":
    main()
