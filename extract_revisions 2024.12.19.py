import fitz  # PyMuPDF
import re
import pandas as pd
from pathlib import Path

def extract_revisions_from_pdfs(folder_path: Path):
    drawing_pattern = re.compile(r"\d{4}_SOC-[A-Z0-9\-]+-ARC0-\d{5}")
    c_rev_pattern = re.compile(r"C\d{2}")
    p_rev_pattern = re.compile(r"P\d{2}")

    # Extract fallback date from folder name (e.g. 2024.12.19 -> 2024/12/19)
    folder_date_match = re.search(r"\d{4}[.\-_]\d{2}[.\-_]\d{2}", str(folder_path))
    fallback_date = ""
    if folder_date_match:
        fallback_date = folder_date_match.group().replace(".", "/").replace("-", "/")

    records = []
    pdf_files = list(folder_path.rglob("*.pdf"))  # Recursively search for all PDFs in the folder and subfolders
    
    if not pdf_files:
        print(f"No PDF files found in {folder_path}")
        return

    # To track duplicates by drawing number
    processed_drawings = set()

    for pdf_file in pdf_files:
        print(f"Processing {pdf_file.name}...")
        try:
            with fitz.open(pdf_file) as doc:
                for page_num, page in enumerate(doc, start=1):
                    text = page.get_text()

                    # Handle "Reference Drawings" exclusion
                    ref_section = ""
                    main_text = text
                    if "REFERENCE DRAWINGS" in text.upper():
                        split_parts = text.upper().split("REFERENCE DRAWINGS")
                        main_text = split_parts[0]
                        ref_section = split_parts[1] if len(split_parts) > 1 else ""

                    # Drawing numbers
                    main_drawings = set(drawing_pattern.findall(main_text))
                    ref_drawings = set(drawing_pattern.findall(ref_section))
                    final_drawings = main_drawings - ref_drawings

                    # Fallback to filename if drawing number not in content
                    if not final_drawings:
                        filename_drawings = drawing_pattern.findall(pdf_file.name)
                        if filename_drawings:
                            final_drawings = set(filename_drawings)

                    # Try to get revisions from content
                    c_revs = c_rev_pattern.findall(text)
                    p_revs = p_rev_pattern.findall(text)

                    # Fallback to filename if missing
                    file_c_rev = c_rev_pattern.search(pdf_file.name)
                    file_p_rev = p_rev_pattern.search(pdf_file.name)

                    c_rev = c_revs[0] if c_revs else (file_c_rev.group() if file_c_rev else "")
                    p_rev = p_revs[-1] if p_revs else (file_p_rev.group() if file_p_rev else "")

                    # Only add unique drawing numbers to the records
                    for drawing in final_drawings:
                        if drawing not in processed_drawings:
                            processed_drawings.add(drawing)
                            records.append({
                                "Drawing Number": drawing,
                                "C Revision": c_rev,
                                "P Revision": p_rev,
                                "PDF File": pdf_file.name,
                                "Page": page_num
                            })

        except Exception as e:
            print(f"Error processing {pdf_file.name}: {e}")

    if records:
        df = pd.DataFrame(records)
        output_file = folder_path / "revision_summary.xlsx"
        df.to_excel(output_file, index=False)
        print(f"\n✅ Extraction complete! Saved to:\n{output_file}")
    else:
        print("No ARC0 drawings found in the PDFs.")

if __name__ == "__main__":
    # Define the folder path where the PDFs are located
    folder = Path(r"C:\Users\jukk\OneDrive - COWI\Documents\Revision check\2024.12.19")
    extract_revisions_from_pdfs(folder)
