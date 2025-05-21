import fitz  # PyMuPDF
import re
import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional, Tuple
import warnings

class LayoutAwareRevisionExtractor:
    def __init__(self):
        self.drawing_pattern = re.compile(r"(\d{4}_SOC-[A-Z0-9\-]+-ARC0-\d{5})")
        self.rev_pattern = re.compile(r"\b(C\d{2}|P\d{2})\b")
        self.date_pattern = re.compile(r"\d{4}[./-]\d{2}[./-]\d{2}")

    def extract_from_filename(self, filename: str) -> Tuple[List[str], str, str]:
        drawings = self.drawing_pattern.findall(filename)
        revs = self.rev_pattern.findall(filename)
        c_rev = next((r for r in revs if r.startswith('C')), "")
        p_rev = next((r for r in revs if r.startswith('P')), "")
        return drawings, c_rev, p_rev

    def extract_from_blocks(self, page) -> List[Dict[str, str]]:
        results = []
        blocks = page.get_text("blocks")

        for block in blocks:
            text = block[4]
            drawings = self.drawing_pattern.findall(text)
            revs = self.rev_pattern.findall(text)

            for drawing in drawings:
                c = next((r for r in revs if r.startswith("C")), "")
                p = next((r for r in revs if r.startswith("P")), "")
                results.append({"drawing": drawing, "c_rev": c, "p_rev": p})

        return results

    def extract_revision_dates(self, text: str, fallback_date: str) -> Tuple[str, str]:
        dates = self.date_pattern.findall(text)
        dates = [d.replace(".", "/").replace("-", "/") for d in dates]
        c_date = dates[0] if dates else fallback_date
        p_date = dates[-1] if dates else fallback_date
        return c_date, p_date

    def parse_folder_date(self, folder_name: str) -> Optional[str]:
        match = self.date_pattern.search(folder_name)
        if not match:
            return None
        date_str = match.group().replace(".", "/").replace("-", "/")
        try:
            datetime.strptime(date_str, "%Y/%m/%d")
            return date_str
        except ValueError:
            return None

    def process_pdf(self, pdf_path: Path, fallback_date: str) -> List[Dict]:
        filename = pdf_path.name
        file_drawings, file_c_rev, file_p_rev = self.extract_from_filename(filename)
        record_map = {}

        try:
            with fitz.open(pdf_path) as doc:
                for page_num, page in enumerate(doc, start=1):
                    block_results = self.extract_from_blocks(page)
                    text = page.get_text()
                    c_date, p_date = self.extract_revision_dates(text, fallback_date)

                    for entry in block_results:
                        drawing = entry["drawing"]
                        c = entry["c_rev"] or file_c_rev or "X"  # Mark 'X' if no revision found
                        p = entry["p_rev"] or file_p_rev or "X"  # Mark 'X' if no revision found

                        if drawing not in record_map:
                            record_map[drawing] = {
                                "drawing": drawing,
                                "c_rev": c,
                                "p_rev": p,
                                "c_rev_date": c_date,
                                "p_rev_date": p_date,
                                "filename": filename,
                                "page": page_num
                            }
        except Exception as e:
            warnings.warn(f"Error processing {pdf_path}: {e}")

        return list(record_map.values())

    def find_all_pdfs(self, folder_path: Path) -> List[Path]:
        # Filter for only ARC0 PDFs
        return [pdf for pdf in folder_path.rglob("*.pdf") if "ARC0" in pdf.name]

    def process_date_folder(self, folder_path: Path) -> Tuple[str, List[Dict]]:
        folder_date = self.parse_folder_date(folder_path.name)
        if not folder_date:
            warnings.warn(f"Skipping invalid folder date: {folder_path}")
            return None, []

        pdfs = self.find_all_pdfs(folder_path)
        all_records = []
        for pdf in pdfs:
            all_records.extend(self.process_pdf(pdf, fallback_date=folder_date))

        print(f"Processed {folder_date}: {len(all_records)} records from {len(pdfs)} PDFs")
        return folder_date, all_records

    def create_master_table(self, data: Dict[str, List[Dict]]) -> pd.DataFrame:
        all_drawings = sorted(set(r["drawing"] for records in data.values() for r in records))
        df = pd.DataFrame({"NUMBER": all_drawings})

        for date in sorted(data.keys(), key=lambda d: datetime.strptime(d, "%Y/%m/%d")):
            temp_df = pd.DataFrame({
                "NUMBER": [r["drawing"] for r in data[date]],
                f"{date} | C Revision": [r["c_rev"] for r in data[date]],
                f"{date} | P Revision": [r["p_rev"] for r in data[date]]
            })
            temp_df = temp_df.drop_duplicates(subset="NUMBER")
            df = df.merge(temp_df, on="NUMBER", how="left")

        return df

def main():
    extractor = LayoutAwareRevisionExtractor()
    main_folder = Path(r"C:\Users\jukk\OneDrive - COWI\Documents\Revision check")

    if not main_folder.exists():
        raise FileNotFoundError("Main folder not found.")

    folders = [f for f in main_folder.iterdir() if f.is_dir() and extractor.parse_folder_date(f.name)]

    all_data = {}
    for folder in folders:
        date, records = extractor.process_date_folder(folder)
        if date and records:
            all_data[date] = records

    if not all_data:
        raise ValueError("No valid drawing data found.")

    master_df = extractor.create_master_table(all_data)
    output_file = main_folder / "Revision_Summary.xlsx"
    master_df.to_excel(output_file, index=False)
    print(f"✅ Saved revision summary to: {output_file}")

if __name__ == "__main__":
    main()
