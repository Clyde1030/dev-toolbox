import os
import re
import logging
import json
from PyPDF2 import PdfMerger

LOGGER = logging.getLogger()

def main():
    configure_logger(LOGGER, 'debug')  # change to debug for more visibility
    pages_to_merge = load_pages_config("pages_config.json")
    
    if not pages_to_merge:
        LOGGER.error("No pages loaded from config. Exiting.")
        return

    combined_pdf = "01_250-36976_signed.pdf"
    combine_pdfs(pages_to_merge, combined_pdf)
    LOGGER.info("PDFs combined successfully.")

def configure_logger(logger, level='info'):
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s - %(message)s', '%Y-%m-%d %H:%M:%S')
    handler = logging.StreamHandler()
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.setLevel(level.upper())    

def load_pages_config(json_file):
    if not os.path.exists(json_file):
        LOGGER.warning("No JSON config found. Exiting.")
        return {}
    with open(json_file, 'r') as f:
        raw_data = json.load(f)

    adjusted_data = {}
    for file, ranges in raw_data.items():
        new_ranges = []
        if not isinstance(ranges, list):
            LOGGER.warning(f"Invalid page range format for {file}, skipping.")
            continue
        for r in ranges:
            if isinstance(r, int):  # shorthand: single page
                start_idx = max(0, r - 1)
                stop_idx = start_idx + 1
                new_ranges.append((start_idx, stop_idx))
            elif isinstance(r, list) and len(r) == 2:  # range: [start, end]
                start_idx = max(0, r[0] - 1)
                stop_idx = r[1]  # stop is inclusive in JSON, exclusive in PyPDF2
                new_ranges.append((start_idx, stop_idx))
            else:
                LOGGER.warning(f"Invalid range {r} for file {file}, skipping.")
        adjusted_data[file] = new_ranges
    # LOGGER.debug("Adjusted page config: %s", adjusted_data)
    print(adjusted_data)
    return adjusted_data

def combine_pdfs(pages_to_merge, output_file):
    merger = PdfMerger()
    folder = os.path.join(os.getcwd(), 'files_to_merge')
    appended_count = 0

    for file, ranges in pages_to_merge.items():  # order preserved from JSON
        pdf_path = os.path.join(folder, file)
        if not os.path.exists(pdf_path):
            LOGGER.warning(f"File {pdf_path} not found, skipping.")
            continue

        if ranges:
            for start, stop in ranges:
                LOGGER.info(f"Appending {file} pages {start+1} to {stop} (human page numbers)")
                merger.append(pdf_path, pages=(start, stop))
                appended_count += 1
        else:
            LOGGER.info(f"Appending all pages from {file}")
            merger.append(pdf_path)
            appended_count += 1

    if appended_count == 0:
        LOGGER.error("No pages appended! Check JSON and file names.")
        return

    merger.write(output_file)
    merger.close()
    LOGGER.info(f"Created: {output_file}")


if __name__ == "__main__":
    main()