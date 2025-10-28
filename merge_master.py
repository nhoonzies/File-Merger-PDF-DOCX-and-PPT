import glob
import os
from docx import Document  # For Word files
from pypdf import PdfWriter  # For PDF files

# --- Configuration ---

# 1. Set the type of merge you want to perform.
#    Options: 'word' or 'pdf'
MERGE_TYPE = 'pdf'

# 2. This is the folder where ALL your files are located.
SOURCE_FOLDER = "files_to_merge"

# 3. This is the base name for the final merged file.
#    The script will add '.docx' or '.pdf' automatically.
OUTPUT_FILENAME_BASE = "master_document"


# --- End of Configuration ---


def get_merge_files(directory, file_extension):
    """
    Finds all files with a specific extension in the source directory
    and sorts them alphabetically.
    """
    # Create the full search path (e.g., "files_to_merge/*.docx")
    search_path = os.path.join(directory, f"*.{file_extension}")

    # Find all files matching the pattern and sort them
    files = sorted(glob.glob(search_path))

    return files


def merge_word_documents(files_list, output_filename):
    """
    Merges a list of .docx files into a single document with page breaks.
    """
    if not files_list:
        print(f"No .docx files found in '{SOURCE_FOLDER}'.")
        return

    print(f"Found {len(files_list)} Word files to merge.")
    master_doc = Document()
    total_files = len(files_list)

    for i, file_path in enumerate(files_list):
        print(f"Merging: {file_path}")
        source_doc = Document(file_path)

        # Copy content
        for element in source_doc.element.body:
            master_doc.element.body.append(element)

        # Add a page break unless it's the last file
        if i < total_files - 1:
            print("  -> Adding page break...")
            master_doc.add_page_break()

    try:
        master_doc.save(output_filename)
        print(f"\nSuccessfully merged {total_files} files into '{output_filename}'")
    except Exception as e:
        print(f"\nAn error occurred while saving the file: {e}")


def merge_pdf_documents(files_list, output_filename):
    """
    Merges a list of .pdf files into a single document.
    """
    if not files_list:
        print(f"No .pdf files found in '{SOURCE_FOLDER}'.")
        return

    print(f"Found {len(files_list)} PDF files to merge.")
    merger = PdfWriter()
    total_files = len(files_list)

    for i, file_path in enumerate(files_list):
        print(f"Merging: {file_path}")
        try:
            merger.append(file_path)
        except Exception as e:
            print(f"  -> ERROR: Could not merge '{file_path}'. Skipping. Error: {e}")

    try:
        print("\nSaving final document...")
        merger.write(output_filename)
        merger.close()
        print(f"Successfully merged {total_files} files into '{output_filename}'")
    except Exception as e:
        print(f"\nAn error occurred while saving the file: {e}")


# --- Main execution ---
if __name__ == "__main__":
    # Check if the source folder exists
    if not os.path.exists(SOURCE_FOLDER):
        os.makedirs(SOURCE_FOLDER)
        print(f"Created source directory: '{SOURCE_FOLDER}'")
        print("Please add your files to this folder and run the script again.")
    else:
        # --- Route to the correct function based on MERGE_TYPE ---

        if MERGE_TYPE.lower() == 'word':
            print("--- Starting WORD merge ---")
            files = get_merge_files(SOURCE_FOLDER, "docx")
            output_file = f"{OUTPUT_FILENAME_BASE}.docx"
            merge_word_documents(files, output_file)

        elif MERGE_TYPE.lower() == 'pdf':
            print("--- Starting PDF merge ---")
            files = get_merge_files(SOURCE_FOLDER, "pdf")
            output_file = f"{OUTPUT_FILENAME_BASE}.pdf"
            merge_pdf_documents(files, output_file)

        else:
            print(f"Error: Invalid MERGE_TYPE: '{MERGE_TYPE}'")
            print("Please open the script and set MERGE_TYPE to either 'word' or 'pdf'.")