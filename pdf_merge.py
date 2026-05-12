"""
PDF Merger
Merges multiple PDF files into a single PDF document.
"""

import os
import re
import sys
from pypdf import PdfWriter, PdfReader
from pathlib import Path



def _natural_sort_key(path):
    """Sort key for natural order: page1, page2, page10, page20 (not page1, page10, page2, page20)."""
    name = path.name if isinstance(path, Path) else os.path.basename(path)
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r'(\d+)', name)]


def merge_pdfs(input_files, output_file="merged_output.pdf"):
    """
    Merges multiple PDF files into a single PDF.
    
    Args:
        input_files: List of paths to PDF files to merge
        output_file: Path for the output merged PDF file (default: merged_output.pdf)
    
    Returns:
        bool: True if successful, False otherwise
    """
    if not input_files:
        print("Error: No PDF files provided.")
        return False
    
    # Validate input files
    valid_files = []
    for file_path in input_files:
        if not os.path.exists(file_path):
            print(f"Warning: File '{file_path}' not found. Skipping.")
            continue
        if not file_path.lower().endswith('.pdf'):
            print(f"Warning: '{file_path}' is not a PDF file. Skipping.")
            continue
        valid_files.append(file_path)
    
    if not valid_files:
        print("Error: No valid PDF files to merge.")
        return False
    
    print(f"\nMerging {len(valid_files)} PDF files...")
    
    try:
        # Create PDF writer object
        writer = PdfWriter()
        
        # Add each PDF file
        for idx, pdf_file in enumerate(valid_files, 1):
            print(f"  [{idx}/{len(valid_files)}] Adding: {os.path.basename(pdf_file)}")
            reader = PdfReader(pdf_file)
            for page in reader.pages:
                writer.add_page(page)
        
        # Write the merged PDF
        print(f"\nWriting merged PDF to: {output_file}")
        with open(output_file, 'wb') as output:
            writer.write(output)
        
        print(f"✓ Successfully merged {len(valid_files)} PDFs into '{output_file}'")
        return True
        
    except Exception as e:
        print(f"✗ Error merging PDFs: {e}")
        import traceback
        traceback.print_exc()
        return False


def merge_pdfs_from_folder(folder_path, output_file="merged_output.pdf", pattern="*.pdf"):
    """
    Merges all PDF files in a folder into a single PDF.
    
    Args:
        folder_path: Path to folder containing PDF files
        output_file: Path for the output merged PDF file
        pattern: File pattern to match (default: *.pdf)
    
    Returns:
        bool: True if successful, False otherwise
    """
    if not os.path.exists(folder_path):
        print(f"Error: Folder '{folder_path}' not found.")
        return False
    
    # Find all PDF files in the folder (natural sort: page1, page2, page10, page20)
    pdf_files = sorted(Path(folder_path).glob(pattern), key=_natural_sort_key)
    pdf_files = [str(f) for f in pdf_files if f.is_file()]
    
    if not pdf_files:
        print(f"Error: No PDF files found in '{folder_path}'")
        return False
    
    print(f"Found {len(pdf_files)} PDF files in folder: {folder_path}")
    return merge_pdfs(pdf_files, output_file)


def main():
    """Main function to handle command-line usage."""
    print("=" * 60)
    print("PDF Merger Tool")
    print("=" * 60)
    
    if len(sys.argv) > 1:
        # Command-line mode
        input_args = sys.argv[1:]
        
        # Check if first argument is a folder
        if len(input_args) == 1 and os.path.isdir(input_args[0]):
            folder_path = input_args[0]
            output_file = os.path.join(folder_path, "merged_output.pdf")
            merge_pdfs_from_folder(folder_path, output_file)
        else:
            # Multiple files provided
            pdf_files = [f.strip('"') for f in input_args]
            merge_pdfs(pdf_files)
    else:
        # Interactive mode
        print("\nOptions:")
        print("  1. Merge specific PDF files")
        print("  2. Merge all PDFs in a folder")
        choice = input("\nEnter your choice (1 or 2): ").strip()
        
        if choice == "1":
            # Get list of files
            print("\nEnter PDF file paths (one per line, empty line to finish):")
            pdf_files = []
            while True:
                file_path = input().strip('"').strip()
                if not file_path:
                    break
                pdf_files.append(file_path)
            
            if pdf_files:
                output_file = input("\nEnter output file name (press Enter for 'merged_output.pdf'): ").strip('"').strip()
                if not output_file:
                    output_file = "merged_output.pdf"
                merge_pdfs(pdf_files, output_file)
            else:
                print("No files provided.")
        
        elif choice == "2":
            # Get folder path
            folder_path = input("\nEnter folder path: ").strip('"').strip()
            output_file = input("Enter output file name (press Enter for 'merged_output.pdf'): ").strip('"').strip()
            if not output_file:
                output_file = os.path.join(folder_path, "merged_output.pdf")
            merge_pdfs_from_folder(folder_path, output_file)
        
        else:
            print("Invalid choice.")


if __name__ == "__main__":
    main()
