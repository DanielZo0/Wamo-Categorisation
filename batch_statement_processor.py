#!/usr/bin/env python3
"""
Bank Statement Categorization Tool
Unified interface for processing PDF and CSV bank statements
"""

import os
import sys
from pathlib import Path
from tkinter import Tk, filedialog
import subprocess


def select_files():
    """
    Open a file selection dialog to choose statement files (supports multiple)
    """
    print("\n" + "="*60)
    print("  BANK STATEMENT CATEGORISATION TOOL")
    print("="*60)
    print("\nOpening file selection dialog...")
    print("Select one or more statement files (PDF or CSV)")
    print("Hold Ctrl to select multiple files, or click Cancel to exit")
    print("="*60)
    
    root = Tk()
    root.withdraw()  # Hide the main window
    root.attributes('-topmost', True)  # Bring dialog to front
    
    file_paths = filedialog.askopenfilenames(
        title="Select Bank Statement Files",
        filetypes=[
            ("All Supported Files", "*.pdf *.csv"),
            ("PDF Files", "*.pdf"),
            ("CSV Files", "*.csv"),
            ("All Files", "*.*")
        ]
    )
    
    root.destroy()
    
    return list(file_paths) if file_paths else []


def detect_file_type(file_path):
    """
    Detect if file is PDF or CSV
    """
    extension = Path(file_path).suffix.lower()
    
    if extension == '.pdf':
        return 'pdf'
    elif extension == '.csv':
        return 'csv'
    else:
        return None


def get_output_path(input_path):
    """
    Generate output filename in the same directory as input
    Format: categorized_[original_name].xlsx
    """
    input_file = Path(input_path)
    input_dir = input_file.parent
    input_stem = input_file.stem
    
    output_name = f"categorized_{input_stem}.xlsx"
    return str(input_dir / output_name)


def process_statement(statement_path, output_path):
    """
    Process the statement using the appropriate script
    """
    file_type = detect_file_type(statement_path)
    
    if file_type == 'pdf':
        script = 'pdf_statement_processor.py'
    elif file_type == 'csv':
        script = 'csv_statement_processor.py'
    else:
        print(f"  [ERROR] Unsupported file type: {Path(statement_path).suffix}")
        return False
    
    # Run the appropriate categorization script
    try:
        result = subprocess.run(
            [sys.executable, script, statement_path, output_path],
            check=True,
            capture_output=False
        )
        return True
    except subprocess.CalledProcessError as e:
        print(f"  [ERROR] Processing failed: {e}")
        return False
    except FileNotFoundError:
        print(f"  [ERROR] Could not find {script}")
        return False


def main():
    """
    Main function with simplified workflow
    """
    try:
        # Check if files were passed as command-line arguments (e.g., double-clicked a PDF)
        if len(sys.argv) > 1:
            # Files were passed as arguments - use them directly
            statement_paths = [arg for arg in sys.argv[1:] if Path(arg).exists()]
            
            if statement_paths:
                print("\n" + "="*60)
                print("  BANK STATEMENT CATEGORISATION TOOL")
                print("="*60)
                print(f"\nProcessing {len(statement_paths)} file(s) passed as arguments")
            else:
                print("\nInvalid file path(s) provided.")
                input("\nPress Enter to exit...")
                sys.exit(0)
        else:
            # No arguments - open file dialog
            statement_paths = select_files()
            
            if not statement_paths:
                print("\nNo files selected. Exiting.")
                input("\nPress Enter to exit...")
                sys.exit(0)
        
        print(f"\n{len(statement_paths)} file(s) selected")
        print("="*60)
        
        # Debug: Show selected files
        for path in statement_paths:
            print(f"  - {Path(path).name}")
        
        successful = 0
        failed = 0
        output_files = []
        
        # Process each file
        for i, statement_path in enumerate(statement_paths, 1):
            file_name = Path(statement_path).name
            file_type = detect_file_type(statement_path)
            
            if not file_type:
                print(f"\n[{i}/{len(statement_paths)}] SKIPPED: {file_name}")
                print(f"  Unsupported file type")
                failed += 1
                continue
            
            output_path = get_output_path(statement_path)
            
            print(f"\n[{i}/{len(statement_paths)}] Processing: {file_name}")
            print(f"  Type: {file_type.upper()}")
            print(f"  Output: {Path(output_path).name}")
            print("-"*60)
            
            success = process_statement(statement_path, output_path)
            
            if success:
                successful += 1
                output_files.append(output_path)
                print(f"  [OK] Completed successfully")
            else:
                failed += 1
                print(f"  [FAILED] Processing error")
        
        # Summary
        print("\n" + "="*60)
        print("  PROCESSING COMPLETE")
        print("="*60)
        print(f"\nSuccessful: {successful}")
        print(f"Failed: {failed}")
        print(f"Total: {len(statement_paths)}")
        
        if output_files:
            print("\nGenerated files:")
            for output_file in output_files:
                print(f"  - {output_file}")
            
            # Get the directory of the first output file
            output_dir = str(Path(output_files[0]).parent)
            print(f"\nOutput location: {output_dir}")
        
        print("\n" + "="*60)
        input("\nPress Enter to open file location and exit...")
        
        # Open file explorer to the output location
        if output_files:
            try:
                if sys.platform == 'win32':
                    # Windows: Open explorer and select the first file
                    subprocess.run(['explorer', '/select,', str(Path(output_files[0]).resolve())])
                elif sys.platform == 'darwin':  # macOS
                    # macOS: Open Finder and select the file
                    subprocess.run(['open', '-R', output_files[0]])
                else:  # linux
                    # Linux: Open file manager to the directory
                    subprocess.run(['xdg-open', output_dir])
            except Exception as e:
                print(f"Could not open file location: {e}")
        
        # Exit after completion
        sys.exit(0)
        
    except KeyboardInterrupt:
        print("\n\nInterrupted by user.")
        sys.exit(0)
    except Exception as e:
        print(f"\nUnexpected error: {e}")
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")
        sys.exit(1)


if __name__ == "__main__":
    main()
