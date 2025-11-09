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
    root = Tk()
    root.withdraw()  # Hide the main window
    root.attributes('-topmost', True)  # Bring dialog to front
    
    print("\n" + "="*60)
    print("  BANK STATEMENT CATEGORIZATION TOOL")
    print("="*60)
    print("\nSelect one or more statement files (PDF or CSV)")
    print("Hold Ctrl/Cmd to select multiple files")
    
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
        script = 'wamo_categorization.py'
    elif file_type == 'csv':
        script = 'bov_categorization.py'
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
        # Immediately open file dialog
        statement_paths = select_files()
        
        if not statement_paths:
            print("\nNo files selected. Exiting.")
            input("\nPress Enter to exit...")
            return
        
        print(f"\n{len(statement_paths)} file(s) selected")
        print("="*60)
        
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
            
            # Ask if user wants to open the first file
            if successful == 1:
                open_msg = "\nWould you like to open the output file? (Y/n): "
            else:
                open_msg = "\nWould you like to open the first output file? (Y/n): "
            
            open_file = input(open_msg).strip().lower()
            if not open_file or open_file in ['y', 'yes']:
                try:
                    if sys.platform == 'win32':
                        os.startfile(output_files[0])
                    elif sys.platform == 'darwin':  # macOS
                        subprocess.run(['open', output_files[0]])
                    else:  # linux
                        subprocess.run(['xdg-open', output_files[0]])
                    print("Opening file...")
                except Exception as e:
                    print(f"Could not open file automatically: {e}")
                    print(f"Please open manually: {output_files[0]}")
        
        print("\n" + "="*60)
        input("\nPress Enter to exit...")
        
    except KeyboardInterrupt:
        print("\n\nInterrupted by user. Goodbye!")
        sys.exit(0)
    except Exception as e:
        print(f"\nUnexpected error: {e}")
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")
        sys.exit(1)


if __name__ == "__main__":
    main()
