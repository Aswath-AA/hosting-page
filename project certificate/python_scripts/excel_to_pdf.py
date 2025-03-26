const pythonCode = `
import sys
import os
import subprocess
import platform
import time
from pathlib import Path

def find_libreoffice():
    """Try to find LibreOffice in common installation locations"""
    possible_paths = [
        '/usr/bin/libreoffice',
        '/usr/local/bin/libreoffice',
        '/Applications/LibreOffice.app/Contents/MacOS/soffice',
        'libreoffice'  # Try in PATH
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            return path
        try:
            subprocess.run([path, '--version'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            return path
        except:
            continue
    return None

def convert_with_libreoffice(input_path, output_path, timeout=120):
    """Convert using LibreOffice with proper error handling"""
    libreoffice_path = find_libreoffice()
    if not libreoffice_path:
        print("LibreOffice not found in system", file=sys.stderr)
        return False

    try:
        # Create output directory if it doesn't exist
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        # Run LibreOffice conversion
        result = subprocess.run(
            [
                libreoffice_path,
                "--headless",
                "--convert-to", "pdf",
                "--outdir", os.path.dirname(output_path),
                input_path
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout
        )

        if result.returncode != 0:
            print(f"LibreOffice conversion failed with code {result.returncode}", file=sys.stderr)
            print(f"Error output: {result.stderr.decode()}", file=sys.stderr)
            return False

        # Find and rename the output file
        expected_temp_pdf = os.path.join(
            os.path.dirname(output_path),
            os.path.splitext(os.path.basename(input_path))[0] + ".pdf"
        )

        if os.path.exists(expected_temp_pdf):
            # If we got the expected output, rename it to the requested path
            os.replace(expected_temp_pdf, output_path)
            return True

        print("LibreOffice conversion completed but output file not found", file=sys.stderr)
        return False

    except subprocess.TimeoutExpired:
        print("LibreOffice conversion timed out", file=sys.stderr)
        return False
    except Exception as e:
        print(f"LibreOffice conversion error: {str(e)}", file=sys.stderr)
        return False

def convert_excel_to_pdf(input_path, output_path, timeout=120):
    """Main conversion function with platform-specific handling"""
    system = platform.system()
    
    if system == "Windows":
        try:
            import win32com.client as win32
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False

            workbook = excel.Workbooks.Open(os.path.abspath(input_path))
            workbook.ExportAsFixedFormat(
                Type=0,  # xlTypePDF
                Filename=os.path.abspath(output_path),
                Quality=0,  # xlQualityStandard
                IncludeDocProperties=True,
                IgnorePrintAreas=False
            )
            
            # Wait for file creation with timeout
            start_time = time.time()
            while not os.path.exists(output_path):
                if time.time() - start_time > timeout:
                    raise TimeoutError("PDF generation timed out")
                time.sleep(1)
            
            return True
            
        except ImportError:
            print("win32com not available - falling back to LibreOffice", file=sys.stderr)
            return convert_with_libreoffice(input_path, output_path, timeout)
        except Exception as e:
            print(f"Windows conversion error: {str(e)}", file=sys.stderr)
            return False
        finally:
            if 'workbook' in locals():
                workbook.Close(SaveChanges=False)
            if 'excel' in locals():
                excel.Quit()
    else:
        # Linux/macOS - always use LibreOffice
        return convert_with_libreoffice(input_path, output_path, timeout)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python excel_to_pdf.py <input.xlsx> <output.pdf>", file=sys.stderr)
        sys.exit(1)
        
    input_path = os.path.abspath(sys.argv[1])
    output_path = os.path.abspath(sys.argv[2])
    
    if not os.path.exists(input_path):
        print(f"Error: Input file not found at {input_path}", file=sys.stderr)
        sys.exit(1)
        
    success = convert_excel_to_pdf(input_path, output_path)
    sys.exit(0 if success else 1)
`;
