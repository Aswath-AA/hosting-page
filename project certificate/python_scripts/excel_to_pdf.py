import win32com.client as win32
import os
import sys
import json

def convert_excel_to_pdf(input_path, output_path):
    excel = None
    workbook = None
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False

        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Excel file not found at {input_path}")

        workbook = excel.Workbooks.Open(input_path)
        
        if workbook.Worksheets.Count < 1:
            raise ValueError("No worksheets found")
        
        workbook.Worksheets(1).ExportAsFixedFormat(
            Type=0,
            Filename=output_path,
            Quality=0,
            IncludeDocProperties=True,
            IgnorePrintAreas=False
        )
        
        return True
    except Exception as e:
        print(f"Conversion error: {str(e)}", file=sys.stderr)
        return False
    finally:
        if workbook:
            workbook.Close(SaveChanges=False)
        if excel:
            excel.Quit()

if __name__ == "__main__":
    try:
        input_path = sys.argv[1]
        output_path = sys.argv[2]
        
        success = convert_excel_to_pdf(input_path, output_path)
        
        # Output JSON result for Node.js to parse
        print(json.dumps({
            "success": success,
            "input": input_path,
            "output": output_path
        }))
        
        sys.exit(0 if success else 1)
    except Exception as e:
        print(json.dumps({"error": str(e)}), file=sys.stderr)
        sys.exit(1)