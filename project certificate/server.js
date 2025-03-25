const express = require("express");
const cors = require("cors");
const bodyParser = require("body-parser");
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const multer = require("multer");
const { exec } = require('child_process');
const util = require('util');
const execPromise = util.promisify(exec);

const app = express();
app.use(cors());
app.use(bodyParser.json());

// Configure Multer for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, path.join(__dirname, "uploads"));
    },
    filename: (req, file, cb) => {
        cb(null, `signature-${Date.now()}.${file.originalname.split(".").pop()}`);
    },
});

const upload = multer({ storage });

// Ensure required directories exist
const exportsDir = path.join(__dirname, "exports");
const uploadsDir = path.join(__dirname, "uploads");
const templatesDir = path.join(__dirname, "templates");
const pythonScriptsDir = path.join(__dirname, "python_scripts");

[exportsDir, uploadsDir, templatesDir, pythonScriptsDir].forEach(dir => {
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
    }
});

// Serve static files
app.use("/exports", express.static(exportsDir));
app.use("/uploads", express.static(uploadsDir));
app.use(express.static(path.join(__dirname, "public")));

// Home route
app.get("/", (req, res) => {
    res.sendFile(path.join(__dirname, "public", "certificate.html"));
});

// Handle form submission
app.post("/update-excel", async (req, res) => {
    try {
        console.log("📌 Received Form Data:", req.body);

        // Sanitize the serial number for filename
        const sanitizedSerialNo = req.body.serialNo.replace(/[^a-zA-Z0-9-_]/g, '_');
        
        // Select Template Based on Mode
        const templateName = req.body.mode === "EN 53" ? "Template_EN_53.xlsx" : "Template_EN_73.xlsx";
        const templatePath = path.join(templatesDir, templateName);
        const excelFilePath = path.join(exportsDir, `${sanitizedSerialNo}_Certificate.xlsx`);
        const pdfFilePath = path.join(exportsDir, `${sanitizedSerialNo}_Certificate.pdf`);

        // Check if Template Exists
        if (!fs.existsSync(templatePath)) {
            console.error("❌ Excel template not found:", templatePath);
            return res.status(404).json({ message: "Excel template not found!" });
        }

        // Load Excel Template
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(templatePath);
        const worksheet = workbook.getWorksheet(1);

        // Insert Form Data
        worksheet.getCell("F10").value = req.body.mode;
        worksheet.getCell("F12").value = req.body.serialNo;
        worksheet.getCell("F16").value = req.body.testedDate;
        worksheet.getCell("F13").value = req.body.year;

        // Save the Updated Excel File
        await workbook.xlsx.writeFile(excelFilePath);
        console.log("✅ Excel file updated from", templateName);

        // Convert using Python (preferred method)
        try {
            await convertExcelToPDFWithPython(excelFilePath, pdfFilePath);
        } catch (pythonError) {
            console.warn("Python conversion failed, falling back to Puppeteer:", pythonError);
            await fallbackHTMLToPDFConversion(worksheet, req.body, req.files, pdfFilePath);
        }

        // Return File Paths with serial number in filename
        res.json({ 
            excelPath: `/exports/${sanitizedSerialNo}_Certificate.xlsx`, 
            pdfPath: `/exports/${sanitizedSerialNo}_Certificate.pdf`,
            serialNo: req.body.serialNo
        });

    } catch (error) {
        console.error("❌ Error processing request:", error);
        res.status(500).json({ message: "Server error while processing the request!" });
    }
});

// Python-based Excel to PDF conversion
async function convertExcelToPDFWithPython(excelPath, pdfPath) {
    const pythonScript = path.join(pythonScriptsDir, "excel_to_pdf.py");
    
    // Create the python script if it doesn't exist
    if (!fs.existsSync(pythonScript)) {
        const pythonCode = `
import win32com.client as win32
import os
import sys

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
        sys.exit(0 if success else 1)
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        sys.exit(1)
`;
        fs.writeFileSync(pythonScript, pythonCode);
    }

    const { stdout, stderr } = await execPromise(`python "${pythonScript}" "${excelPath}" "${pdfPath}"`);
    if (stderr) {
        throw new Error(stderr);
    }
    console.log("✅ PDF file generated with Python:", pdfPath);
}


// Start Server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 Server running on http://localhost:${PORT}`);
    console.log(`Python scripts directory: ${pythonScriptsDir}`);
});
