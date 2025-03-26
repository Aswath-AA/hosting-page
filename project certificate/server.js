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

// Serve static files with download headers
app.use("/exports", express.static(exportsDir, {
    setHeaders: (res, filePath) => {
        res.setHeader('Content-Disposition', `attachment; filename="${path.basename(filePath)}"`);
    }
}));
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

        if (!fs.existsSync(templatePath)) {
            console.error(`Template file missing: ${templatePath}`);
            return res.status(400).json({ 
                error: "Template not found",
                details: `Please upload ${templateName} to the templates directory`
            });
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

        // Attempt PDF conversion with multiple fallbacks
        const pdfSuccess = await convertExcelToPDF(excelFilePath, pdfFilePath);
        
        if (!pdfSuccess) {
            console.warn("PDF generation failed, returning Excel only");
            return res.json({ 
                excelPath: `/exports/${sanitizedSerialNo}_Certificate.xlsx`,
                serialNo: req.body.serialNo
            });
        }

        res.json({ 
            excelPath: `/exports/${sanitizedSerialNo}_Certificate.xlsx`, 
            pdfPath: `/exports/${sanitizedSerialNo}_Certificate.pdf`,
            serialNo: req.body.serialNo
        });

    } catch (error) {
        console.error("❌ Error processing request:", error);
        res.status(500).json({ 
            message: "Server error while processing the request!",
            details: error.message 
        });
    }
});

// Enhanced PDF conversion function with multiple fallbacks
async function convertExcelToPDF(excelPath, pdfPath) {
    try {
        console.log(`Attempting to convert ${excelPath} to PDF`);
        
        // Method 1: Try LibreOffice (works on Linux/Windows)
        try {
            console.log('Trying LibreOffice conversion...');
            await execPromise(`libreoffice --headless --convert-to pdf "${excelPath}" --outdir "${path.dirname(pdfPath)}"`);
            console.log('✅ PDF generated via LibreOffice');
            
            // Verify the file was created
            if (fs.existsSync(pdfPath)) {
                return true;
            }
        } catch (libreOfficeError) {
            console.warn('⚠️ LibreOffice failed:', libreOfficeError.message);
        }
        
        // Method 2: Try Python script (Windows with Excel installed)
        try {
            console.log('Trying Python conversion...');
            const pythonSuccess = await convertExcelToPDFWithPython(excelPath, pdfPath);
            if (pythonSuccess && fs.existsSync(pdfPath)) {
                return true;
            }
        } catch (pythonError) {
            console.warn('⚠️ Python conversion failed:', pythonError.message);
        }
        
        // Method 3: Fallback to HTML-to-PDF
        console.log('Falling back to HTML-to-PDF');
        try {
            await fallbackHTMLToPDFConversion(excelPath, pdfPath);
            return fs.existsSync(pdfPath);
        } catch (htmlPdfError) {
            console.error('⚠️ HTML-to-PDF failed:', htmlPdfError.message);
            return false;
        }
    } catch (error) {
        console.error('❌ All PDF conversion methods failed:', error);
        return false;
    }
}

// HTML-to-PDF fallback conversion
async function fallbackHTMLToPDFConversion(excelPath, pdfPath) {
    const pdf = require('html-pdf');
    const html = `
        <html>
            <head>
                <style>
                    body { font-family: Arial, sans-serif; margin: 2cm; }
                    h1 { color: #0066cc; }
                    table { border-collapse: collapse; width: 100%; }
                    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
                </style>
            </head>
            <body>
                <h1>Certificate of Conformity</h1>
                <p>This is an automatically generated certificate based on the Excel file:</p>
                <p><strong>${path.basename(excelPath)}</strong></p>
                <p>Please note: This is a simplified version. The full certificate is available in Excel format.</p>
            </body>
        </html>
    `;
    
    const options = {
        format: 'A4',
        border: {
            top: '1cm',
            right: '1cm',
            bottom: '1cm',
            left: '1cm'
        }
    };
    
    return new Promise((resolve, reject) => {
        pdf.create(html, options).toFile(pdfPath, (err, res) => {
            if (err) reject(err);
            else resolve(res);
        });
    });
}

// Python-based Excel to PDF conversion (Windows only)
async function convertExcelToPDFWithPython(excelPath, pdfPath) {
    try {
        const pythonScript = path.join(pythonScriptsDir, "excel_to_pdf.py");
        
        // Create the python script if it doesn't exist
        if (!fs.existsSync(pythonScript)) {
            const pythonCode = `
import win32com.client as win32
import os
import sys
import time

def convert_excel_to_pdf(input_path, output_path, timeout=120):
    excel = None
    start_time = time.time()
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
        
        # Wait for file to be created
        while not os.path.exists(output_path):
            if time.time() - start_time > timeout:
                raise TimeoutError("PDF generation timed out")
            time.sleep(1)
        
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
        return fs.existsSync(pdfPath);
    } catch (error) {
        console.error("Python conversion error:", error);
        return false;
    }
}

// File cleanup endpoint (optional)
app.post("/cleanup", (req, res) => {
    try {
        const files = fs.readdirSync(exportsDir);
        const now = Date.now();
        const oneHour = 60 * 60 * 1000;
        
        files.forEach(file => {
            const filePath = path.join(exportsDir, file);
            const stat = fs.statSync(filePath);
            if (now - stat.mtimeMs > oneHour) {
                fs.unlinkSync(filePath);
                console.log(`Deleted old file: ${file}`);
            }
        });
        
        res.json({ success: true, deleted: files.length });
    } catch (error) {
        console.error("Cleanup error:", error);
        res.status(500).json({ error: "Cleanup failed", details: error.message });
    }
});

// Start Server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 Server running on http://localhost:${PORT}`);
    console.log(`Templates directory: ${templatesDir}`);
    console.log(`Exports directory: ${exportsDir}`);
});
