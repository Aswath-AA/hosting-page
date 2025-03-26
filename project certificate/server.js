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

        // Verify template exists
        if (!fs.existsSync(templatePath)) {
            console.error(`Template file missing: ${templatePath}`);
            return res.status(400).json({ 
                success: false,
                error: "Template not found",
                details: `Please upload ${templateName} to the templates directory`
            });
        }

        // Load and update Excel template
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

        // Attempt PDF conversion
        let pdfSuccess = false;
        let pdfError = null;
        
        try {
            pdfSuccess = await convertExcelToPDF(excelFilePath, pdfFilePath);
        } catch (error) {
            console.error("PDF conversion error:", error);
            pdfError = error.message;
        }

        // Prepare response
        const response = {
            success: true,
            excelPath: `/exports/${sanitizedSerialNo}_Certificate.xlsx`,
            serialNo: req.body.serialNo
        };

        if (pdfSuccess) {
            response.pdfPath = `/exports/${sanitizedSerialNo}_Certificate.pdf`;
        } else {
            response.pdfError = pdfError || "PDF generation failed";
        }

        res.json(response);

    } catch (error) {
        console.error("❌ Error processing request:", error);
        res.status(500).json({ 
            success: false,
            message: "Server error while processing the request!",
            details: error.message 
        });
    }
});

// Enhanced PDF conversion function with multiple fallbacks
async function convertExcelToPDF(excelPath, pdfPath) {
    try {
        console.log(`Attempting to convert ${excelPath} to PDF`);
        
        // Method 1: Try LibreOffice first (works on Linux/Windows/macOS)
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

// Robust Python-based Excel to PDF conversion
async function convertExcelToPDFWithPython(excelPath, pdfPath) {
    try {
        const pythonScript = path.join(pythonScriptsDir, "excel_to_pdf.py");
        
        // Create the python script if it doesn't exist
        if (!fs.existsSync(pythonScript)) {
            const pythonCode = `
import sys
import os
import time
import platform

def convert_excel_to_pdf(input_path, output_path, timeout=120):
    system = platform.system()
    
    # Windows implementation using pywin32
    if system == "Windows":
        try:
            import win32com.client as win32
            
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False

            workbook = excel.Workbooks.Open(input_path)
            workbook.Worksheets(1).ExportAsFixedFormat(
                Type=0, Filename=output_path, Quality=0,
                IncludeDocProperties=True, IgnorePrintAreas=False
            )
            
            # Wait for file creation
            start_time = time.time()
            while not os.path.exists(output_path):
                if time.time() - start_time > timeout:
                    raise TimeoutError("PDF generation timed out")
                time.sleep(1)
            
            return True
            
        except Exception as e:
            print(f"Windows conversion error: {e}", file=sys.stderr)
            return False
        finally:
            if 'workbook' in locals():
                workbook.Close(SaveChanges=False)
            if 'excel' in locals():
                excel.Quit()
    
    # macOS/Linux implementation using libreoffice
    else:
        try:
            import subprocess
            result = subprocess.run([
                "libreoffice", "--headless", "--convert-to", "pdf",
                "--outdir", os.path.dirname(output_path), input_path
            ], capture_output=True, text=True, timeout=timeout)
            
            if result.returncode == 0:
                temp_pdf = os.path.join(
                    os.path.dirname(output_path),
                    os.path.basename(input_path).replace('.xlsx', '.pdf')
                if os.path.exists(temp_pdf):
                    os.rename(temp_pdf, output_path)
                    return True
            return False
            
        except Exception as e:
            print(f"libreoffice failed: {e}", file=sys.stderr)
            return False

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python excel_to_pdf.py <input.xlsx> <output.pdf>", file=sys.stderr)
        sys.exit(1)
        
    input_path = sys.argv[1]
    output_path = sys.argv[2]
    
    if not os.path.exists(input_path):
        print(f"Error: Input file not found at {input_path}", file=sys.stderr)
        sys.exit(1)
        
    success = convert_excel_to_pdf(input_path, output_path)
    sys.exit(0 if success else 1)
`;
            fs.writeFileSync(pythonScript, pythonCode);
        }

        // Determine Python command (python or python3)
        let pythonCmd = 'python';
        try {
            await execPromise('python --version');
        } catch {
            try {
                await execPromise('python3 --version');
                pythonCmd = 'python3';
            } catch {
                throw new Error("Python is not available on this system");
            }
        }

        // Run the conversion
        const { stdout, stderr } = await execPromise(
            `${pythonCmd} "${pythonScript}" "${excelPath}" "${pdfPath}"`
        );
        
        if (stderr) console.error("Python stderr:", stderr);
        return fs.existsSync(pdfPath);
        
    } catch (error) {
        console.error("Python conversion error:", error);
        return false;
    }
}

// File cleanup endpoint
app.post("/cleanup", (req, res) => {
    try {
        const files = fs.readdirSync(exportsDir);
        const now = Date.now();
        const oneHour = 60 * 60 * 1000;
        let deletedCount = 0;
        
        files.forEach(file => {
            const filePath = path.join(exportsDir, file);
            try {
                const stat = fs.statSync(filePath);
                if (now - stat.mtimeMs > oneHour) {
                    fs.unlinkSync(filePath);
                    console.log(`Deleted old file: ${file}`);
                    deletedCount++;
                }
            } catch (err) {
                console.error(`Error deleting ${file}:`, err);
            }
        });
        
        res.json({ success: true, deleted: deletedCount });
    } catch (error) {
        console.error("Cleanup error:", error);
        res.status(500).json({ error: "Cleanup failed", details: error.message });
    }
});

// Health check endpoint
app.get("/health", (req, res) => {
    res.status(200).json({ 
        status: "healthy",
        timestamp: new Date().toISOString(),
        directories: {
            exports: exportsDir,
            uploads: uploadsDir,
            templates: templatesDir
        }
    });
});

// Start Server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 Server running on http://localhost:${PORT}`);
    console.log(`📁 Templates directory: ${templatesDir}`);
    console.log(`📁 Exports directory: ${exportsDir}`);
    console.log(`🐍 Python scripts directory: ${pythonScriptsDir}`);
});
