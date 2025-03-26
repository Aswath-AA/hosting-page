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

const isWindows = process.platform === 'win32';
const isRender = process.env.RENDER === 'true'; // Example for Render.com

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
            response.pdfPath = `${sanitizedSerialNo}_Certificate.pdf`;
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
async function convertExcelToPDF(excelPath, pdfPath, formData) {
    console.log(`Attempting to convert ${excelPath} to PDF`);
    
    // Method 1: Try LibreOffice first
    try {
        console.log('Trying LibreOffice conversion...');
        await execPromise(`libreoffice --headless --convert-to pdf "${excelPath}" --outdir "${path.dirname(pdfPath)}"`);
        
        // Verify and rename the output
        const tempPdf = path.join(
            path.dirname(pdfPath),
            path.basename(excelPath).replace('.xlsx', '.pdf')
        );
        
        if (fs.existsSync(tempPdf)) {
            fs.renameSync(tempPdf, pdfPath);
            console.log('✅ PDF generated via LibreOffice');
            return true;
        }
    } catch (libreOfficeError) {
        console.warn('⚠️ LibreOffice failed:', libreOfficeError.message);
    }
    
    // Method 2: Try Python script
    try {
        console.log('Trying Python conversion...');
        const pythonSuccess = await convertExcelToPDFWithPython(excelPath, pdfPath);
        if (pythonSuccess) {
            console.log('✅ PDF generated via Python');
            return true;
        }
    } catch (pythonError) {
        console.warn('⚠️ Python conversion failed:', pythonError.message);
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
