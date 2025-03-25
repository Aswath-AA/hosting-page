const express = require("express");
const cors = require("cors");
const bodyParser = require("body-parser");
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const { exec } = require('child_process');
const util = require('util');
const execPromise = util.promisify(exec);
const PDFDocument = require('pdfkit'); // For fallback PDF generation

const app = express();
app.use(cors());
app.use(bodyParser.json());

// Configure directories
const exportsDir = path.join(__dirname, "exports");
const templatesDir = path.join(__dirname, "templates");

// Ensure directories exist
[exportsDir, templatesDir].forEach(dir => {
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
    }
});

// Serve static files
app.use("/exports", express.static(exportsDir));
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

        // Check if template exists
        if (!fs.existsSync(templatePath)) {
            console.error(`❌ Template file missing: ${templatePath}`);
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

        // Convert to PDF
        try {
            await convertExcelToPDF(excelFilePath, pdfFilePath);
            console.log("✅ PDF generated successfully");
        } catch (pdfError) {
            console.error("❌ PDF generation failed:", pdfError);
            await createFallbackPDF(pdfFilePath, req.body);
        }

        // Return file information
        res.json({ 
            success: true,
            excelPath: `/exports/${sanitizedSerialNo}_Certificate.xlsx`, 
            pdfPath: `/exports/${sanitizedSerialNo}_Certificate.pdf`,
            serialNo: req.body.serialNo,
            filename: `${sanitizedSerialNo}_Certificate.pdf`
        });

    } catch (error) {
        console.error("❌ Error processing request:", error);
        res.status(500).json({ 
            success: false,
            message: "Server error while processing the request!",
            error: error.message
        });
    }
});

// PDF Conversion Functions
async function convertExcelToPDF(excelPath, pdfPath) {
    try {
        // Method 1: Try LibreOffice first
        console.log("Attempting LibreOffice conversion...");
        await execPromise(`libreoffice --headless --convert-to pdf "${excelPath}" --outdir "${path.dirname(pdfPath)}"`);
        
        // Verify PDF was created
        if (!fs.existsSync(pdfPath)) {
            throw new Error("LibreOffice conversion failed - no PDF created");
        }
        return true;
    } catch (error) {
        console.log("LibreOffice failed, trying Python fallback...");
        try {
            await convertExcelToPDFWithPython(excelPath, pdfPath);
            return true;
        } catch (pythonError) {
            console.log("Python conversion failed, will use basic PDF fallback");
            throw pythonError;
        }
    }
}

async function convertExcelToPDFWithPython(excelPath, pdfPath) {
    const pythonScript = path.join(__dirname, "excel_to_pdf.py");
    
    if (!fs.existsSync(pythonScript)) {
        fs.writeFileSync(pythonScript, `
# Python script for Excel to PDF conversion
import sys
from win32com.client import Dispatch

def convert_excel_to_pdf(input_path, output_path):
    try:
        excel = Dispatch('Excel.Application')
        excel.Visible = False
        workbook = excel.Workbooks.Open(input_path)
        workbook.ExportAsFixedFormat(0, output_path)
        workbook.Close(False)
        excel.Quit()
        return True
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return False

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python excel_to_pdf.py <input.xlsx> <output.pdf>", file=sys.stderr)
        sys.exit(1)
    success = convert_excel_to_pdf(sys.argv[1], sys.argv[2])
    sys.exit(0 if success else 1)
`);
    }

    const { stdout, stderr } = await execPromise(`python "${pythonScript}" "${excelPath}" "${pdfPath}"`);
    if (stderr) {
        throw new Error(stderr);
    }
    if (!fs.existsSync(pdfPath)) {
        throw new Error("Python conversion failed - no PDF created");
    }
    return true;
}

async function createFallbackPDF(pdfPath, formData) {
    return new Promise((resolve, reject) => {
        const doc = new PDFDocument();
        const stream = fs.createWriteStream(pdfPath);
        
        doc.pipe(stream);
        
        // Add basic certificate content
        doc.fontSize(20).text('ELGi Certificate', { align: 'center' });
        doc.moveDown();
        doc.fontSize(14).text(`Serial Number: ${formData.serialNo}`);
        doc.text(`Mode: ${formData.mode}`);
        doc.text(`Tested Date: ${formData.testedDate}`);
        doc.text(`Year of Manufacturing: ${formData.year}`);
        
        doc.end();
        
        stream.on('finish', () => resolve());
        stream.on('error', (err) => reject(err));
    });
}

// Start Server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 Server running on http://localhost:${PORT}`);
    console.log(`Templates directory: ${templatesDir}`);
    console.log(`Exports directory: ${exportsDir}`);
});
