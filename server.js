const express = require("express");
const cors = require("cors");
const pdfParse = require("pdf-parse");
const fs = require("fs");
const path = require("path");
const json2xls = require("json2xls");
const multer = require("multer");
const XLSX = require("xlsx");
const dotenv = require("dotenv");
dotenv.config();

const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, "public"); // Files will be saved in the 'public' directory
  },
  filename: function (req, file, cb) {
    // Use the original file name or generate a new one
    cb(null, Date.now() + "-" + file.originalname);
  },
});

// Filter to only allow PDF uploads
const fileFilter = (req, file, cb) => {
  cb(null, true);
};

const upload = multer({ storage: storage, fileFilter: fileFilter });

const PORT = 5200;

const app = express();
app.use(cors());

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.use(express.static("public"));

app.get("/", (req, res) => {
  res.send("Hello World!");
});

// Function to combine text from every two pages
const combinePages = (text) => {
  // console.log(text)
  const blocks = text.split(/P R E - B I L L/);
  // console.log(blocks)
  return blocks
    .slice(1)
    .filter((block) => block.trim() !== "")
    .map((block) => "PRE-BILL" + block);
};
// Function to read Excel file and convert to JSON
function readExcelFile(filePath) {
  const workbook = XLSX.readFile(filePath);
  const sheetNameList = workbook.SheetNames;
  const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[0]]);
  return jsonData;
}
// Function to extract data for each person using regex
const extractData = (combinedRecords, pdfName) => {
  const records = [];

  // Regex patterns to match data blocks
  const fileRegex = /File #: (\S+)/;
  const feesRegex = /Total Fees\s*\n\s*([0-9.,]+)/;
  const disbursementsRegex = /Total Taxable Disbursements\s*\$([0-9.,]+)/;
  const hstRegex = /Total HST on Disbursements\s*\$([0-9.,]+)/;
  const totalRegex = /Total\s*\$([0-9.,]+)/;
  const arBalanceRegex = /A\/R Balance:\s*\$([0-9.,]+)/;
  const clientRegex = /Approved By: _______\s*\n\s*([A-Z][a-z]+ [A-Z][a-z]+)/;
  const dateRegex = /Date:\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})/;

  combinedRecords.forEach((block, page) => {
    const data = {};
    const fileMatch = block.match(fileRegex) || block.match(/File #:\s*(\S+)/);
    const clientMatch =
      block.match(clientRegex) ||
      block.match(/^\s*PRE-BILL\s*\n\s*(.*?)\s*\n/m);
    const feesMatch = block.match(feesRegex);
    const disbursementsMatch = block.match(disbursementsRegex);
    const hstMatch = block.match(hstRegex);
    const totalMatch = block.match(totalRegex);
    const arBalanceMatch = block.match(arBalanceRegex);
    const dateMatch = block.match(dateRegex);

    data.File = fileMatch ? fileMatch[1].replace(/Page/i, "") : null;
    data.Client = clientMatch ? clientMatch[1].trim() : null;
    data.Fees = feesMatch ? feesMatch[1] : 0;
    data.Date = dateMatch ? dateMatch[1] : null;

    data["Total Taxable Disbursements"] = disbursementsMatch
      ? disbursementsMatch[1]
      : 0;
    data["Total HST On Disbursements"] = hstMatch ? hstMatch[1] : 0;
    data.Total = totalMatch ? totalMatch[1] : 0;
    data["Ar Balance"] = arBalanceMatch ? arBalanceMatch[1] : 0;
    data.Link = `${process.env.FRONTEND_URL}?pdf=${pdfName}&page=${
      page * 2 + 1
    }`;

    records.push(data);
  });

  return records;
};

app.post("/upload-xlFile", upload.single("xlFile"), async (req, res) => {
  try {
    const { file } = req;
    const { pdfFile } = req.body;
    const xlFileUrl = path.join(__dirname, "public", file.filename);
    const pdfFileUrl = path.join(__dirname, "public", pdfFile);

    const pdfBuffer = fs.readFileSync(pdfFileUrl);

    const data = await pdfParse(pdfBuffer);
    const text = data.text;
    const combinedRecords = combinePages(text, 2);
    const JSONData = extractData(combinedRecords, pdfFile);

    const xlsxJSONData = readExcelFile(xlFileUrl);

    xlsxJSONData.forEach((record) => {
      const file = record.File;
      const link = JSONData.find((data) => data.File === file)?.Link;
      record.Link = link;
    });

    // // Convert updated JSON back to XLS
    const xls = json2xls(xlsxJSONData);
    const xlsxName = `${pdfFile}-updated.xlsx`;
    const xlsxPath = path.join(__dirname, "public", xlsxName);
    fs.writeFileSync(xlsxPath, xls, "binary");

    res.json({ file: file, data: JSONData, xlsxURL: xlsxName });
  } catch (error) {
    console.error("Error uploading file:", error);
    res.status(500).send("Failed to upload file");
  }
});

app.post("/upload-pdf", upload.single("pdfFile"), (req, res) => {
  try {
    res.send({ message: "File uploaded successfully", file: req.file });
  } catch (error) {
    console.error("Error uploading file:", error);
    res.status(500).send("Failed to upload file");
  }
});

app.listen(PORT, () => {
  console.log("Server is running on port " + PORT);
});
