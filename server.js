const multer = require("multer");
const XLSX = require("xlsx");
const express = require("express");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");
const mime = require("mime-types");

const app = express();
app.use(cors());
app.use(express.json());

// Upload Folder
const upload = multer({
  dest: "uploads/"
});

const CREDENTIALS_PATH = path.join(__dirname, "client_secret.json");
const TOKEN_PATH = path.join(__dirname, "token.json");

const SCOPES = [
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/drive.file"
];
const SPREADSHEET_ID = "1E_cBuqFpccpPE-sDmEtTrhwFIVvobCh1G_BBiiycwEo";
const SHEET_NAME = "Sheet1";
async function getAuthClient() {

  const credentials = JSON.parse(fs.readFileSync(CREDENTIALS_PATH));
  const token = JSON.parse(fs.readFileSync(TOKEN_PATH));

  const keys = credentials.web || credentials.installed;

  const auth = new google.auth.OAuth2(
    keys.client_id,
    keys.client_secret,
    keys.redirect_uris[0]
  );

  auth.setCredentials(token);

  return auth;
}

// Home Route
app.use(express.static(__dirname));

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// Google Authorization
app.get("/auth", (req, res) => {
  try {
    const credentials = JSON.parse(fs.readFileSync(CREDENTIALS_PATH));
    const keys = credentials.web || credentials.installed;

    const auth = new google.auth.OAuth2(
      keys.client_id,
      keys.client_secret,
      keys.redirect_uris[0]
    );

    const authUrl = auth.generateAuthUrl({
      access_type: "offline",
      scope: SCOPES
    });

    res.redirect(authUrl);

  } catch (err) {
    res.status(500).send(err.message);
  }
});

// OAuth Callback
app.get("/oauth2callback", async (req, res) => {
  try {

    const code = req.query.code;

    const credentials = JSON.parse(fs.readFileSync(CREDENTIALS_PATH));
    const keys = credentials.web || credentials.installed;

    const auth = new google.auth.OAuth2(
      keys.client_id,
      keys.client_secret,
      keys.redirect_uris[0]
    );

    const { tokens } = await auth.getToken(code);

    fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens));

    res.send("✅ token.json created successfully!");

  } catch (err) {
    res.status(500).send(err.message);
  }
});

// Upload test.txt to Google Drive
app.get("/upload", async (req, res) => {

  try {

    const credentials = JSON.parse(fs.readFileSync(CREDENTIALS_PATH));
    const token = JSON.parse(fs.readFileSync(TOKEN_PATH));

    const keys = credentials.web || credentials.installed;

    const auth = new google.auth.OAuth2(
      keys.client_id,
      keys.client_secret,
      keys.redirect_uris[0]
    );

    auth.setCredentials(token);

    const drive = google.drive({
      version: "v3",
      auth
    });

    const fileMetadata = {
      name: "test.txt"
    };

    const media = {
      mimeType: mime.lookup("test.txt"),
      body: fs.createReadStream(path.join(__dirname, "test.txt"))
    };

    const response = await drive.files.create({
      requestBody: fileMetadata,
      media: media
    });

    res.send("✅ Uploaded Successfully! File ID: " + response.data.id);

  } catch (err) {
    console.error(err);
    res.status(500).send(err.message);
  }

});

// Excel Upload Route
app.post("/uploadExcel", upload.single("file"), async (req, res) => {

  try {

    if (!req.file) {
      return res.status(400).send("No file uploaded");
    }

    const workbook = XLSX.readFile(req.file.path);

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const data = XLSX.utils.sheet_to_json(sheet);

    console.log("Excel Data:");
    console.log(data);

    res.json({
      success: true,
      totalRows: data.length,
      data: data
    });

  } catch (err) {
    console.error(err);
    res.status(500).send(err.message);
  }

});
app.post("/api/leads", async (req, res) => {
  console.log("Request received");
console.log(req.body);

  try {

    const { fullName, mobileNumber, loanAmount, loanType } = req.body;

    if (!fullName || !mobileNumber || !loanAmount || !loanType) {
      return res.status(400).json({
        success: false,
        message: "All fields are required"
      });
    }

    const auth = await getAuthClient();

    const sheets = google.sheets({
      version: "v4",
      auth
    });

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A:E`,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[
          new Date().toLocaleString("en-IN"),
          fullName,
          mobileNumber,
          loanAmount,
          loanType
        ]]
      }
    });
    console.log("Google Sheet me data add ho gaya");

    res.json({
      success: true,
      message: "Lead saved successfully"
    });

  } catch (err) {

    console.error(err);

    res.status(500).json({
      success: false,
      message: err.message
    });

  }

});
const PORT = 5000;

app.listen(PORT, () => {

  console.log("=================================");

  console.log("👉 http://localhost:5000/auth");
  console.log("👉 http://localhost:5000/upload");
  console.log("👉 POST http://localhost:5000/uploadExcel");
  console.log("👉 POST http://localhost:5000/api/leads");
  console.log("=================================");

});