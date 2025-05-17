const express = require('express');
const fs = require('fs');
const xlsx = require('xlsx');
const bodyParser = require('body-parser');
const app = express();
const PORT = 3000;

app.use(bodyParser.json());
app.use(express.static(__dirname)); // Serve index.html

const EXCEL_FILE = 'submissions.xlsx';

app.post('/submit', (req, res) => {
  const data = req.body;

  let workbook;
  let worksheet;

  if (fs.existsSync(EXCEL_FILE)) {
    workbook = xlsx.readFile(EXCEL_FILE);
    worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = xlsx.utils.sheet_to_json(worksheet);
    jsonData.push(data);
    const newSheet = xlsx.utils.json_to_sheet(jsonData);
    workbook.Sheets[workbook.SheetNames[0]] = newSheet;
  } else {
    const newSheet = xlsx.utils.json_to_sheet([data]);
    workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, newSheet, 'Submissions');
  }

  xlsx.writeFile(workbook, EXCEL_FILE);
  res.sendStatus(200);
});

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
