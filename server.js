const express = require("express");
const bodyParser = require("body-parser");
const ExcelJS = require("exceljs");
const fs = require("fs");

const app = express();
app.use(bodyParser.json());

const file = "orders.xlsx";

// Create Excel file if not exists
if (!fs.existsSync(file)) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Orders");
  sheet.columns = [
    { header: "Name", key: "name" },
    { header: "Class", key: "class" },
    { header: "Item", key: "item" },
    { header: "Quantity", key: "quantity" },
    { header: "Payment", key: "payment" },
    { header: "Time", key: "time" },
    { header: "Date", key: "date" },
  ];
  workbook.xlsx.writeFile(file);
}

// Place order API
app.post("/place-order", async (req, res) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(file);
  const sheet = workbook.getWorksheet("Orders");

  sheet.addRow(req.body);
  await workbook.xlsx.writeFile(file);

  res.json({ success: true });
});

// Admin download Excel (password protected)
app.get("/admin/:password", (req, res) => {
  if (req.params.password !== "youradminpassword") {
    return res.status(403).send("Forbidden");
  }
  res.download(file);
});

app.listen(3000, () => console.log("Server running on port 3000"));
