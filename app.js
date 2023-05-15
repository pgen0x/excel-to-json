const express = require("express");
const app = express();
const bodyParser = require("body-parser");
const multer = require("multer");
const xlsToJson = require("xls-to-json-lc");
const xlsxToJson = require("xlsx-to-json-lc");

app.use(bodyParser.json());

// Multer configuration untuk menentukan tempat penyimpanan dan format nama file
const storage = multer.diskStorage({
  destination: "./uploads/",
  filename: function (req, file, cb) {
    const datetimestamp = Date.now();
    cb(
      null,
      `${file.fieldname}-${datetimestamp}.${
        file.originalname.split(".")[file.originalname.split(".").length - 1]
      }`
    );
  },
});

// Middleware multer untuk filter ekstensi file
const upload = multer({
  storage: storage,
  fileFilter: function (req, file, callback) {
    const fileExtension =
      file.originalname.split(".")[file.originalname.split(".").length - 1];
    if (["xls", "xlsx"].indexOf(fileExtension) === -1) {
      return callback(
        new Error("Ekstensi file salah, hanya upload file xls atau xlsx")
      );
    }
    callback(null, true);
  },
}).single("file"); // hanya satu file yang diunggah dan nama fieldnya adalah "file"

// Endpoint untuk mengupload file
app.post("/upload", function (req, res) {
  upload(req, res, function (err) {
    if (err) {
      console.log(err);
      return res.json({ error_code: 1, err_desc: err });
    }
    if (!req.file) {
      return res.json({
        error_code: 1,
        err_desc: "Tidak ada file yang di unggah",
      });
    }
    const fileExtension =
      req.file.originalname.split(".")[
        req.file.originalname.split(".").length - 1
      ];
    const excelToJson = fileExtension === "xlsx" ? xlsxToJson : xlsToJson;

    console.log(req.file.path);
    try {
      excelToJson(
        {
          input: req.file.path,
          output: null, // jika ingin menjadikan format .json contoh: output.json
          lowerCaseHeaders: true,
        },
        function (err, result) {
          console.log(result);
          if (err) {
            return res.json({ error_code: 1, err_desc: err, data: null });
          }
          res.json({ error_code: 0, err_desc: null, data: result });
        }
      );
    } catch (e) {
      res.json({ error_code: 1, err_desc: "File excel tidak dapat dibaca" });
    }
  });
});

// Endpoint untuk halaman utama
app.get("/", function (req, res) {
  res.sendFile(__dirname + "/index.html");
});

const PORT = 3000;
app.listen(PORT, function () {
  console.log(`Listening on port ${PORT}`);
});
