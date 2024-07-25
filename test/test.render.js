const carbone = require("../lib");
const path = require("path");
const fs = require("fs");
const helper = require("../lib/helper");
const spawn = require("child_process").spawn;

carbone.set({ lang: "en-us" });

const tempDir = path.join(__dirname, "temp");
const templateFile = path.join(__dirname, "datasets", "test_sample.docx");
const dataFile = path.join(__dirname, "datasets", "test_sample.json");
const resultFile = path.join(
  tempDir,
  `test_sample_${new Date().toISOString().replaceAll(/[TZ:.-]/g, "")}`
);
const data = JSON.parse(fs.readFileSync(dataFile, "utf8"));
const options = {};

if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir);

describe("Carbone Render Tests", function () {
  describe("render", function () {
    describe("Docx to Docx", function () {
      after(function (done) {
        helper.rmDirRecursive(`${resultFile}_zip`);
        done();
      });
      it("should render successfully", function (done) {
        carbone.render(templateFile, data, options, (err, result) => {
          if (err) {
            console.error(err);
            done();
            return;
          }
          fs.writeFileSync(`${resultFile}.docx`, result);
          extractDocumentXml(done);
        });
      });
    });
    describe("Docx to PDF", function () {
      it("should render successfully", function (done) {
        carbone.render(
          templateFile,
          data,
          { ...options, convertTo: "PDF" },
          (err, result) => {
            if (err) console.error(err);
            else fs.writeFileSync(`${resultFile}.pdf`, result);
            done();
          }
        );
      });
    });
  });
});

function extractDocumentXml(done) {
  unzipSystem(`${resultFile}.docx`, `${resultFile}_zip`, function (err, files) {
    fs.writeFileSync(`${resultFile}.xml`, files["word/document.xml"]);
    done();
  });
}

function unzipSystem(filePath, destPath, callback) {
  var _unzippedFiles = {};
  var _unzip = spawn("unzip", ["-o", filePath, "-d", destPath]);
  _unzip.on("error", function () {
    throw Error(
      "\n\nPlease install unzip program to execute tests. Ex: sudo apt install unzip\n\n"
    );
  });
  _unzip.stderr.on("data", function (data) {
    throw Error(data);
  });
  _unzip.on("exit", function () {
    var _filesToParse = helper.walkDirSync(destPath);
    for (var i = 0; i < _filesToParse.length; i++) {
      var _file = _filesToParse[i];
      var _content = fs.readFileSync(_file, "utf8");
      _unzippedFiles[path.relative(destPath, _file)] = _content;
    }
    callback(null, _unzippedFiles);
  });
}
