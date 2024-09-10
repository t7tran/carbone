const carbone = require("../lib");
const path = require("path");
const fs = require("fs");
const helper = require("../lib/helper");
const spawn = require("child_process").spawn;
const { comparePdfToSnapshot } = require("pdf-visual-diff");

carbone.set({ lang: "en-us" });

const tempDir = path.join(__dirname, "temp");
const docxTemplateFile = path.join(__dirname, "datasets", "test_sample.docx");
const htmlTemplateFile = path.join(__dirname, "datasets", "test_sample.html");
const dataFile = path.join(__dirname, "datasets", "test_sample.json");
const resultFile = path.join(
  tempDir,
  `test_sample_${new Date().toISOString().replaceAll(/[TZ:.-]/g, "")}`
);
const jsonData = fs
  .readFileSync(dataFile, "utf8")
  .replaceAll(/"__([a-zA-Z]+\.[a-z]+)__"/g, (_placeholder, part) =>
    JSON.stringify(
      fs.readFileSync(
        path.join(__dirname, "datasets", `test_sample_${part}`),
        "utf8"
      )
    )
  );
const data = JSON.parse(jsonData);
const options = {};

fs.writeFileSync(`${resultFile}.json`, jsonData);

if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir);

describe("Carbone Render Tests", function () {
  describe("render", function () {
    describe("Docx to Docx", function () {
      after(function (done) {
        helper.rmDirRecursive(`${resultFile}_zip`);
        done();
      });
      it("should render successfully", function (done) {
        carbone.render(docxTemplateFile, data, options, (err, result) => {
          if (err) {
            console.error(err);
            return done(err);
          }
          fs.writeFileSync(`${resultFile}.docx`, result);
          extractDocumentXml(done);
        });
      });
    });
    describe("Docx to PDF", function () {
      it("should render successfully", function (done) {
        carbone.render(
          docxTemplateFile,
          data,
          { ...options, convertTo: "PDF" },
          (err, result) => {
            if (err) console.error(err);
            else fs.writeFileSync(`${resultFile}.pdf`, result);
            done(err);
          }
        );
      });
      process.env.REMOTE_CONTAINERS && // XXX only verify inside test container
        it("should cause no visual differences", function (done) {
          comparePdfToSnapshot(
            `${resultFile}.pdf`,
            path.join(__dirname, "datasets", "test_sample"),
            "test_sample"
          )
            .then((x) => {
              helper.assert(x, true);
              done();
            })
            .catch((e) => done(e));
        });
    });
    describe("HTML to PDF", function () {
      it("should render successfully", function (done) {
        carbone.render(
          htmlTemplateFile,
          data,
          { ...options, convertTo: "PDF" },
          (err, result) => {
            if (err) console.error(err);
            else fs.writeFileSync(`${resultFile}.pdf`, result);
            done(err);
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
