const carbone = require("../lib");
const path = require("path");
const fs = require("fs");

carbone.set({ lang: "en-us" });

const tempDir = path.join(__dirname, "temp");
const templateFile = path.join(__dirname, "datasets", "test_sample.docx");
const dataFile = path.join(__dirname, "datasets", "test_sample.json");
const resultFile = path.join(
  tempDir,
  `test_sample_${new Date().toISOString().replaceAll(/[TZ:.-]/g, "")}.pdf`
);
const data = JSON.parse(fs.readFileSync(dataFile, "utf8"));
const options = {
  convertTo: "PDF",
};

if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir);

describe("Carbone Render Tests", function () {
  describe("render", function () {
    describe("Docx to PDF", function () {
      it("should render successfully", function (done) {
        carbone.render(templateFile, data, options, (err, result) => {
          if (err) console.error(err);
          else fs.writeFileSync(resultFile, result);
          done();
        });
      });
    });
  });
});
