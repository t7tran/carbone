const carbone = require("../lib");
const path = require("path");
const fs = require("fs");
const helper = require("../lib/helper");

carbone.set({ lang: "en-us" });

const tempDir = path.join(__dirname, "temp");
const file = name=> path.join(__dirname, "datasets", name);
const options = {};

const writeResult = (file, result) => {
  const [name, ext] = file.split(".");
  fs.writeFileSync(path.join(tempDir, `${name}_${new Date().toISOString().replaceAll(/[TZ:.-]/g, "")}.${ext}`), result);
};

describe("Extra Tests", function () {
  describe("render", function () {
    describe("HTML to PDF with syntax error", function () {
      it("should render successfully", function (done) {
        carbone.render(
          file("test_syntax.html"),
          file("test_sample.json"),
          { ...options, convertTo: "PDF" },
          (err, result) => {
            helper.assert(!err, true);
            writeResult("test_syntax.pdf", result);
            done(err);
          }
        );
      });
    });
  });
});

