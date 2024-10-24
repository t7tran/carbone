const crypto = require("crypto");
const axios = require("axios");
const path = require("path");
const fs = require("fs");
const html2xml = require("./html2xml");

const docxAssets = path.resolve(__dirname, "..", "assets", "docx");

const postprocessor = {
  process: function (template, _data, _options, callback) {
    var processor;
    switch (template.extension) {
      case "odt":
        processor = new OdtPostProcessor(template);
        break;
      case "docx":
        processor = new DocxPostProcessor(template);
        break;
      case "html":
        processor = new HtmlPostProcessor(template);
        break;
      case "xlsx":
      case "ods":
      default:
        return callback();
    }
    processor.processImages((err) => {
      if (err) return callback(err);
      processor.processHTML((err) => callback(err));
    });
  },
};

class FileStore {
  cache = [];
  images = {};
  fetchErrorMessage = undefined;
  constructor() {}

  fileBaseName(base64) {
    const hash = crypto.createHash("sha256").update(base64).digest("hex");
    const newfile = !this.cache.includes(hash);
    if (newfile) this.cache.push(hash);
    return [hash, newfile];
  }

  imagesFetched(fetchedImages) {
    this.images = fetchedImages
      .filter((i) => !!i)
      .reduce((m, i) => {
        m[i.url] = i;
        return m;
      }, {});
    this.fetchErrorMessage = fetchedImages.find((i) => i && i.error)?.error;
  }

  getImage(data) {
    const [, , mime, content, url] =
      /(data:([^;]+);base64,(.*)|(https:\/\/.*))/.exec(data) || [];
    if (mime && content) return { mime, content };
    if (!url) return {};
    return this.images[url] || {};
  }
}

function fetchImages(urls) {
  if (!urls || !urls.length) return Promise.resolve([]);
  const fetches = [...new Set(urls || [])]
    .map((url) => ({ url, fixed_url: url.replace(/&amp;/g, "&") }))
    .map(({ url, fixed_url }) =>
      axios({
        method: "get",
        url: fixed_url,
        responseType: "arraybuffer",
        timeout: 60000,
        maxContentLength:
          Number(process.env.CARBONE_MAX_IMAGE_URL || 10485760) || 10485760, // default 10MiB
      })
        .then(({ data, headers }) =>
          (headers["content-type"] || "").startsWith("image/")
            ? {
                url,
                mime: headers["content-type"],
                // FIXME need to scan for viruses
                content: Buffer.from(data, "base64"),
              }
            : {
                url,
                error: "Unrecognisable image content type",
              }
        )
        .catch((cause) => ({
          url,
          error: cause.message || "Image fetching error",
          cause,
        }))
    );
  return Promise.allSettled(fetches);
}

const allframes = /<draw:frame (.*?)<\/draw:frame>/g;

class Processor {
  filestore = new FileStore();
  constructor(template) {
    this.template = template;
  }
}

class OdtPostProcessor extends Processor {
  constructor(template) {
    super(template);
  }
  processImages(callback) {
    const contentXml = this.template.files.find(
      (f) => f.name === "content.xml"
    );
    if (!contentXml) return callback();
    const manifestXml = this.template.files.find(
      (f) => f.name === "META-INF/manifest.xml"
    );

    const frameUrls = (contentXml.data.match(allframes) || [])
      .map((drawFrame) => {
        const [, url] =
          /<svg:title>(https:\/\/(.*?))<\/svg:title>/.exec(drawFrame) || [];
        return url;
      })
      .filter((url) => !!url);

    fetchImages(frameUrls)
      .then((resolved) => {
        this.filestore.imagesFetched(resolved.map((r) => r.value));
        // Use base64 data to create new file and update references
        contentXml.data = contentXml.data.replaceAll(allframes, (drawFrame) => {
          const { mime, content } = this.filestore.getImage(
            (/<svg:title>(.*?)<\/svg:title>/.exec(drawFrame) || [])[1]
          );
          if (!content || !mime) return drawFrame;
          const [, extension] = mime.split("/");
          // Add new image to Pictures folder
          const [basename, newfile] = this.filestore.fileBaseName(content);
          const imgFile = `Pictures/${basename}.${extension}`;
          if (newfile) {
            this.template.files.push({
              name: imgFile,
              isMarked: false,
              data: Buffer.from(content, "base64"),
              parent: "",
            });
            // Update manifest.xml file
            manifestXml.data = manifestXml.data.replace(
              /((.|\n)*)(<\/manifest:manifest>)/,
              function (_match, p1, _p2, p3) {
                return [
                  p1,
                  `<manifest:file-entry manifest:full-path="${imgFile}" manifest:media-type="${mime}"/>`,
                  p3,
                ].join("");
              }
            );
          }
          return drawFrame
            .replace(/<svg:title>.*?<\/svg:title>/, "<svg:title></svg:title>")
            .replace(/draw:mime-type="[^"]+"/, `draw:mime-type="${mime}"`)
            .replace(/xlink:href="[^"]+"/, `xlink:href="${imgFile}"`);
        });
        return callback(
          this.filestore.fetchErrorMessage
            ? new Error(this.filestore.fetchErrorMessage)
            : undefined
        );
      })
      .catch((e) => {
        callback(new Error("Image processing error"));
      });
  }
  processHTML(callback) {
    // TODO implement support for ODT
  }
}

const alldrawings = /<w:drawing>(.*?)<\/w:drawing>/g;
const allparagraphs = /<w:p[^>]*\/>|<w:p.*?>.*?<\/w:p>/g;
const numberingXmlFilePath = "word/numbering.xml";
const allParagraphStyles = /<w:style w:type="paragraph" w:styleId="[^"]+"/g;
const headingCount = 6;

class DocxPostProcessor extends Processor {
  constructor(template) {
    super(template);
    this.documentXmlFile = this.template.files.find(
      (f) => f.name === "word/document.xml"
    );
    this.numberingXmlFile = this.template.files.find(
      (f) => f.name === numberingXmlFilePath
    );
    this.documentXmlRelsFile = this.template.files.find(
      (f) => f.name === "word/_rels/document.xml.rels"
    );
    this.stylesXmlFile = this.template.files.find(
      (f) => f.name === "word/styles.xml"
    );
    this.contentTypesXmlFile = this.template.files.find(
      (f) => f.name === "[Content_Types].xml"
    );

    this.paragraphStyleIds = new Array(headingCount)
      .fill("")
      .map((_, i) => `Heading${i + 1}`)
      .concat(
        (this.stylesXmlFile.data.match(allParagraphStyles) || []).map(
          (pstyle) => /w:styleId="([^"]+)"/.exec(pstyle)[1]
        )
      );
  }
  processImages(callback) {
    if (!this.documentXmlFile) return callback();

    const drawingUrls = (this.documentXmlFile.data.match(alldrawings) || [])
      .map((drawing) => {
        const [, , url] =
          /(title|descr)="(https:\/\/(.*?))"/.exec(drawing) || [];
        return url;
      })
      .filter((url) => !!url);

    fetchImages(drawingUrls)
      .then((resolved) => {
        this.filestore.imagesFetched(resolved.map((r) => r.value));
        this.documentXmlFile.data = this.documentXmlFile.data.replaceAll(
          alldrawings,
          (drawing) => {
            const { mime, content } = this.filestore.getImage(
              (/(title|descr)="(.*?)"/.exec(drawing) || [])[2]
            );
            const [, relationshipId] = /embed="(.*?)"/.exec(drawing) || [];
            if (!content || !mime || !relationshipId) return drawing;
            const [, extension] = mime.split("/");
            // Save image to media folder
            const [basename, newfile] = this.filestore.fileBaseName(content);
            const imgFile = `word/media/${basename}.${extension}`;
            if (newfile) {
              this.template.files.push({
                name: imgFile,
                isMarked: false,
                data: Buffer.from(content, "base64"),
                parent: "",
              });
              // Add new entry in word/_rels/document.xml.rels file
              this.documentXmlRelsFile.data =
                this.documentXmlRelsFile.data.replace(
                  "</Relationships>",
                  `<Relationship Id="${basename}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="/${imgFile}"/></Relationships>`
                );
            }
            return drawing
              .replace(
                /(title|descr)="(data:[^;]+;base64,.*?|https:\/\/.*?)"/g,
                '$1=""'
              )
              .replace(/embed="(.*?)"/g, `embed="${basename}"`);
          }
        );
        return callback(
          this.filestore.fetchErrorMessage
            ? new Error(this.filestore.fetchErrorMessage)
            : undefined
        );
      })
      .catch((e) => {
        callback(new Error("Image processing error"));
      });
  }
  processHTML(callback) {
    try {
      this._decodeBase64Html();
      this._processNumberingXml();
      this._processHeadingStyles();
    } catch (e) {
      return callback(e);
    }
    return callback();
  }
  _decodeBase64Html() {
    if (!this.documentXmlFile) return;

    this.documentXmlFile.data = this.documentXmlFile.data.replaceAll(
      allparagraphs,
      (paragraph) => {
        const [, base64] = paragraph.match(/<w:p.*?>.*<w:t>html:([A-Za-z0-9+\/]+=*):html<\/w:t>.*<\/w:p>/) || [];
        if (!base64) return paragraph;

        const [, pstyle] = paragraph.match(/<w:p.*?>.*?<w:pPr>.*?<w:pStyle w:val="([A-Za-z0-9_]+)"/) || [];
        const [, pstyles] = paragraph.match(/<w:p.*?>.*?<w:pPr>.*?<w:rPr>(.*?)<\/w:rPr>/) || [];
        const [, tstyles] = paragraph.match(/.*<w:rPr>(.*?)<\/w:rPr>.*?<w:t>html:[A-Za-z0-9+\/]+=*/) || [];

        const processedHtml = Buffer.from(base64, "base64").toString("utf8");
        return new html2xml(processedHtml, {pstyle, pstyles, tstyles, validStyles: this.paragraphStyleIds}).getXML();
      }
    );
  }
  _processNumberingXml() {
    const numberingXmlTemplate = fs.readFileSync(
      path.resolve(docxAssets, "numbering.xml"),
      "utf8"
    );
    const [abstractNum1001] =
      numberingXmlTemplate.match(
        /<w:abstractNum w:abstractNumId="1001".*?<\/w:abstractNum>/s
      ) || [];
    if (this.numberingXmlFile) {
      const [abstractNums] =
        numberingXmlTemplate.match(/<w:abstractNum.*<\/w:abstractNum>/s) || [];
      this.numberingXmlFile.data = this.numberingXmlFile.data.replace(
        "<w:abstractNum",
        abstractNums + "<w:abstractNum"
      );
    } else {
      this.numberingXmlFile = {
        name: numberingXmlFilePath,
        isMarked: false,
        data: numberingXmlTemplate,
        parent: "",
      };
      this.template.files.push(this.numberingXmlFile);
      this.contentTypesXmlFile.data = this.contentTypesXmlFile.data.replace(
        "<Override ",
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override '
      );
      this.documentXmlRelsFile.data = this.documentXmlRelsFile.data.replace(
        "</Relationships>",
        '<Relationship Id="rId1000" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'
      );
    }
    let additionalAbstractNums = "";
    let numStyles = "";
    const lastNumId = new html2xml().getLastNumId();
    for (let i = 1002; i <= lastNumId; i++)
      additionalAbstractNums += abstractNum1001.replace("1001", i);
    if (additionalAbstractNums)
      this.numberingXmlFile.data = this.numberingXmlFile.data.replace(
        "<w:abstractNum",
        additionalAbstractNums + "<w:abstractNum"
      );
    for (let i = 1000; i <= lastNumId; i++)
      numStyles += `<w:num w:numId="${i}"><w:abstractNumId w:val="${i}"/></w:num>`;
    this.numberingXmlFile.data = this.numberingXmlFile.data.replace(
      "</w:numbering>",
      numStyles + "</w:numbering>"
    );
  }
  _processHeadingStyles() {
    // FIXME Check if styles to be added.
    if (!this.stylesXmlFile) return;

    const headings = (level) => `
      <w:style w:type="paragraph" w:styleId="Heading${level}">
        <w:name w:val="Heading ${level}"/>
        <w:basedOn w:val="Heading"/>
        <w:next w:val="TextBody"/>
        <w:qFormat/>
        <w:pPr>
          <w:numPr>
            <w:ilvl w:val="0"/>
            <w:numId w:val="1"/>
          </w:numPr>
          <w:spacing w:before="240" w:after="120"/>
          <w:outlineLvl w:val="0"/>
        </w:pPr>
        <w:rPr>
          <w:b/>
          <w:bCs/>
          <w:sz w:val="${36 - 4 * level}"/>
          <w:szCs w:val="${36 - 4 * level}"/>
        </w:rPr>
      </w:style>
    `;
    const xml = new Array(headingCount)
      .fill("")
      .map((_, i) => headings(i + 1).trim())
      .join("");
    this.stylesXmlFile.data = this.stylesXmlFile.data.replace(
      /(.*)(<\/w:styles>)/,
      `$1${xml}$2`
    );
  }
}

class HtmlPostProcessor extends Processor {
  constructor(template) {
    super(template);
  }
  processImages(callback) {
    return callback();
  }
  processHTML(callback) {
    try {
      this._decodeBase64Html();
    } catch (e) {
      return callback(e);
    }
    return callback();
  }
  _decodeBase64Html() {
    const htmlFile = this.template.files[0];
    htmlFile.data = htmlFile.data.replaceAll(
      /html:([A-Za-z0-9+/]+=*):html/g,
      (_, base64html) => Buffer.from(base64html, "base64").toString("utf8")
    );
  }
}

module.exports = postprocessor;
