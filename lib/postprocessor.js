const crypto = require("crypto");
const axios = require("axios");

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
      case "xlsx":
      case "ods":
      default:
        return callback();
    }
    processor.processImages((err) => callback(err));
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
}

const alldrawings = /<w:drawing>(.*?)<\/w:drawing>/g;
const allrels = /<Relationship (.*?)\/>/g;

class DocxPostProcessor extends Processor {
  constructor(template) {
    super(template);
  }
  processImages(callback) {
    const documentXmlFile = this.template.files.find(
      (f) => f.name === "word/document.xml"
    );
    if (!documentXmlFile) return callback();
    const documentXmlRelsFile = this.template.files.find(
      (f) => f.name === "word/_rels/document.xml.rels"
    );

    const drawingUrls = (documentXmlFile.data.match(alldrawings) || [])
      .map((drawing) => {
        const [, , url] =
          /(title|descr)="(https:\/\/(.*?))"/.exec(drawing) || [];
        return url;
      })
      .filter((url) => !!url);

    fetchImages(drawingUrls)
      .then((resolved) => {
        this.filestore.imagesFetched(resolved.map((r) => r.value));
        documentXmlFile.data = documentXmlFile.data.replaceAll(
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
            const imgFile = `media/${basename}.${extension}`;
            if (newfile) {
              this.template.files.push({
                name: imgFile,
                isMarked: false,
                data: Buffer.from(content, "base64"),
                parent: "",
              });
              // Update corresponding entry in word/_rels/document.xml.rels file
              documentXmlRelsFile.data = documentXmlRelsFile.data.replaceAll(
                allrels,
                function (relationship) {
                  const [, id] = /Id="(.*?)"/.exec(relationship) || [];
                  if (id != relationshipId) return relationship;
                  return relationship.replace(
                    /Target=".*?"/g,
                    `Target="/${imgFile}"`
                  );
                }
              );
            }
            return drawing.replace(
              /(title|descr)="(data:[^;]+;base64,.*?|https:\/\/.*?)"/g,
              '$1=""'
            );
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
}

module.exports = postprocessor;
