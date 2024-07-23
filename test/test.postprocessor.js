var postprocessor = require('../lib/postprocessor');
var helper = require('../lib/helper');

const PNG_BASE64 = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAAXNSR0IArs4c6QAAAA1JREFUGFdj+L+U4T8ABu8CpCYJ1DQAAAAASUVORK5CYII=';

describe('postprocessor', function () {
  describe('process', function () {
    describe('ODT postprocessing', function () {
      var _xml = `<xml>
          <office:body>
            <office:text text:use-soft-page-breaks="true">
              <table:table-cell office:value-type="string" table:style-name="Table6.A2">
                <text:p text:style-name="P19">bla
                  <text:span text:style-name="T134">nam</text:span>
                  <text:soft-page-break/>
                  </text:p>
              </table:table-cell>
              <table:table-cell office:value-type="string" table:style-name="Table6.A2">
                <text:p text:style-name="P29">
                  <text:span text:style-name="T42">sd</text:span>
                  <text:soft-page-break></text:soft-page-break>
                  <text:span text:style-name="T43">position</text:span>
                  <text:soft-page-break/>
                  <text:span text:style-name="T42"></text:span>
                </text:p>
              </table:table-cell>
              <draw:frame draw:style-name="fr2" draw:name="Logo" text:anchor-type="page" text:anchor-page-number="1" svg:x="2.073cm" svg:y="2.445cm" svg:width="2.482cm" svg:height="2.482cm" draw:z-index="0"><draw:image xlink:href="Pictures/10000001000002000000020001BE103637D30330.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad" draw:mime-type="image/png"/><svg:title>__DATA__</svg:title></draw:frame>
            </office:text>
          </office:body>
        </xml>`
      ;
      it('should replace draw:frame with base64 data', function (done) {
          var contentXml = { name : 'content.xml', parent : '', data : _xml.replace('__DATA__', PNG_BASE64)};
          var _report = {
              isZipped   : true,
              filename   : 'template.odt',
          embeddings : [],
          extension  : 'odt',
          files      : [
            contentXml,
            { name : 'META-INF/manifest.xml', parent : '', data : ''}
          ]
        };
        postprocessor.process(_report, {}, {});
        helper.assert(/draw:mime-type="image\/png"/.test(contentXml.data), true);
        helper.assert(/xlink:href="Pictures\/[0-9a-f]+.png"/.test(contentXml.data), true);
        helper.assert(/<svg:title><\/svg:title>/.test(contentXml.data), true);
        helper.assert(/Pictures\/[0-9a-f]+.png/.test(_report.files[2].name), true);
        helper.assert(`data:image/png;base64,${_report.files[2].data.toString("base64")}`, PNG_BASE64);
        done();
      });
    });
    describe('DOCX postprocessing', function () {
      var _xml = `<xml>
          <office:body>
            <office:text text:use-soft-page-breaks="true">
              <table:table-cell office:value-type="string" table:style-name="Table6.A2">
                <text:p text:style-name="P19">bla
                  <text:span text:style-name="T134">nam</text:span>
                  <text:soft-page-break/>
                  </text:p>
              </table:table-cell>
              <table:table-cell office:value-type="string" table:style-name="Table6.A2">
                <text:p text:style-name="P29">
                  <text:span text:style-name="T42">sd</text:span>
                  <text:soft-page-break></text:soft-page-break>
                  <text:span text:style-name="T43">position</text:span>
                  <text:soft-page-break/>
                  <text:span text:style-name="T42"></text:span>
                </text:p>
              </table:table-cell>
              <w:drawing><wp:inline wp14:editId="1075CE7C" wp14:anchorId="207D284F"><wp:extent cx="1009650" cy="1009650" /><wp:effectExtent l="0" t="0" r="0" b="0" /><wp:docPr id="1771682389" name="" title="__DATA__" /><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1" /></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="" /><pic:cNvPicPr /></pic:nvPicPr><pic:blipFill><a:blip r:embed="R1b736920b34e4e63"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a14:useLocalDpi val="0" /></a:ext></a:extLst></a:blip><a:stretch><a:fillRect /></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0" /><a:ext cx="1009650" cy="1009650" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>
            </office:text>
          </office:body>
        </xml>`
      ;
      it('should replace w:drawing with base64 data', function (done) {
          var documentXml = { name : 'word/document.xml', parent : '', data : _xml.replace('__DATA__', PNG_BASE64)};
          var _report = {
              isZipped   : true,
              filename   : 'template.docx',
          embeddings : [],
          extension  : 'docx',
          files      : [
            documentXml,
            { name : 'word/_rels/document.xml.rels', parent : '', data : ''}
          ]
        };
        postprocessor.process(_report, {}, {});
        helper.assert(/title=""/.test(documentXml.data), true);
        helper.assert(/media\/[0-9a-f]+.png/.test(_report.files[2].name), true);
        helper.assert(`data:image/png;base64,${_report.files[2].data.toString("base64")}`, PNG_BASE64);
        done();
      });
    });
  });
});

