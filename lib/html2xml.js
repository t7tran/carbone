// credit: https://github.com/CryptoNinjaGeek/carbone.git
const html2json = require('html2json').html2json;
const htmlEntities = require('./htmlentities');

const TAGS = {
  docx : {
    start_text : '<w:t>',
    end_text : '</w:t>',
    b : '<w:b/>',
    u : '<w:u/>',
    i : '<w:i/>',
    color : (hex) => { return `<w:color w:val="${hex}"/>` },
  }
}
// w:sz in half-points
const PT_CONV = {
  px: 3/4 * 2, // 16px = 12pt
  pt: 2
}
const NEW_PARAGRAPH = "<w:p>";
const TEXT_ALIGNMENT = {
  right: "right",
  left: "left",
  center: "center",
  justify: "both",
};

let listId = 1
const headings = ["h1", "h2", "h3", "h4", "h5", "h6"];
let lastNumId = 1000;

const inheritStyles = (styles) => {
  if (!styles) return '';

  const [fonts] = styles.match(/<w:rFonts .*?>/) || [];
  const [lang] = styles.match(/<w:lang .*?>/) || [];
  return `${fonts||''}${lang||''}`;
};

class html2xml {

  constructor(html, styles) {
    this.html = html
    this.validStyles = (styles || {}).validStyles || []
    this.pstyle = (styles || {}).pstyle
    this.pstyles = inheritStyles((styles || {}).pstyles)
    this.tstyles = inheritStyles((styles || {}).tstyles)
    this.result = []
    this.TAGS = TAGS
    this.hasError = false
    this.json = {}
    this.iter = 0
    this.hasTable = false;
  }

  toJSON() {

    try {
      this.html = htmlEntities.decode(this.html);
      // XXX commented next 2 lines as it has been taken care of by sanitisation/minifying process in the formatter
      // this.html = this.html.replace(/\n/g, "")
      // this.html = this.html.replace(/\t/g, "")
      this.json = html2json(this.html);
    }
    catch(error) {
      this.hasError = true
    }

  }

  getXML() {

    this.toJSON()

    const temp = ["<w:p>"]
    if(!this.hasError && this.json) {
      this.convert(this.json)
      const lastResult = this.result.pop();
      if (lastResult.endsWith(NEW_PARAGRAPH))
        this.result.push(lastResult.substring(0, lastResult.length - NEW_PARAGRAPH.length));
      else {
        this.result.push(lastResult);
        this.result.push("</w:p>");
      }
      if (this.hasTable)
        // XXX remove empty paragraph before tables
        temp.push(this.result.join("").replaceAll("<w:p><!--table--></w:p>", ""));
      else
        temp.push(this.result.join(""));
    }
    else {
      const html = this.html.replace(/<[\s\S]*?>/gi, "")
      temp.push(`<w:r><w:t>${html}</w:t></w:r></w:p>`)
    }

    const xml = temp.join("")

    return xml;

  }

  buildStyles(styles) {
    const temp = Array.isArray(styles) ? styles : typeof styles === "string" ? [styles] : []
    const _styles = temp.join("").split(";")
    const props = {}
    _styles.map( style => {
      if( typeof style === "string" && style.length > 0) {
        var arg_val = style.split(':')
        if( Array.isArray(arg_val) && arg_val.length === 2) {
          var arg = arg_val[0]
          var val = arg_val[1]
          props[arg] = val
        }
      }
    })
    return props
  }

  getLastNumId() {
    return lastNumId;
  }

  issueNumId() {
    return ++lastNumId;
  }

  isEmptyParagraph(element) {
    if(element.node === "text" && typeof element.text === "string") return false;
    if (!Array.isArray(element.child) || !element.child.length) return true;

    return element.child.every(c => this.isEmptyParagraph(c));
  }

  markNewParagraph(element) {
    if (!Array.isArray(element.child) || !element.child.length) return;

    const lastChild = element.child[element.child.length - 1];
    lastChild.properties = lastChild.properties || {};
    lastChild.properties.newParagraph = true;

    this.markNewParagraph(lastChild);
  }

  setProps(child, deep = 0) {

    if(child.tag === "p") {
      if (child.attr?.class) {
        const [first] = Array.isArray(child.attr.class) ? child.attr.class : [child.attr.class]
        if (this.validStyles.includes(first)) child.properties.class = first
      }
      if (this.isEmptyParagraph(child)) {
        child.child = [
          { node: "text", text: "", properties: { newParagraph: true } },
        ];
      } else if (Array.isArray(child.child) && child.child.length > 0) {
        child.child[0].properties = child.child[0].properties || {};
        child.child[0].properties.firstItem = true;
        this.markNewParagraph(child);
      }
    }
    else if(child.tag === "table") {
      child.properties.hasBorder = child.attr?.border
      child.properties.cellpadding = child.attr?.cellpadding
      child.properties.cellspacing = child.attr?.cellspacing
      child.properties.tableStyle = child.attr?.style
      child.properties.table = true
    }
    else if(child.tag === "tbody") {
      child.properties.tbodyStyle = child.attr?.style
      child.properties.tbody = true
      child.child.map( (_child, i) => {
        _child.properties = _child.properties || {}
        _child.properties.row = i
        if(i === child.child.length-1) _child.properties.lastRow = true
      })
    }
    else if(child.tag === "tr") {
      child.properties.trStyle = child.attr?.style
      child.properties.numCols = child.child.length
      child.child.map( (_child, i) => {
        _child.properties = _child.properties || {}
        _child.properties.col = i
        if(i === child.child.length-1) _child.properties.lastCol = true
      })
    }
    else if(child.tag === "td") {
      child.properties.tdStyle = child.attr?.style
    }
    else if(headings.includes(child.tag)) {
      const level = headings.findIndex(h => h === child.tag) + 1;
      child.properties.headingStyle = `Heading${level}`
    }
    else if(child.tag === "strong" || child.tag === "b") {
      child.properties.bold = true
    }
    else if(child.tag === "em" || child.tag === "i") {
      child.properties.italic = true
    }
    else if(child.tag === "u") {
      child.properties.underline = true
    }
    else if(child.tag === "sub") {
      child.properties.subScript = true
    }
    else if(child.tag === "sup") {
      child.properties.supScript = true
    }
    else if(child.tag === "s") {
      child.properties.strike = true
    }
    else if(child.tag === "span") {

    }
    else if(child.tag === "ul") {
      child.properties.listType = "bullet"
      child.properties.listId = child.properties.listId || listId++
    }
    else if(child.tag === "ol") {
      child.properties.listType = "numbering"
      child.properties.listId = child.properties.listId || listId++
      child.properties.numId = this.issueNumId();
    }
    else if(child.tag === "li") {
      child.properties.list = true
      child.properties.newParagraph = true
      child.properties.deep = child.parent.properties.deep + 1 || deep

      if(child.properties.listType === "bullet") {
        child.properties.listTypeXML = "Paragraphedeliste"
        child.properties.numId = 1000;
      }
      else if(child.properties.listType === "numbering") {
        child.properties.listTypeXML = "ListParagraph"
        child.properties.numId = child.parent.properties.numId;
      }
      else {

      }
    }
  }

  buildText(child) {

    const props = child.properties || {}
    var text = []
    text.push(`<w:r>`)

    if(props) {

      if(props.list) {

        const list = `
          <w:pPr>
            <w:pStyle w:val="${this.pstyle || 'Normal' || props.listTypeXML}"/>
            <w:numPr>
              <w:ilvl w:val="${props.deep}"/>
              <w:numId w:val="${props.numId || 1}"/>
            </w:numPr>
            <w:rPr></w:rPr>
          </w:pPr>
        `
        text.pop();
        text.push(list);
        text.push(`<w:r>`);

      }

      if (props.headingStyle) {
        props.newParagraph = true
        const heading = `
          <w:pPr>
            <w:pStyle w:val="${props.headingStyle}"/>
            <w:numPr>
              <w:ilvl w:val="0"/>
              <w:numId w:val="0"/>
            </w:numPr>
            <w:ind w:left="0" w:hanging="0"/>
            <w:rPr></w:rPr>
          </w:pPr>
        `
        text.pop();
        text.push(heading)
        text.push(`<w:r>`)
      }

      text.push('<w:rPr>')
      text.push(this.tstyles);
      if(props.strike) text.push('<w:strike/>')
      else if(props.dstrike) text.push('<w:dstrike w:val="true" />')
      if(props.italic) text.push('<w:i/><w:iCs/>');
      if(props.bold) text.push('<w:b/><w:bCs/>');
      if(props.underline) text.push('<w:u w:val="single"/>')

      props.formatting_options = { style: props.class };
      if(props.style) {

        let styles = this.buildStyles(props.style)

        for (var arg in styles) {
          if (styles.hasOwnProperty(arg)) {
            const val = styles[arg]
            if(arg === "color") {
              var rgb = val.startsWith("#") ? val.substring(1) : null
              if (rgb !== null) {
                text.push(`<w:color w:val="${rgb.toUpperCase()}"/>`)
              }
            }
            else if(arg === "background-color") {
              var rgb = val.startsWith("#") ? val.substring(1) : null
              if (rgb !== null) {
                text.push(`<w:shd w:val="clear" w:color="33FF49" w:fill="${rgb.toUpperCase()}"/>`)
              }
            }
            if (arg === "font-size") {
              const [, size, unit] = val.match(/^(\d+\.?\d+?)(px|pt)$/) || [];
              if (size) {
                const factor = PT_CONV[(unit || "pt").toLowerCase()];
                const _size = +size * factor;
                text.push(`<w:sz w:val="${_size}"/><w:szCs w:val="${_size}"/>`);
              }
            }
            let align;
            if (arg === "text-align" && (align = TEXT_ALIGNMENT[val])) {
              props.formatting_options = { ...props.formatting_options, align };
            }
            if (arg === "text-indent") {
              const [, pc] = val.match(/^(\d+(.\d+)?)%$/) || [];
              if (pc) props.formatting_options = { ...props.formatting_options, indent : Math.round(72 * parseFloat(pc)) };
            }
            // 12pt = 240 unit
            if (arg === "margin-top") {
              const [, pt] = val.match(/^(\d+(.\d+)?)pt$/) || [];
              if (pt) props.formatting_options = { ...props.formatting_options, spacingBefore : Math.round(240 * parseFloat(pt) / 12) };
            }
            if (arg === "margin-bottom") {
              const [, pt] = val.match(/^(\d+(.\d+)?)pt$/) || [];
              if (pt) props.formatting_options = { ...props.formatting_options, spacingAfter : Math.round(240 * parseFloat(pt) / 12) };
            }
            if (arg === "line-height") {
              const [, pt] = val.match(/^(\d+(.\d+)?)pt$/) || [];
              if (pt) props.formatting_options = { ...props.formatting_options, spacingLine : Math.round(240 * parseFloat(pt) / 12) };
            }
          }
        }

      }
      text.push('</w:rPr>')

      text.push(`<w:t xml:space="preserve">${child.text}</w:t>`)

      if (props.formatting_options) {
        const fopts = [];
        fopts.push(`<w:pPr>`);
        fopts.push(`<w:pStyle w:val="${props.formatting_options.style || this.pstyle || "Normal"}"/>`);
        if ("spacingBefore" in props.formatting_options || "spacingAfter" in props.formatting_options || "spacingLine" in props.formatting_options) {
          let spacing = '';
          if ("spacingLine"   in props.formatting_options) spacing += ` w:line="${props.formatting_options.spacingLine}"`;
          if ("spacingBefore" in props.formatting_options) spacing += ` w:before="${props.formatting_options.spacingBefore}"`;
          if ("spacingAfter"  in props.formatting_options) spacing += ` w:after="${props.formatting_options.spacingBefore}"`;
          fopts.push(`<w:spacing ${spacing}/>`);
        }
        if (props.formatting_options.indent)
          fopts.push(`<w:ind w:left="${props.formatting_options.indent}" w:hanging="0"/>`);
        if (props.formatting_options.align)
          fopts.push(`<w:jc w:val="${props.formatting_options.align}"/>`);
        fopts.push(`<w:rPr>${this.pstyles}</w:rPr>`);
        fopts.push(`</w:pPr>`);

        for (let fopt of fopts.reverse()) text.unshift(fopt);
      }
      text.push(`</w:r>`)

      if(props.newParagraph) {
        text.push(`</w:p>`)
        text.push(NEW_PARAGRAPH);
      }

      /// Si jamais c'est dans un tableau
      if(props.table) {

        this.hasTable = true;
        const tdStyles = this.buildStyles(props.tdStyle)
        const tableStyles = this.buildStyles(props.tableStyle)
        const table_total_width = 9000; // 100% de la largeur

        // const table_width = tableStyles.width.endsWith("px") ? tableStyles.width.substring(-2) :
        // const column_width = tdStyles.width.endsWith("px") ? tdStyles.width.substring(-2) :

          if(props.col === 0 && props.row === 0) {
            const wsz = 8;
            this.result.push("<!--table--></w:p>\n<w:tbl><w:tblPr>")
            if(props.hasBorder) {
              this.result.push(`
                <w:tblBorders>
                  <w:top w:val="single" w:sz="${wsz}" w:space="0" w:color="000000" />
                  <w:start w:val="single" w:sz="${wsz}" w:space="0" w:color="000000" />
                  <w:bottom w:val="single" w:sz="${wsz}" w:space="0" w:color="000000" />
                  <w:end w:val="single" w:sz="${wsz}" w:space="0" w:color="000000" />
                  <w:insideH w:val="single" w:sz="${wsz-2}" w:space="0" w:color="000000" />
                  <w:insideV w:val="single" w:sz="${wsz-2}" w:space="0" w:color="000000" />
                </w:tblBorders>
              `)
            }
            this.result.push(`
              <w:tblW w:w="0" w:type="auto"/>
              </w:tblPr>
            \n\t<w:tr>\n\t\t<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr>\n\t\t\t<w:p>`)
          }
          else if(props.col > 0) this.result.push(`</w:p>\n\t\t</w:tc>\n\t\t<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr>\n\t\t\t<w:p>`)

          this.result.push(text.join(""))

          if(props.lastCol && props.lastRow) this.result.push(`</w:p>\n\t\t</w:tc>\n\t</w:tr>\n</w:tbl>\n<w:p>`)
          else if(props.lastCol) this.result.push(`</w:p>\n\t\t</w:tc>\n\t</w:tr>\n\t<w:tr>\n\t\t<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr>\n\t\t\t<w:p>`)

      }
      else {
        this.result.push(text.join(""))
      }

    }

  }

  convert(item) {
    const tableTags = ["table", "tbody", "tr", "td", "thead", "th", "tfoot"]

    if(Array.isArray(item.child) && item.child.length>0) {
      item.child.map( (child, i) => {
        child.parent = item
        const parentProps = { ...child.parent.properties };
        delete parentProps.newParagraph;
        child.properties = Object.assign(child.properties || {}, parentProps)

        if(child.attr?.style && tableTags.indexOf(child.tag) < 0 ) {
          child.properties.style = child.attr.style
        }

        if(child.node === "element") {
          this.setProps(child)
          this.convert(child)
        }

        else if(child.node === "text" && typeof child.text === "string") {
          this.buildText(child)
        }
      })
    }
  }
}





module.exports = html2xml
