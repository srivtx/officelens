/// <reference types="bun" />
import { describe, expect, test } from "bun:test";
import { strToU8, zipSync } from "fflate";
import { openOoxml } from "../src/package.ts";
import { auditDocx } from "../src/docx.ts";
import { audit } from "../src/audit.ts";
import { toSarif } from "../src/sarif.ts";

const enc = (value: string): Uint8Array => strToU8(value);

const RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

const NS =
  `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
  `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
  `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
  `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;

function buildDocx(documentXml: string, stylesXml: string): Uint8Array {
  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": enc(CONTENT_TYPES),
    "_rels/.rels": enc(RELS),
    "word/document.xml": enc(documentXml),
    "word/styles.xml": enc(stylesXml),
  };
  return zipSync(files);
}

const BAD_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}>
  <w:body>
    <w:p><w:r><w:t>Body text with no heading structure at all.</w:t></w:r></w:p>
    <w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="Picture 1"/></wp:inline></w:drawing></w:r></w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
    <w:p><w:hyperlink r:id="rId9"><w:r><w:t>https://example.com/page</w:t></w:r></w:hyperlink></w:p>
  </w:body>
</w:document>`;

const BAD_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${NS}>
  <w:style w:type="paragraph" w:styleId="Normal"/>
</w:styles>`;

const GOOD_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}>
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Document title</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Section</w:t></w:r></w:p>
    <w:p><w:r><w:t>Some accessible body text.</w:t></w:r></w:p>
    <w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="Picture 1" descr="A cat sitting on a mat"/></wp:inline></w:drawing></w:r></w:p>
    <w:tbl>
      <w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
    <w:p><w:hyperlink r:id="rId9"><w:r><w:t>Read the guide</w:t></w:r></w:hyperlink></w:p>
  </w:body>
</w:document>`;

const GOOD_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${NS}>
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Heading1"/>
</w:styles>`;

describe("auditDocx", () => {
  test("reports defects in a bad document", () => {
    const pkg = openOoxml(buildDocx(BAD_DOCUMENT, BAD_STYLES));
    const codes = auditDocx(pkg).map((issue) => issue.code);
    expect(codes).toContain("DOCX-ALT-001");
    expect(codes).toContain("DOCX-HEAD-002");
    expect(codes).toContain("DOCX-TBL-005");
    expect(codes).toContain("DOCX-LINK-006");
  });

  test("reports nothing for a good document", () => {
    const pkg = openOoxml(buildDocx(GOOD_DOCUMENT, GOOD_STYLES));
    expect(auditDocx(pkg)).toEqual([]);
  });
});

const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const OFFICE_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const NS_EXT =
  `xmlns:w="${W_NS}" ` +
  `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
  `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
  `xmlns:r="${OFFICE_REL}" ` +
  `xmlns:v="urn:schemas-microsoft-com:vml" ` +
  `xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"`;

const BASE_CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`;

function pkg(entries: Record<string, string | Uint8Array>): Uint8Array {
  const files: Record<string, Uint8Array> = {};
  for (const [path, value] of Object.entries(entries)) {
    files[path] = typeof value === "string" ? enc(value) : value;
  }
  return zipSync(files);
}

function rootRels(main: string, extra = ""): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${REL_NS}">
  <Relationship Id="rId1" Type="${OFFICE_REL}/officeDocument" Target="${main}"/>
  ${extra}
</Relationships>`;
}

function documentRels(rels: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${REL_NS}">${rels}</Relationships>`;
}

function docxDoc(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS_EXT}><w:body>${body}</w:body></w:document>`;
}

const TEXTBOX_DRAWING =
  `<w:p><w:r><w:drawing><wp:inline><wp:docPr id="7" name="TextBox 1"/>` +
  `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
  `<wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>Visible text box copy</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp>` +
  `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;

const vmlImage = (alt: string): string =>
  `<w:p><w:r><w:pict><v:shape id="s1" type="#_x0000_t75"${alt}>` +
  `<v:imagedata r:id="rId1"/></v:shape></w:pict></w:r></w:p>`;

const linkParagraph = (text: string): string =>
  `<w:p><w:hyperlink r:id="rId9"><w:r><w:t>${text}</w:t></w:r></w:hyperlink></w:p>`;

const STYLES_WITH_DEFAULTS_LANG = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${NS_EXT}><w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="en-US"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`;

describe("auditDocx locations", () => {
  test("uses the resolved main part path, not word/document.xml", () => {
    const main = "word/custom-main.xml";
    const data = pkg({
      "[Content_Types].xml": BASE_CONTENT_TYPES,
      "_rels/.rels": rootRels(main),
      [main]: docxDoc(
        `<w:p><w:r><w:t>Body copy.</w:t></w:r></w:p>` +
          `<w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="Picture 1"/></wp:inline></w:drawing></w:r></w:p>`,
      ),
    });

    const alt = auditDocx(openOoxml(data)).find(
      (issue) => issue.code === "DOCX-ALT-001",
    );
    expect(alt?.location).toBe(main);

    const sarif = toSarif(audit(data, "custom.docx"), "officelens", "0.1.0");
    const sarifAlt = sarif.runs[0]!.results.find(
      (result) => result.ruleId === "DOCX-ALT-001",
    );
    expect(sarifAlt?.locations[0]!.physicalLocation.artifactLocation.uri).toBe(
      main,
    );
  });
});

describe("auditDocx alt text", () => {
  test("does not flag a drawing with a non-empty text body", () => {
    const data = pkg({
      "[Content_Types].xml": BASE_CONTENT_TYPES,
      "_rels/.rels": rootRels("word/document.xml"),
      "word/document.xml": docxDoc(TEXTBOX_DRAWING),
    });
    const codes = auditDocx(openOoxml(data)).map((issue) => issue.code);
    expect(codes).not.toContain("DOCX-ALT-001");
  });

  test("flags a VML image without alt text", () => {
    const data = pkg({
      "[Content_Types].xml": BASE_CONTENT_TYPES,
      "_rels/.rels": rootRels("word/document.xml"),
      "word/document.xml": docxDoc(
        `<w:p><w:r><w:t>Body copy.</w:t></w:r></w:p>` + vmlImage(""),
      ),
    });
    const codes = auditDocx(openOoxml(data)).map((issue) => issue.code);
    expect(codes).toContain("DOCX-ALT-001");
  });

  test("accepts a VML image with alt text", () => {
    const data = pkg({
      "[Content_Types].xml": BASE_CONTENT_TYPES,
      "_rels/.rels": rootRels("word/document.xml"),
      "word/document.xml": docxDoc(
        `<w:p><w:r><w:t>Body copy.</w:t></w:r></w:p>` +
          vmlImage(` alt="A chart of quarterly revenue"`),
      ),
    });
    const codes = auditDocx(openOoxml(data)).map((issue) => issue.code);
    expect(codes).not.toContain("DOCX-ALT-001");
  });
});

describe("auditDocx language", () => {
  test("does not count a bare eastAsia lang in the body as a default", () => {
    const data = pkg({
      "[Content_Types].xml": BASE_CONTENT_TYPES,
      "_rels/.rels": rootRels("word/document.xml"),
      "word/document.xml": docxDoc(
        `<w:p><w:pPr><w:rPr><w:lang w:eastAsia="ja-JP"/></w:rPr></w:pPr>` +
          `<w:r><w:t>Hello there.</w:t></w:r></w:p>`,
      ),
    });
    const codes = auditDocx(openOoxml(data)).map((issue) => issue.code);
    expect(codes).toContain("DOCX-LANG-004");
  });

  test("accepts a docDefaults language with w:val", () => {
    const data = pkg({
      "[Content_Types].xml": BASE_CONTENT_TYPES,
      "_rels/.rels": rootRels("word/document.xml"),
      "word/document.xml": docxDoc(`<w:p><w:r><w:t>Hello there.</w:t></w:r></w:p>`),
      "word/styles.xml": STYLES_WITH_DEFAULTS_LANG,
    });
    const codes = auditDocx(openOoxml(data)).map((issue) => issue.code);
    expect(codes).not.toContain("DOCX-LANG-004");
  });

  test("accepts dc:language in core properties", () => {
    const core = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:language>en-US</dc:language></cp:coreProperties>`;
    const data = pkg({
      "[Content_Types].xml": BASE_CONTENT_TYPES,
      "_rels/.rels": rootRels(
        "word/document.xml",
        `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>`,
      ),
      "word/document.xml": docxDoc(`<w:p><w:r><w:t>Hello there.</w:t></w:r></w:p>`),
      "docProps/core.xml": core,
    });
    const codes = auditDocx(openOoxml(data)).map((issue) => issue.code);
    expect(codes).not.toContain("DOCX-LANG-004");
  });

  test("accepts w:themeFontLang in settings", () => {
    const settings = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings ${NS_EXT}><w:themeFontLang w:val="en-US"/></w:settings>`;
    const data = pkg({
      "[Content_Types].xml": BASE_CONTENT_TYPES,
      "_rels/.rels": rootRels("word/document.xml"),
      "word/document.xml": docxDoc(`<w:p><w:r><w:t>Hello there.</w:t></w:r></w:p>`),
      "word/_rels/document.xml.rels": documentRels(
        `<Relationship Id="rId3" Type="${OFFICE_REL}/settings" Target="settings.xml"/>`,
      ),
      "word/settings.xml": settings,
    });
    const codes = auditDocx(openOoxml(data)).map((issue) => issue.code);
    expect(codes).not.toContain("DOCX-LANG-004");
  });
});

describe("auditDocx link text", () => {
  test("flags a link whose text embeds a URL", () => {
    const data = pkg({
      "[Content_Types].xml": BASE_CONTENT_TYPES,
      "_rels/.rels": rootRels("word/document.xml"),
      "word/document.xml": docxDoc(linkParagraph("Click here: https://example.com")),
    });
    const codes = auditDocx(openOoxml(data)).map((issue) => issue.code);
    expect(codes).toContain("DOCX-LINK-006");
  });

  test("flags a bare domain as link text", () => {
    const data = pkg({
      "[Content_Types].xml": BASE_CONTENT_TYPES,
      "_rels/.rels": rootRels("word/document.xml"),
      "word/document.xml": docxDoc(linkParagraph("example.com")),
    });
    const codes = auditDocx(openOoxml(data)).map((issue) => issue.code);
    expect(codes).toContain("DOCX-LINK-006");
  });

  test("flags a mailto link", () => {
    const data = pkg({
      "[Content_Types].xml": BASE_CONTENT_TYPES,
      "_rels/.rels": rootRels("word/document.xml"),
      "word/document.xml": docxDoc(linkParagraph("mailto:team@example.com")),
    });
    const codes = auditDocx(openOoxml(data)).map((issue) => issue.code);
    expect(codes).toContain("DOCX-LINK-006");
  });

  test("does not flag descriptive link text", () => {
    const data = pkg({
      "[Content_Types].xml": BASE_CONTENT_TYPES,
      "_rels/.rels": rootRels("word/document.xml"),
      "word/document.xml": docxDoc(linkParagraph("Read the accessibility guide")),
    });
    const codes = auditDocx(openOoxml(data)).map((issue) => issue.code);
    expect(codes).not.toContain("DOCX-LINK-006");
  });
});

describe("auditDocx related parts", () => {
  test("audits images in header parts and reports their path", () => {
    const header = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr ${NS_EXT}><w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="Logo"/></wp:inline></w:drawing></w:r></w:p></w:hdr>`;
    const data = pkg({
      "[Content_Types].xml": BASE_CONTENT_TYPES,
      "_rels/.rels": rootRels("word/document.xml"),
      "word/document.xml": docxDoc(`<w:p><w:r><w:t>Body copy.</w:t></w:r></w:p>`),
      "word/_rels/document.xml.rels": documentRels(
        `<Relationship Id="rId10" Type="${OFFICE_REL}/header" Target="header1.xml"/>`,
      ),
      "word/header1.xml": header,
    });
    const issues = auditDocx(openOoxml(data));
    const alt = issues.find((issue) => issue.code === "DOCX-ALT-001");
    expect(alt?.location).toBe("word/header1.xml");
  });

  test("resolves a nonstandard styles part from relationships", () => {
    const data = pkg({
      "[Content_Types].xml": BASE_CONTENT_TYPES,
      "_rels/.rels": rootRels("word/document.xml"),
      "word/document.xml": docxDoc(`<w:p><w:r><w:t>Body copy.</w:t></w:r></w:p>`),
      "word/_rels/document.xml.rels": documentRels(
        `<Relationship Id="rId11" Type="${OFFICE_REL}/styles" Target="mystyles.xml"/>`,
      ),
      "word/mystyles.xml": STYLES_WITH_DEFAULTS_LANG,
    });
    const codes = auditDocx(openOoxml(data)).map((issue) => issue.code);
    expect(codes).not.toContain("DOCX-LANG-004");
  });
});
