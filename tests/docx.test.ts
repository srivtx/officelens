/// <reference types="bun" />
import { describe, expect, test } from "bun:test";
import { strToU8, zipSync } from "fflate";
import { openOoxml } from "../src/package.ts";
import { auditDocx } from "../src/docx.ts";

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
