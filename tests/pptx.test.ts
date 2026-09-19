import { describe, expect, test } from "bun:test";
import { strToU8, zipSync } from "fflate";
import { auditPptx } from "../src/pptx";
import { audit } from "../src/audit";
import { openOoxml } from "../src/package";

const NS_DECL = [
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
].join(" ");

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;

const PRESENTATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation ${NS_DECL}>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
</p:presentation>`;

const PRESENTATION_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`;

const BAD_SLIDE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld ${NS_DECL}>
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:pic>
        <p:nvPicPr><p:cNvPr id="2" name="Picture 1"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
        <p:blipFill/>
        <p:spPr/>
      </p:pic>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="3" name="Table 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <a:graphic><a:graphicData><a:tbl><a:tblPr/><a:tblGrid/></a:tbl></a:graphicData></a:graphic>
      </p:graphicFrame>
    </p:spTree>
  </p:cSld>
</p:sld>`;

const GOOD_SLIDE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld ${NS_DECL}>
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Hello</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:pic>
        <p:nvPicPr><p:cNvPr id="3" name="Picture 1" descr="A nice picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
        <p:blipFill/>
        <p:spPr/>
      </p:pic>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="4" name="Table 1" descr="Data table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <a:graphic><a:graphicData><a:tbl><a:tblPr firstRow="1"/><a:tblGrid/></a:tbl></a:graphicData></a:graphic>
      </p:graphicFrame>
    </p:spTree>
  </p:cSld>
</p:sld>`;

function makePptx(slideXml: string): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8(CONTENT_TYPES),
    "_rels/.rels": strToU8(ROOT_RELS),
    "ppt/presentation.xml": strToU8(PRESENTATION),
    "ppt/_rels/presentation.xml.rels": strToU8(PRESENTATION_RELS),
    "ppt/slides/slide1.xml": strToU8(slideXml),
  });
}

describe("auditPptx", () => {
  test("reports alt text, title and table defects on a bad fixture", () => {
    const issues = auditPptx(openOoxml(makePptx(BAD_SLIDE)));
    const codes = issues.map((issue) => issue.code);
    expect(codes).toContain("PPTX-ALT-001");
    expect(codes).toContain("PPTX-TITLE-002");
    expect(codes).toContain("PPTX-TBL-003");
  });

  test("reports no defects on a good fixture", () => {
    const issues = auditPptx(openOoxml(makePptx(GOOD_SLIDE)));
    expect(issues.filter((issue) => issue.code.startsWith("PPTX-"))).toHaveLength(0);
  });
});

describe("audit", () => {
  test("detects pptx and surfaces the defects end to end", () => {
    const result = audit(makePptx(BAD_SLIDE), "bad.pptx");
    expect(result.kind).toBe("pptx");
    const codes = result.issues.map((issue) => issue.code);
    expect(codes).toContain("PPTX-ALT-001");
    expect(codes).toContain("PPTX-TITLE-002");
    expect(codes).toContain("PPTX-TBL-003");
  });

  test("good fixture produces no pptx issues through audit", () => {
    const result = audit(makePptx(GOOD_SLIDE), "good.pptx");
    expect(result.kind).toBe("pptx");
    expect(result.issues.filter((issue) => issue.code.startsWith("PPTX-"))).toHaveLength(0);
  });
});
