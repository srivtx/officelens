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

const BASE_CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
</Types>`;

function deck(entries: Record<string, string>): Uint8Array {
  const files: Record<string, Uint8Array> = {};
  for (const [path, value] of Object.entries(entries)) files[path] = strToU8(value);
  return zipSync(files);
}

function presentation(slideIds: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation ${NS_DECL}><p:sldIdLst>${slideIds}</p:sldIdLst></p:presentation>`;
}

function presentationRels(rels: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
}

const DECK_ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;

function slide(inner: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld ${NS_DECL}><p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr/>${inner}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

const PICTURE = `<p:pic><p:nvPicPr><p:cNvPr id="9" name="Picture 1" descr="A chart"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic>`;

function textBox(lang?: string): string {
  const rPr = lang ? `<a:rPr lang="${lang}"/>` : "";
  return `<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>${rPr}<a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp>`;
}

const LAYOUT_WITH_TITLE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout ${NS_DECL} type="title"><p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody/></p:sp>
</p:spTree></p:cSld></p:sldLayout>`;

const MASTER_WITH_LANG = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster ${NS_DECL}><p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
</p:spTree></p:cSld>
<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr lang="en-US"/></a:lvl1pPr></p:titleStyle><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>`;

const MASTER_NO_LANG = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster ${NS_DECL}><p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
</p:spTree></p:cSld></p:sldMaster>`;

describe("auditPptx title inheritance", () => {
  test("does not flag a slide whose layout supplies the title", () => {
    const data = deck({
      "[Content_Types].xml": BASE_CONTENT_TYPES,
      "_rels/.rels": DECK_ROOT_RELS,
      "ppt/presentation.xml": presentation(`<p:sldId id="256" r:id="rId1"/>`),
      "ppt/_rels/presentation.xml.rels": presentationRels(
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>`,
      ),
      "ppt/slides/slide1.xml": slide(PICTURE),
      "ppt/slides/_rels/slide1.xml.rels": presentationRels(
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`,
      ),
      "ppt/slideLayouts/slideLayout1.xml": LAYOUT_WITH_TITLE,
    });
    const issues = auditPptx(openOoxml(data));
    expect(issues.map((issue) => issue.code)).not.toContain("PPTX-TITLE-002");
  });

  test("reports a missing title as a warning, not an error", () => {
    const issues = auditPptx(openOoxml(makePptx(BAD_SLIDE)));
    const title = issues.find((issue) => issue.code === "PPTX-TITLE-002");
    expect(title?.severity).toBe("warning");
  });
});

describe("auditPptx language", () => {
  test("honors a language defined in master txStyles", () => {
    const data = deck({
      "[Content_Types].xml": BASE_CONTENT_TYPES,
      "_rels/.rels": DECK_ROOT_RELS,
      "ppt/presentation.xml": presentation(`<p:sldId id="256" r:id="rId1"/>`),
      "ppt/_rels/presentation.xml.rels": presentationRels(
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>`,
      ),
      "ppt/slides/slide1.xml": slide(textBox()),
      "ppt/slideMasters/slideMaster1.xml": MASTER_WITH_LANG,
    });
    const codes = auditPptx(openOoxml(data)).map((issue) => issue.code);
    expect(codes).not.toContain("PPTX-LANG-004");
  });

  test("honors a language defined in a layout txStyles", () => {
    const layoutWithLang = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout ${NS_DECL}><p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
</p:spTree></p:cSld>
<p:txStyles><p:titleStyle/><p:bodyStyle><a:lvl1pPr><a:defRPr lang="fr-FR"/></a:lvl1pPr></p:bodyStyle><p:otherStyle/></p:txStyles>
</p:sldLayout>`;
    const data = deck({
      "[Content_Types].xml": BASE_CONTENT_TYPES,
      "_rels/.rels": DECK_ROOT_RELS,
      "ppt/presentation.xml": presentation(`<p:sldId id="256" r:id="rId1"/>`),
      "ppt/_rels/presentation.xml.rels": presentationRels(
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>`,
      ),
      "ppt/slides/slide1.xml": slide(textBox()),
      "ppt/slides/_rels/slide1.xml.rels": presentationRels(
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`,
      ),
      "ppt/slideLayouts/slideLayout1.xml": layoutWithLang,
    });
    const codes = auditPptx(openOoxml(data)).map((issue) => issue.code);
    expect(codes).not.toContain("PPTX-LANG-004");
  });

  test("reports language per slide and points at the slide part", () => {
    const data = deck({
      "[Content_Types].xml": BASE_CONTENT_TYPES,
      "_rels/.rels": DECK_ROOT_RELS,
      "ppt/presentation.xml": presentation(
        `<p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/>`,
      ),
      "ppt/_rels/presentation.xml.rels": presentationRels(
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
          `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>`,
      ),
      "ppt/slides/slide1.xml": slide(textBox("en-US")),
      "ppt/slides/slide2.xml": slide(textBox()),
      "ppt/slideMasters/slideMaster1.xml": MASTER_NO_LANG,
    });
    const lang = auditPptx(openOoxml(data)).filter(
      (issue) => issue.code === "PPTX-LANG-004",
    );
    expect(lang).toHaveLength(1);
    expect(lang[0]!.location).toBe("ppt/slides/slide2.xml");
  });
});
