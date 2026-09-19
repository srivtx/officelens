import { mkdirSync, writeFileSync } from "node:fs";
import { strToU8, zipSync } from "fflate";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WP_NS =
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const PNG_BYTES = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49,
  0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06,
  0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4, 0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44,
  0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0d,
  0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42,
  0x60, 0x82,
]);

type Part = string | Uint8Array;

function zip(parts: Record<string, Part>): Uint8Array {
  const out: Record<string, Uint8Array> = {};
  for (const [path, data] of Object.entries(parts)) {
    out[path] = typeof data === "string" ? strToU8(data) : data;
  }
  return zipSync(out);
}

function docxContentTypes(hasHeadings: boolean): string {
  return (
    XML_HEADER +
    `<Types xmlns="${CT_NS}">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Default Extension="png" ContentType="image/png"/>` +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
    `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
    `</Types>`
  );
}

function docxRootRels(): string {
  return (
    XML_HEADER +
    `<Relationships xmlns="${REL_NS}">` +
    `<Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>` +
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
    `</Relationships>`
  );
}

function docxDocumentRels(): string {
  return (
    XML_HEADER +
    `<Relationships xmlns="${REL_NS}">` +
    `<Relationship Id="rId1" Type="${R_NS}/image" Target="media/image1.png"/>` +
    `<Relationship Id="rId5" Type="${R_NS}/hyperlink" Target="https://example.com" TargetMode="External"/>` +
    `</Relationships>`
  );
}

function docxCoreProps(title: string): string {
  return (
    XML_HEADER +
    `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">` +
    `<dc:title>${title}</dc:title>` +
    `<dc:language>en-US</dc:language>` +
    `</cp:coreProperties>`
  );
}

function docxStyles(hasHeadings: boolean, withLang: boolean): string {
  const lang = withLang ? `<w:lang w:val="en-US" w:eastAsia="en-US"/>` : "";
  const headingStyles = hasHeadings
    ? `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>` +
      `<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>`
    : "";
  return (
    XML_HEADER +
    `<w:styles xmlns:w="${W_NS}">` +
    `<w:docDefaults><w:rPrDefault><w:rPr>${lang}</w:rPr></w:rPrDefault></w:docDefaults>` +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
    headingStyles +
    `</w:styles>`
  );
}

function docxDrawing(withAlt: boolean): string {
  const alt = withAlt
    ? ` descr="A sample image used as a placeholder graphic"`
    : "";
  return (
    `<w:p><w:r><w:drawing>` +
    `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="914400" cy="914400"/>` +
    `<wp:docPr id="1" name="Picture 1"${alt}/>` +
    `<a:graphic><a:graphicData uri="${PIC_NS}">` +
    `<pic:pic>` +
    `<pic:nvPicPr><pic:cNvPr id="1" name="Picture 1"${alt}/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic>` +
    `</a:graphicData></a:graphic>` +
    `</wp:inline>` +
    `</w:drawing></w:r></w:p>`
  );
}

function docxTable(withHeader: boolean): string {
  const headerRowProps = withHeader ? `<w:trPr><w:tblHeader/></w:trPr>` : "";
  const cell = (text: string) =>
    `<w:tc><w:tcPr/><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  return (
    `<w:tbl>` +
    `<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
    `<w:tr>${headerRowProps}${cell("Name")}${cell("Value")}</w:tr>` +
    `<w:tr>${cell("Alpha")}${cell("One")}</w:tr>` +
    `</w:tbl>`
  );
}

function docxDocument(options: {
  hasHeadings: boolean;
  withAlt: boolean;
  linkText: string;
  withHeader: boolean;
}): string {
  const heading = (level: 1 | 2, text: string) =>
    `<w:p><w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const paragraph = (text: string) =>
    `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

  const intro = options.hasHeadings
    ? heading(1, "Introduction") + paragraph("Some introductory body text.")
    : paragraph("Introduction") + paragraph("Some introductory body text.");

  const section = options.hasHeadings
    ? heading(2, "Details")
    : "";

  const link =
    `<w:p><w:r><w:t>Read more: </w:t></w:r>` +
    `<w:hyperlink r:id="rId5"><w:r><w:t>${options.linkText}</w:t></w:r></w:hyperlink></w:p>`;

  return (
    XML_HEADER +
    `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}">` +
    `<w:body>` +
    intro +
    section +
    docxDrawing(options.withAlt) +
    link +
    docxTable(options.withHeader) +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>` +
    `</w:body></w:document>`
  );
}

export function makeBadDocx(): Uint8Array {
  return zip({
    "[Content_Types].xml": docxContentTypes(false),
    "_rels/.rels": docxRootRels(),
    "docProps/core.xml": docxCoreProps("Inaccessible Document"),
    "word/document.xml": docxDocument({
      hasHeadings: false,
      withAlt: false,
      linkText: "https://example.com",
      withHeader: false,
    }),
    "word/_rels/document.xml.rels": docxDocumentRels(),
    "word/styles.xml": docxStyles(false, false),
    "word/media/image1.png": PNG_BYTES,
  });
}

export function makeGoodDocx(): Uint8Array {
  return zip({
    "[Content_Types].xml": docxContentTypes(true),
    "_rels/.rels": docxRootRels(),
    "docProps/core.xml": docxCoreProps("Accessible Document"),
    "word/document.xml": docxDocument({
      hasHeadings: true,
      withAlt: true,
      linkText: "Example",
      withHeader: true,
    }),
    "word/_rels/document.xml.rels": docxDocumentRels(),
    "word/styles.xml": docxStyles(true, true),
    "word/media/image1.png": PNG_BYTES,
  });
}

function pptxContentTypes(): string {
  return (
    XML_HEADER +
    `<Types xmlns="${CT_NS}">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Default Extension="png" ContentType="image/png"/>` +
    `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
    `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
    `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
    `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
    `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
    `</Types>`
  );
}

function pptxRootRels(): string {
  return (
    XML_HEADER +
    `<Relationships xmlns="${REL_NS}">` +
    `<Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="ppt/presentation.xml"/>` +
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
    `</Relationships>`
  );
}

function pptxPresentation(): string {
  return (
    XML_HEADER +
    `<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">` +
    `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
    `<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>` +
    `<p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>` +
    `<p:notesSz cx="6858000" cy="9144000"/>` +
    `</p:presentation>`
  );
}

function pptxPresentationRels(): string {
  return (
    XML_HEADER +
    `<Relationships xmlns="${REL_NS}">` +
    `<Relationship Id="rId1" Type="${R_NS}/slideMaster" Target="slideMasters/slideMaster1.xml"/>` +
    `<Relationship Id="rId2" Type="${R_NS}/slide" Target="slides/slide1.xml"/>` +
    `</Relationships>`
  );
}

function pptxSlideMaster(): string {
  return (
    XML_HEADER +
    `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">` +
    `<p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr/>` +
    `</p:spTree></p:cSld>` +
    `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
    `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>` +
    `</p:sldMaster>`
  );
}

function pptxSlideMasterRels(): string {
  return (
    XML_HEADER +
    `<Relationships xmlns="${REL_NS}">` +
    `<Relationship Id="rId1" Type="${R_NS}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
    `</Relationships>`
  );
}

function pptxSlideLayout(): string {
  return (
    XML_HEADER +
    `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="title" preserve="1">` +
    `<p:cSld name="Title Slide"><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr/>` +
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>` +
    `</p:spTree></p:cSld>` +
    `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>` +
    `</p:sldLayout>`
  );
}

function pptxSlideLayoutRels(): string {
  return (
    XML_HEADER +
    `<Relationships xmlns="${REL_NS}">` +
    `<Relationship Id="rId1" Type="${R_NS}/slideMaster" Target="../slideMasters/slideMaster1.xml"/>` +
    `</Relationships>`
  );
}

function pptxSlideRels(): string {
  return (
    XML_HEADER +
    `<Relationships xmlns="${REL_NS}">` +
    `<Relationship Id="rId1" Type="${R_NS}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
    `<Relationship Id="rId2" Type="${R_NS}/image" Target="../media/image1.png"/>` +
    `</Relationships>`
  );
}

function pptxSlide(options: {
  withTitle: boolean;
  withAlt: boolean;
  withHeader: boolean;
}): string {
  const alt = options.withAlt ? ` descr="A sample image used in the deck"` : "";
  const tableAlt = options.withAlt ? ` descr="A data table with two columns"` : "";
  const title = options.withTitle
    ? `<p:sp><p:nvSpPr><p:cNvPr id="4" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Accessible Deck</a:t></a:r></a:p></p:txBody></p:sp>`
    : "";
  const tblPr = options.withHeader ? `<a:tblPr firstRow="1"/>` : `<a:tblPr/>`;
  return (
    XML_HEADER +
    `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">` +
    `<p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr/>` +
    title +
    `<p:pic>` +
    `<p:nvPicPr><p:cNvPr id="2" name="Picture 1"${alt}/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `</p:pic>` +
    `<p:graphicFrame>` +
    `<p:nvGraphicFramePr><p:cNvPr id="3" name="Table 1"${tableAlt}/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">` +
    `<a:tbl>${tblPr}` +
    `<a:tblGrid><a:gridCol w="1000000"/><a:gridCol w="1000000"/></a:tblGrid>` +
    `<a:tr h="370840"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Name</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Value</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>` +
    `<a:tr h="370840"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Alpha</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>One</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>` +
    `</a:tbl>` +
    `</a:graphicData></a:graphic>` +
    `</p:graphicFrame>` +
    `</p:spTree></p:cSld>` +
    `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>` +
    `</p:sld>`
  );
}

export function makeBadPptx(): Uint8Array {
  return zip({
    "[Content_Types].xml": pptxContentTypes(),
    "_rels/.rels": pptxRootRels(),
    "docProps/core.xml": docxCoreProps("Inaccessible Deck"),
    "ppt/presentation.xml": pptxPresentation(),
    "ppt/_rels/presentation.xml.rels": pptxPresentationRels(),
    "ppt/slides/slide1.xml": pptxSlide({
      withTitle: false,
      withAlt: false,
      withHeader: false,
    }),
    "ppt/slides/_rels/slide1.xml.rels": pptxSlideRels(),
    "ppt/slideMasters/slideMaster1.xml": pptxSlideMaster(),
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": pptxSlideMasterRels(),
    "ppt/slideLayouts/slideLayout1.xml": pptxSlideLayout(),
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": pptxSlideLayoutRels(),
    "ppt/media/image1.png": PNG_BYTES,
  });
}

export function makeGoodPptx(): Uint8Array {
  return zip({
    "[Content_Types].xml": pptxContentTypes(),
    "_rels/.rels": pptxRootRels(),
    "docProps/core.xml": docxCoreProps("Accessible Deck"),
    "ppt/presentation.xml": pptxPresentation(),
    "ppt/_rels/presentation.xml.rels": pptxPresentationRels(),
    "ppt/slides/slide1.xml": pptxSlide({
      withTitle: true,
      withAlt: true,
      withHeader: true,
    }),
    "ppt/slides/_rels/slide1.xml.rels": pptxSlideRels(),
    "ppt/slideMasters/slideMaster1.xml": pptxSlideMaster(),
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": pptxSlideMasterRels(),
    "ppt/slideLayouts/slideLayout1.xml": pptxSlideLayout(),
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": pptxSlideLayoutRels(),
    "ppt/media/image1.png": PNG_BYTES,
  });
}

if (import.meta.main) {
  const fixturesDir = new URL("../fixtures/", import.meta.url);
  mkdirSync(fixturesDir, { recursive: true });
  writeFileSync(new URL("bad.docx", fixturesDir), makeBadDocx());
  writeFileSync(new URL("good.docx", fixturesDir), makeGoodDocx());
  writeFileSync(new URL("bad.pptx", fixturesDir), makeBadPptx());
  writeFileSync(new URL("good.pptx", fixturesDir), makeGoodPptx());
  console.log(
    "Wrote fixtures/bad.docx, fixtures/good.docx, fixtures/bad.pptx, fixtures/good.pptx",
  );
}
