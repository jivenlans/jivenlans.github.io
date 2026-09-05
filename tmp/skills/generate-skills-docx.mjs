import fs from "node:fs";
import path from "node:path";

const outPath = path.resolve("documents", "Jivenlans_Tabien_Skills_Bullet_Summary.docx");

const NS = {
  cp: "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
  dc: "http://purl.org/dc/elements/1.1/",
  dcterms: "http://purl.org/dc/terms/",
  dcmitype: "http://purl.org/dc/dcmitype/",
  xsi: "http://www.w3.org/2001/XMLSchema-instance",
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
};

const technicalSkills = [
  ["Application Types Developed", "Web applications, Web APIs and services, SOAP and REST services, background services, Windows services, and console applications"],
  ["Languages", "C#, JavaScript, TypeScript, SQL, HTML, and CSS"],
  ["Backend and Frameworks", ".NET Core, .NET 8 and 9, ASP.NET Core Web API, ASP.NET MVC, ASP.NET Web Forms, Blazor, Entity Framework Core, and xUnit"],
  ["Frontend and UI", "React, TypeScript, JavaScript, HTML, CSS, Bootstrap, AJAX, jQuery, and SignalR"],
  ["API and Testing Tools", "REST API, GraphQL, SOAP, Swagger, Postman, SOAPUI, and Mockoon"],
  ["Databases and Data Access", "SQL Server, PostgreSQL, MySQL, SQLite, MongoDB, Google Cloud Datastore, Azure Blob Storage, Entity Framework Core, and ADO.NET"],
  ["Security and Authentication", "JWT, PGP keys, ASP.NET Identity, SQL encryption, and data encryption and decryption with Bouncy Castle"],
  ["Messaging and Background Jobs", "NATS, Kafka, Azure Queues, Hangfire jobs, SQL jobs, Windows services, and console applications"],
  ["Cloud and Serverless", "Azure Web App Service, Azure Functions, Google Cloud Run Functions, Railway, Cloudflare, GitHub Pages, and Supabase"],
  ["Version Control", "Git, GitHub repositories, Azure DevOps repositories, and TFS repositories"],
  ["Deployment and DevOps", "Docker, GitHub Actions, Azure DevOps Pipelines, Jenkins, Fortify, Azure Portal, and Azure Storage Explorer"],
  ["Monitoring and Logging", "Azure Application Insights, Google Cloud Logs Explorer, Datadog, Serilog, and log4net"],
  ["External Communications", "CDyne, Twilio, Mailgun, FTP, and SFTP"],
  ["Reporting", "SSRS and ClosedXML"],
  ["AI Assisted Development", "Codex, GitHub Copilot, ChatGPT, Copilot, and agentic AI workflows"],
];

const designSkills = [
  ["Software Design", "Object-oriented programming, SOLID principles, DRY, design patterns, dependency injection, and inversion of control"],
  ["Software Architecture", "N-tier architecture, layered architecture, Clean Architecture, CQRS, and microservices"],
  ["Software Methodologies", "Waterfall, Agile, and Scrum delivery"],
];

const professionalSkills = [
  "Leadership and team mentoring",
  "Problem solving and debugging",
  "Communication and collaboration",
  "Analytical thinking for business scenarios",
  "Production support and operational troubleshooting",
  "Code reviews and developer guidance",
  "Time management",
];

function esc(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function rPr({ bold = false, italic = false, size = 21, color = "000000", font = "Aptos" } = {}) {
  return [
    "<w:rPr>",
    `<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/>`,
    bold ? "<w:b/>" : "",
    italic ? "<w:i/>" : "",
    `<w:color w:val="${color}"/>`,
    `<w:sz w:val="${size}"/>`,
    `<w:szCs w:val="${size}"/>`,
    "</w:rPr>",
  ].join("");
}

function pPr({
  style = "",
  before = 0,
  after = 120,
  line = 276,
  align = "",
  keepNext = false,
  bullet = false,
  left = 0,
  hanging = 0,
} = {}) {
  return [
    "<w:pPr>",
    style ? `<w:pStyle w:val="${style}"/>` : "",
    keepNext ? "<w:keepNext/>" : "",
    bullet ? "<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>" : "",
    `<w:spacing w:before="${before}" w:after="${after}" w:line="${line}" w:lineRule="auto"/>`,
    left ? `<w:ind w:left="${left}"${hanging ? ` w:hanging="${hanging}"` : ""}/>` : "",
    align ? `<w:jc w:val="${align}"/>` : "",
    "</w:pPr>",
  ].join("");
}

function run(text, options = {}) {
  return `<w:r>${rPr(options)}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
}

function paragraph(parts, options = {}) {
  const runs = (Array.isArray(parts) ? parts : [{ text: parts }]).map((part) => run(part.text, { ...options, ...part })).join("");
  return `<w:p>${pPr(options)}${runs}</w:p>`;
}

function sectionHeading(text) {
  return paragraph(text, {
    style: "Heading1",
    bold: true,
    size: 25,
    before: 300,
    after: 100,
    keepNext: true,
  });
}

function skillBullet(label, text) {
  return paragraph(
    [
      { text: `${label}: `, bold: true },
      { text },
    ],
    {
      bullet: true,
      size: 21,
      after: 95,
      left: 360,
      hanging: 360,
    },
  );
}

function plainBullet(text) {
  return paragraph(text, {
    bullet: true,
    size: 21,
    after: 80,
    left: 360,
    hanging: 360,
  });
}

function buildBody() {
  return [
    paragraph("Jivenlans Tabien Skills Bullet Summary", {
      style: "Title",
      bold: true,
      size: 34,
      after: 160,
      line: 320,
    }),
    paragraph("This document summarizes the core technical, software design, and professional skills of Jivenlans Tabien for resume, portfolio, and role application use.", {
      size: 22,
      after: 80,
      line: 300,
    }),
    paragraph("The skill set is centered on .NET application delivery, APIs, integrations, cloud deployment, production support, and practical engineering leadership.", {
      size: 22,
      after: 180,
      line: 300,
    }),
    sectionHeading("Technical Skills"),
    ...technicalSkills.map(([label, text]) => skillBullet(label, text)),
    sectionHeading("Software Design and Delivery"),
    ...designSkills.map(([label, text]) => skillBullet(label, text)),
    sectionHeading("Professional Skills"),
    ...professionalSkills.map((text) => plainBullet(text)),
  ].join("");
}

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS.w}" xmlns:r="${NS.r}">
  <w:body>
    ${buildBody()}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/>
      <w:cols w:space="720"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>`;

const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${NS.w}">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:cs="Aptos"/>
        <w:sz w:val="21"/>
        <w:szCs w:val="21"/>
        <w:color w:val="000000"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="120" w:line="276" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="0" w:after="160"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Aptos Display" w:hAnsi="Aptos Display" w:cs="Aptos Display"/>
      <w:b/>
      <w:color w:val="000000"/>
      <w:sz w:val="34"/>
      <w:szCs w:val="34"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="300" w:after="100"/>
      <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Aptos Display" w:hAnsi="Aptos Display" w:cs="Aptos Display"/>
      <w:b/>
      <w:color w:val="000000"/>
      <w:sz w:val="25"/>
      <w:szCs w:val="25"/>
    </w:rPr>
  </w:style>
</w:styles>`;

const numberingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${NS.w}">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="singleLevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#8226;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:tabs>
          <w:tab w:val="num" w:pos="360"/>
        </w:tabs>
        <w:ind w:left="360" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>
      </w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>`;

const settingsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="${NS.w}">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
</w:settings>`;

const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;

const documentRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`;

const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;

const coreXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="${NS.cp}" xmlns:dc="${NS.dc}" xmlns:dcterms="${NS.dcterms}" xmlns:dcmitype="${NS.dcmitype}" xmlns:xsi="${NS.xsi}">
  <dc:title>Jivenlans Tabien Skills Bullet Summary</dc:title>
  <dc:creator>Jivenlans Tabien</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-09-05T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-09-05T00:00:00Z</dcterms:modified>
</cp:coreProperties>`;

const appXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <Company/>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>`;

const files = [
  { name: "[Content_Types].xml", data: Buffer.from(contentTypesXml, "utf8") },
  { name: "_rels/.rels", data: Buffer.from(relsXml, "utf8") },
  { name: "docProps/core.xml", data: Buffer.from(coreXml, "utf8") },
  { name: "docProps/app.xml", data: Buffer.from(appXml, "utf8") },
  { name: "word/document.xml", data: Buffer.from(documentXml, "utf8") },
  { name: "word/_rels/document.xml.rels", data: Buffer.from(documentRelsXml, "utf8") },
  { name: "word/styles.xml", data: Buffer.from(stylesXml, "utf8") },
  { name: "word/numbering.xml", data: Buffer.from(numberingXml, "utf8") },
  { name: "word/settings.xml", data: Buffer.from(settingsXml, "utf8") },
];

function crc32(buffer) {
  let crc = -1;
  for (const byte of buffer) {
    crc ^= byte;
    for (let i = 0; i < 8; i += 1) {
      crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
    }
  }
  return (crc ^ -1) >>> 0;
}

function dosDateTime(date = new Date()) {
  const time = (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2);
  const dosDate = ((date.getFullYear() - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate();
  return { time, date: dosDate };
}

function writeZip(entries, targetPath) {
  const chunks = [];
  const central = [];
  let offset = 0;
  const { time, date } = dosDateTime();

  for (const entry of entries) {
    const name = Buffer.from(entry.name, "utf8");
    const data = entry.data;
    const crc = crc32(data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6);
    local.writeUInt16LE(0, 8);
    local.writeUInt16LE(time, 10);
    local.writeUInt16LE(date, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(name.length, 26);
    local.writeUInt16LE(0, 28);
    chunks.push(local, name, data);

    const cen = Buffer.alloc(46);
    cen.writeUInt32LE(0x02014b50, 0);
    cen.writeUInt16LE(20, 4);
    cen.writeUInt16LE(20, 6);
    cen.writeUInt16LE(0, 8);
    cen.writeUInt16LE(0, 10);
    cen.writeUInt16LE(time, 12);
    cen.writeUInt16LE(date, 14);
    cen.writeUInt32LE(crc, 16);
    cen.writeUInt32LE(data.length, 20);
    cen.writeUInt32LE(data.length, 24);
    cen.writeUInt16LE(name.length, 28);
    cen.writeUInt16LE(0, 30);
    cen.writeUInt16LE(0, 32);
    cen.writeUInt16LE(0, 34);
    cen.writeUInt16LE(0, 36);
    cen.writeUInt32LE(0, 38);
    cen.writeUInt32LE(offset, 42);
    central.push(cen, name);
    offset += local.length + name.length + data.length;
  }

  const centralStart = offset;
  const centralSize = central.reduce((sum, chunk) => sum + chunk.length, 0);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(centralSize, 12);
  eocd.writeUInt32LE(centralStart, 16);
  eocd.writeUInt16LE(0, 20);

  fs.mkdirSync(path.dirname(targetPath), { recursive: true });
  fs.writeFileSync(targetPath, Buffer.concat([...chunks, ...central, eocd]));
}

writeZip(files, outPath);
console.log(`Wrote ${outPath}`);
