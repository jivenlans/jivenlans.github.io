import fs from "node:fs";
import path from "node:path";

const outPath = path.resolve("documents", "Resume - Jivenlans Tabien.docx");
const imagePath = path.resolve("images", "profile_picture.jpg");

const NS = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
};

function esc(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function rPr({ bold = false, size = 20, color = "000000", font = "Georgia" } = {}) {
  return `<w:rPr>${bold ? "<w:b/>" : ""}<w:color w:val="${color}"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/></w:rPr>`;
}

function pPr({ align = "", before = 0, after = 80, left = 0, hanging = 0, shading = "", keepNext = false } = {}) {
  return `<w:pPr>${align ? `<w:jc w:val="${align}"/>` : ""}<w:spacing w:before="${before}" w:after="${after}"/>${left ? `<w:ind w:left="${left}"${hanging ? ` w:hanging="${hanging}"` : ""}/>` : ""}${shading ? `<w:shd w:val="clear" w:color="auto" w:fill="${shading}"/>` : ""}${keepNext ? "<w:keepNext/>" : ""}</w:pPr>`;
}

function textRun(text, options = {}) {
  return `<w:r>${rPr(options)}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
}

function para(text = "", options = {}) {
  return `<w:p>${pPr(options)}${textRun(text, options)}</w:p>`;
}

function mixedPara(parts, options = {}) {
  return `<w:p>${pPr(options)}${parts.map((part) => textRun(part.text, { ...options, ...part })).join("")}</w:p>`;
}

function bullet(text, options = {}) {
  return mixedPara(
    [
      { text: "\u2022 ", bold: false },
      ...(Array.isArray(text) ? text : [{ text }]),
    ],
    { ...options, left: options.left ?? 260, hanging: options.hanging ?? 140, after: options.after ?? 45 },
  );
}

function heading(text, color = "000000", fill = "E9E9E9") {
  return para(text.toUpperCase(), {
    bold: true,
    size: 22,
    color,
    align: "center",
    shading: fill,
    before: 140,
    after: 110,
    keepNext: true,
  });
}

function leftHeading(text) {
  return heading(text, "FFFFFF", "214F65");
}

function rightHeading(text) {
  return heading(text, "000000", "E9E9E9");
}

function imagePara() {
  if (!fs.existsSync(imagePath)) return "";
  return `<w:p>${pPr({ align: "center", before: 240, after: 260 })}<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1295400" cy="1295400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Profile Picture"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${NS.a}" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="${NS.a}"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="${NS.pic}"><pic:nvPicPr><pic:cNvPr id="0" name="profile_picture.jpg"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1295400" cy="1295400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
}

function contactLine(text) {
  return para(text, { color: "FFFFFF", size: 19, after: 115 });
}

function sidePageOne() {
  return [
    para("Jivenlans Tabien", { bold: true, size: 36, color: "FFFFFF", before: 650, after: 160 }),
    para("Senior Software Developer", { size: 26, color: "FFFFFF", after: 220 }),
    imagePara(),
    contactLine("jivenlans@gmail.com"),
    contactLine("+639358777395"),
    contactLine("Manila, Philippines"),
    contactLine("03/07/1996"),
    contactLine("linkedin.com/in/jivenlans-tabien"),
    contactLine("https://jivenlans.github.io/"),
    leftHeading("Profile"),
    para("Senior .NET Developer with over 9 years in designing and delivering scalable applications like Web, APIs, and background service. Proficient in .NET Core, Blazor and Javascript with expertise in building performant, user-friendly solutions. Skilled in leading teams, mentoring developers, and ensuring seamless integration of systems.", { color: "FFFFFF", size: 19, after: 360 }),
    leftHeading("Education"),
    para("Information Technology", { bold: true, color: "FFFFFF", size: 19, after: 10 }),
    para("Adamson University", { color: "FFFFFF", size: 19, after: 10 }),
    para("06/2012 - 03/2016 | Manila, Philippines", { color: "FFFFFF", size: 19, after: 350 }),
    leftHeading("Trainings Attended"),
    bullet("Xamarin Mobile App Development Training Program (2017)", { color: "FFFFFF", size: 18, after: 85 }),
    bullet("SQL Server Training Program (2017)", { color: "FFFFFF", size: 18, after: 85 }),
    bullet("Outsystems Training Program (2018)", { color: "FFFFFF", size: 18, after: 85 }),
    bullet("Blazor Training (2024)", { color: "FFFFFF", size: 18, after: 85 }),
    bullet("Docker Training (2024)", { color: "FFFFFF", size: 18, after: 85 }),
    bullet("Unit Testing - XUnit & NUnit (2025)", { color: "FFFFFF", size: 18, after: 40 }),
  ].join("");
}

function skillBullet(label, text) {
  return bullet([
    { text: label, bold: true },
    { text: ` - ${text}` },
  ], { size: 19, after: 95 });
}

function sideBlank() {
  return para("", { color: "FFFFFF", size: 18, before: 12000, after: 0 });
}

function expHeader(company, role, datePlace) {
  return [
    para(company, { bold: true, size: 19, after: 20 }),
    para(role, { size: 19, after: 20 }),
    para(datePlace, { size: 18, after: 50 }),
  ].join("");
}

function pageOneMain() {
  return [
    rightHeading("Skills"),
    para("Technical Skills", { bold: true, size: 19, after: 100 }),
    skillBullet("Application Types Developed", "Web, Web APIs/Services (SOAP & REST), Background Services/Console Applications"),
    skillBullet("Languages", "C# (.NET Framework, .NET Core), Javascript, SQL, HTML/CSS"),
    skillBullet("Web Frameworks & Technologies", "ASP.NET Core API, ASP.NET MVC, ASP.NET Webforms, Blazor, Swagger, AJAX, SignalR, JQuery, Bootstrap"),
    skillBullet("Databases and Data Access", "SQL Server, SQLite, PostgreSQL, MySQL, Entity Framework Core, ADO.NET"),
    skillBullet("Logging", "Serilog, Datadog, Log4net"),
    skillBullet("Asynchronous Messaging & Background Jobs", "NATS, Kafka, Azure Queues, Hangfire, SQL Jobs"),
    skillBullet("Reporting", "SSRS, ClosedXML"),
    skillBullet("File Transfer", "SFTP, FTP"),
    skillBullet("Security & Authentication", "JWT, PGP Keys, ASP.Net Identity, SQL Encryption, Data Encryption/Decryption (Bouncy Castle)"),
    skillBullet("Software Development and API Testing Tools", "Visual Studio, Visual Studio Code, Mockoon, Postman, SOAPUI"),
    skillBullet("Version Control & DevOps Tools", "Git, TFS"),
    skillBullet("DevOps Tools", "Jenkins, Fortify, Github Actions, Docker, Azure Portal, Azure Storage Explorer"),
    skillBullet("Software Design", "Object-Oriented Programming (OOP), SOLID Principles, DRY, Design Patterns, Dependency Injection, Inversion of Control (IoC)"),
    skillBullet("Software Architectures", "N-Tier (Layered) Architecture, Clean Architecture, CQRS"),
    skillBullet("Software Development Life Cycle", "Waterfall, Agile"),
    para("Soft Skills", { bold: true, size: 19, before: 150, after: 70 }),
    bullet("Leadership and Team Mentoring", { size: 18, after: 35 }),
    bullet("Problem-Solving and Debugging", { size: 18, after: 35 }),
    bullet("Communication and Collaboration", { size: 18, after: 35 }),
    bullet("Analytical Thinking - Analyzing business scenarios", { size: 18, after: 35 }),
    bullet("Time Management", { size: 18, after: 100 }),
    rightHeading("Professional Experience"),
    expHeader("EssilorLuxottica (via Talentium Inc.)", "Senior Analyst Programmer/Contractor", "07/2025 - Present | Filinvest City, Alabang Muntinlupa, Philippines"),
    bullet("Maintained Web, SOAP & REST APIs and Console apps", { size: 18, after: 35 }),
    bullet("Maintained middleware DLLs", { size: 18, after: 35 }),
  ].join("");
}

function pageTwoMain() {
  const idem = [
    "Developed Web, SOAP & REST APIs, Windows services and Console apps",
    "Developed Blazor Server application with SSO Integration",
    "Designed and implemented applications from ground-up",
    "Spearheaded the migration of legacy systems to .NET Core",
    "Maintained legacy systems",
    "Collaborated with cross-functional teams to deliver projects within Agile frameworks",
    "Integration with other application and third-party application",
    "Conducted demo/presentation of implemented projects to Product Owner",
    "Conducted Product support",
    "Conducted code reviews",
    "Mentored other devs",
  ];
  const smits = [
    "Developed Web, SOAP & REST APIs, Windows services and Console apps",
    "Integration with SAP",
    "Designed and implemented applications from ground-up",
    "Maintained legacy systems",
    "Collaborated with cross-functional teams to deliver projects within Agile frameworks",
    "Conducted demo/presentation of implemented projects to clients",
    "Conducted Product support",
    "Conducted code reviews",
    "Mentored other devs",
  ];
  const service101 = [
    "Developed Web, REST APIs and Mobile application",
    "Integration with SAP",
    "Conducted Product Support",
    "Conducted demo/presentation of implemented projects to clients",
  ];
  return [
    bullet("Maintained legacy systems", { size: 19, after: 45 }),
    bullet("Collaborated with cross-functional teams to deliver projects within Agile frameworks", { size: 19, after: 45 }),
    bullet("Integration with other application and third-party application", { size: 19, after: 45 }),
    bullet("Conducted demo/presentation of implemented projects to Product Owner", { size: 19, after: 45 }),
    bullet("Conducted Product support", { size: 19, after: 45 }),
    bullet("Conducted code reviews", { size: 19, after: 210 }),
    expHeader("IDEMIA Philippines Inc.", "Senior Software Developer", "09/2021 - 07/2025 | Makati, Philippines"),
    ...idem.map((item) => bullet(item, { size: 18, after: 35 })),
    para("", { after: 180 }),
    expHeader("SMITS Inc. - IT Company of San Miguel Corporation", ".Net Developer", "09/2018 - 08/2021 | Pasig, Philippines"),
    ...smits.map((item) => bullet(item, { size: 18, after: 35 })),
    para("", { after: 175 }),
    expHeader("Service 101 Plus Consulting Inc.", "Enterprise Systems Specialist", "03/2016 - 07/2018 | BGC Taguig, Philippines"),
    ...service101.map((item) => bullet(item, { size: 18, after: 35 })),
  ].join("");
}

function pageThreeMain() {
  return [
    expHeader("Infoman Inc.", "Web Developer (Intern)", "06/2015 - 09/2015 | Makati, Philippines"),
    bullet("Web Application Development", { size: 18, after: 40 }),
  ].join("");
}

function cell(content, width, fill, margins = { top: 720, left: 760, bottom: 360, right: 520 }) {
  return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="${fill}"/><w:tcMar><w:top w:w="${margins.top}" w:type="dxa"/><w:left w:w="${margins.left}" w:type="dxa"/><w:bottom w:w="${margins.bottom}" w:type="dxa"/><w:right w:w="${margins.right}" w:type="dxa"/></w:tcMar><w:vAlign w:val="top"/></w:tcPr>${content}</w:tc>`;
}

function pageTable(leftContent, rightContent) {
  return `<w:tbl><w:tblPr><w:tblW w:w="12240" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="5184"/><w:gridCol w:w="7056"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="15400" w:hRule="atLeast"/></w:trPr>${cell(leftContent, 5184, "0E3D52")}${cell(rightContent, 7056, "FFFFFF", { top: 740, left: 610, bottom: 360, right: 610 })}</w:tr></w:tbl>`;
}

function pageBreak() {
  return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
}

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}">
  <w:body>
    ${pageTable(sidePageOne(), pageOneMain())}
    ${pageBreak()}
    ${pageTable(sideBlank(), pageTwoMain())}
    ${pageBreak()}
    ${pageTable(sideBlank(), pageThreeMain())}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/>
      <w:cols w:space="720"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>`;

const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${NS.w}">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="20"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="80"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
</w:styles>`;

const settingsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="${NS.w}"><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/></w:settings>`;

const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;

const documentRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/profile_picture.jpg"/></Relationships>`;

const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`;

const files = [
  { name: "[Content_Types].xml", data: Buffer.from(contentTypesXml, "utf8") },
  { name: "_rels/.rels", data: Buffer.from(relsXml, "utf8") },
  { name: "word/document.xml", data: Buffer.from(documentXml, "utf8") },
  { name: "word/_rels/document.xml.rels", data: Buffer.from(documentRelsXml, "utf8") },
  { name: "word/styles.xml", data: Buffer.from(stylesXml, "utf8") },
  { name: "word/settings.xml", data: Buffer.from(settingsXml, "utf8") },
];

if (fs.existsSync(imagePath)) {
  files.push({ name: "word/media/profile_picture.jpg", data: fs.readFileSync(imagePath) });
}

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
