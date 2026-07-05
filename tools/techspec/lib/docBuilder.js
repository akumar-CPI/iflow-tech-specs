"use strict";
// Node port of src/docBuilder.js — identical document structure/content,
// swapped from the browser `docx` global to `require("docx")`, and from
// Packer.toBlob to Packer.toBuffer.

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, ImageRun, AlignmentType, VerticalAlign,
} = require("docx");

const NAVY = "1F3864";
const HEAD2_BLUE = "2F5496";
const HEAD3_GREEN = "375623";
const BORDER_GREY = "CCCCCC";

const cellBorders = {
  top: { style: BorderStyle.SINGLE, size: 4, color: BORDER_GREY },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: BORDER_GREY },
  left: { style: BorderStyle.SINGLE, size: 4, color: BORDER_GREY },
  right: { style: BorderStyle.SINGLE, size: 4, color: BORDER_GREY },
};

const title = (text) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 0, after: 360 },
  children: [new TextRun({ text, bold: true, size: 52, color: NAVY })],
});
const h2 = (text) => new Paragraph({
  spacing: { before: 280, after: 80 },
  children: [new TextRun({ text, bold: true, size: 32, color: HEAD2_BLUE })],
});
const h3 = (text) => new Paragraph({
  spacing: { before: 200, after: 60 },
  children: [new TextRun({ text, bold: true, size: 26, color: HEAD3_GREEN })],
});
const body = (text, opts = {}) => new Paragraph({
  spacing: { after: 120 },
  children: [new TextRun({ text: text || "", size: 22, italics: !!opts.italics })],
});
const spacer = () => new Paragraph({ spacing: { after: 80 }, children: [] });
const bullet = (text) => new Paragraph({
  spacing: { after: 60 }, indent: { left: 360 },
  children: [new TextRun({ text: `\u2022 ${text}`, size: 22 })],
});
const tocItem = (num, text) => new Paragraph({
  indent: { left: 720, hanging: 360 }, spacing: { after: 80 },
  children: [new TextRun({ text: `${num}. `, bold: true, size: 22 }), new TextRun({ text, size: 22 })],
});

function headerCell(text, w) {
  return new TableCell({
    width: { size: w, type: WidthType.PERCENTAGE },
    borders: cellBorders,
    shading: { type: ShadingType.CLEAR, color: "auto", fill: NAVY },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: "FFFFFF", size: 18 })] })],
  });
}
function dataCell(text, w) {
  return new TableCell({
    width: { size: w, type: WidthType.PERCENTAGE },
    borders: cellBorders,
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [new TextRun({ text: String(text ?? ""), size: 18 })] })],
  });
}
function makeTable(headers, rows, widths) {
  const w = widths || headers.map(() => Math.floor(100 / headers.length));
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((h, i) => headerCell(h, w[i])) }),
      ...rows.map((r) => new TableRow({ children: r.map((c, i) => dataCell(c, w[i])) })),
    ],
  });
}
function imageParagraph(buffer, w, h) {
  return new Paragraph({ children: [new ImageRun({ data: buffer, transformation: { width: w, height: h }, type: "png" })] });
}

function humanAuthMethod(props) {
  return props.authenticationMethod || props.authentication || "\u2014";
}
function humanEndpoint(props) {
  return props.httpAddressWithoutQuery || (props.host ? `${props.host}${props.path || ""}` : "\u2014");
}
function stepRows(steps) {
  return steps.map((s, i) => [String(i + 1), s.name || "(unnamed)", s.type, s.aiDescription || ""]);
}

async function build({ parsed, ai, diagrams, docTitle }) {
  const toc = [
    "Overview", "Interface Overview", "High-Level iFlow Design", "Message Flow Diagram",
    "Externalized Parameters", "Dependencies & Referenced Artifacts", "Deployment Configuration Checklist",
    "Sender & Receiver Channels", "Mappings & Transformations", "Security Configuration",
    "Error Handling & Logging", "Test Plan", "Version & Metadata", "Review & Recommendations",
  ];

  const children = [];
  const finalTitle = docTitle && docTitle.trim() ? docTitle.trim() : `${parsed.iflowName} — Technical Specification`;
  children.push(title(finalTitle));

  children.push(h2("Change History"));
  const today = new Date().toISOString().slice(0, 10);
  children.push(makeTable(
    ["Version", "Date", "Author", "Change Description"],
    [["1.0", today, "CI Pipeline", "Auto-generated from iFlow artifact via GitHub Actions"]],
    [12, 15, 20, 53]
  ));
  children.push(spacer(), spacer());

  children.push(h2("Table of Contents"));
  toc.forEach((t, i) => children.push(tocItem(i + 1, t)));
  children.push(spacer());

  children.push(h2("1. Overview"));
  children.push(body(ai.overview));
  children.push(spacer());

  children.push(h2("2. Interface Overview"));
  const ifRows = parsed.messageFlows.map((mf) => [
    mf.props.direction || "\u2014",
    mf.name || mf.props.system || "\u2014",
    mf.props.ComponentType || mf.props.TransportProtocol || "\u2014",
    humanAuthMethod(mf.props),
    humanEndpoint(mf.props),
  ]);
  if (ifRows.length) {
    children.push(makeTable(
      ["Direction", "System / Partner", "Adapter Type", "Authentication", "Endpoint / Address"],
      ifRows, [20, 24, 14, 20, 22]
    ));
  } else {
    children.push(body("No sender/receiver message flows were found in this iFlow.", { italics: true }));
  }
  children.push(spacer(), spacer());

  children.push(h2("3. High-Level iFlow Design"));
  if (diagrams.highLevel) {
    children.push(h2("High-Level Design"));
    children.push(imageParagraph(diagrams.highLevel.buffer, 555, Math.round(555 * (diagrams.highLevel.height / diagrams.highLevel.width))));
  }
  if (diagrams.detailed) {
    children.push(h2("Message Flow Diagram"));
    children.push(imageParagraph(diagrams.detailed.buffer, 555, Math.round(555 * (diagrams.detailed.height / diagrams.detailed.width))));
  }
  children.push(body(ai.highLevelSummary));
  children.push(spacer());

  children.push(h2("4. Message Flow Diagram"));
  children.push(spacer());
  parsed.processes.forEach((proc, pi) => {
    children.push(h3(`4.${pi + 1} ${proc.name || "Integration Process"}`));
    const rows = stepRows(proc.steps);
    if (rows.length) {
      children.push(makeTable(["Step #", "Step Name", "Type", "Description"], rows, [8, 24, 20, 48]));
    } else {
      children.push(body("No steps found in this process.", { italics: true }));
    }
    children.push(spacer());
    (proc.subProcesses || []).forEach((sub, si) => {
      children.push(h3(`4.${pi + 1}.${si + 1} ${sub.name || "Subprocess"}`));
      const subRows = stepRows(sub.steps);
      if (subRows.length) {
        children.push(makeTable(["Step #", "Step Name", "Type", "Description"], subRows, [8, 24, 20, 48]));
      }
      children.push(spacer());
    });
  });

  children.push(h2("5. Externalized Parameters"));
  if (parsed.parametersDefined && parsed.parametersDefined.length) {
    children.push(makeTable(["Parameter Name"], parsed.parametersDefined.map((p) => [p.name]), [100]));
  } else {
    children.push(body("No externalized parameters are defined in this iFlow.", { italics: true }));
  }
  children.push(spacer());

  children.push(h2("6. Dependencies & Referenced Artifacts"));
  children.push(body(ai.dependenciesNote));
  children.push(spacer());

  children.push(h2("7. Deployment Configuration Checklist"));
  (ai.deploymentChecklist || []).forEach((item) => children.push(bullet(item)));
  children.push(spacer());

  children.push(h2("8. Sender & Receiver Channels"));
  const chanRows = parsed.messageFlows.map((mf) => {
    const p = mf.props;
    const details = [
      p.TransportProtocolVersion ? `version ${p.TransportProtocolVersion}` : null,
      p.httpMethod ? `method ${p.httpMethod}` : null,
      p.httpRequestTimeout ? `timeout ${p.httpRequestTimeout} ms` : null,
      p.connectTimeout ? `connect timeout ${p.connectTimeout} ms` : null,
      p.maximumReconnectAttempts ? `max reconnects ${p.maximumReconnectAttempts}` : null,
      p.authenticationMethod ? `auth ${p.authenticationMethod}` : (p.authentication ? `auth ${p.authentication}` : null),
    ].filter(Boolean).join(", ");
    return [mf.name || "\u2014", p.direction || "\u2014", p.ComponentType || p.TransportProtocol || "\u2014", details || "\u2014"];
  });
  if (chanRows.length) {
    children.push(makeTable(["Channel", "Direction", "Adapter", "Key Settings"], chanRows, [16, 20, 18, 46]));
  } else {
    children.push(body("No adapter channels detected.", { italics: true }));
  }
  children.push(spacer(), spacer());

  children.push(h2("9. Mappings & Transformations"));
  children.push(body(ai.mappingsNote));
  (ai.mappingsBullets || []).forEach((b) => children.push(bullet(b)));
  children.push(spacer());

  children.push(h2("10. Security Configuration"));
  if (ai.securityRows && ai.securityRows.length) {
    children.push(makeTable(["Area", "Configuration", "Observation"], ai.securityRows, [20, 38, 42]));
  }
  children.push(spacer(), spacer());

  children.push(h2("11. Error Handling & Logging"));
  (ai.errorHandlingBullets || []).forEach((b) => children.push(bullet(b)));
  children.push(spacer());

  children.push(h2("12. Test Plan"));
  children.push(body("Suggested test scenarios derived from the iFlow structure (extend with business-specific cases):"));
  children.push(spacer());
  if (ai.testPlanRows && ai.testPlanRows.length) {
    children.push(makeTable(["#", "Test Scenario", "Type", "Expected Result"], ai.testPlanRows, [6, 32, 18, 44]));
  }
  children.push(spacer(), spacer());

  children.push(h2("13. Version & Metadata"));
  children.push(makeTable(
    ["Field", "Value"],
    [
      ["iFlow Name", parsed.iflowName || "\u2014"],
      ["Bundle Version", parsed.bundleVersion || "\u2014"],
      ["Source", "Auto-generated by CI pipeline (GitHub Actions)"],
      ["Generated On", today],
      ["Status", "Draft — review before publishing"],
    ],
    [35, 65]
  ));
  children.push(spacer(), spacer());

  children.push(h2("14. Review & Recommendations"));
  children.push(body("Automated best-practice review (verify against your project standards):"));
  children.push(spacer());
  if (ai.reviewRows && ai.reviewRows.length) {
    children.push(makeTable(["Severity", "Area", "Finding", "Recommendation"], ai.reviewRows, [10, 14, 38, 38]));
  }
  children.push(spacer(), spacer());

  children.push(h2("Appendix"));
  children.push(spacer());
  (ai.appendixBullets || []).forEach((b) => children.push(bullet(b)));
  if (!ai.appendixBullets || !ai.appendixBullets.length) {
    children.push(body("List any assumptions, open points, reference documents, or SAP Notes relevant to this integration.", { italics: true }));
  }

  const doc = new Document({
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      children,
    }],
  });

  return Packer.toBuffer(doc);
}

module.exports = { build };
