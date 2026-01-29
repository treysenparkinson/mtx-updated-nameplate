import type { Context } from "@netlify/functions";
import { S3Client, PutObjectCommand } from "@aws-sdk/client-s3";
import * as XLSX from "xlsx";

const s3 = new S3Client({
  region: process.env.MY_AWS_REGION || "us-east-1",
  credentials: {
    accessKeyId: process.env.MY_AWS_ACCESS_KEY_ID || "",
    secretAccessKey: process.env.MY_AWS_SECRET_ACCESS_KEY || "",
  },
});

const BUCKET = process.env.S3_BUCKET || "matrix-systems-labels";
const ZAPIER_WEBHOOK = process.env.ZAPIER_WEBHOOK_URL || "";

interface TextLine {
  text: string;
  fontSize: number;
  x: number;
  y: number;
}

interface Label {
  id: number;
  height: number;
  width: number;
  font: string;
  textLines: TextLine[];
  labelColor: string;
  textColor: string;
  corners: string;
  quantity: number;
}

interface RequestBody {
  refId: string;
  contactName: string;
  contactEmail: string;
  notes: string;
  labels: Label[];
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function getColorName(hex: string): string {
  const colors: Record<string, string> = {
    "#22c55e": "Green",
    "#ef4444": "Red",
    "#eab308": "Yellow",
    "#3b82f6": "Blue",
    "#1a1a1a": "Black",
    "#ffffff": "White",
    "#f97316": "Orange",
    "#6b7280": "Gray",
  };
  return colors[hex.toLowerCase()] || hex;
}

export default async (req: Request, _context: Context) => {
  if (req.method !== "POST") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), {
      status: 405,
      headers: { "Content-Type": "application/json" },
    });
  }

  try {
    const body: RequestBody = await req.json();
    const { refId, contactName, contactEmail, notes, labels } = body;

    if (!refId || !labels || labels.length === 0) {
      return new Response(
        JSON.stringify({ error: "Missing required fields" }),
        { status: 400, headers: { "Content-Type": "application/json" } }
      );
    }

    const timestamp = Date.now();
    const formattedDate = new Date().toLocaleString("en-US", {
      timeZone: "America/Chicago",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
    });

    // Generate Excel file
    const workbook = XLSX.utils.book_new();
    
    // Build rows - one row per quantity
    const rows: string[][] = [];
    
    // Header row
    rows.push([`Reference ID: ${refId}`]);
    rows.push([]); // Empty row
    rows.push([
      "Height",
      "Width",
      "Label Color",
      "Text Color",
      "Font",
      "Corners",
      "Text 1",
      "Size 1",
      "Text 2",
      "Size 2",
      "Text 3",
      "Size 3",
      "Text 4",
      "Size 4",
      "Text 5",
      "Size 5",
      "Text 6",
      "Size 6",
      "Pos X 1",
      "Pos Y 1",
      "Pos X 2",
      "Pos Y 2",
      "Pos X 3",
      "Pos Y 3",
    ]);

    // Data rows
    for (const label of labels) {
      for (let i = 0; i < label.quantity; i++) {
        const row: string[] = [
          label.height.toString(),
          label.width.toString(),
          getColorName(label.labelColor),
          getColorName(label.textColor),
          label.font,
          label.corners,
        ];

        // Add text lines (up to 6)
        for (let j = 0; j < 6; j++) {
          const line = label.textLines[j];
          row.push(line?.text || "");
          row.push(line?.fontSize?.toString() || "");
        }

        // Add positions (up to 3)
        for (let j = 0; j < 3; j++) {
          const line = label.textLines[j];
          row.push(line?.x?.toFixed(1) || "");
          row.push(line?.y?.toFixed(1) || "");
        }

        rows.push(row);
      }
    }

    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    
    // Set column widths
    worksheet["!cols"] = [
      { wch: 8 },  // Height
      { wch: 8 },  // Width
      { wch: 12 }, // Label Color
      { wch: 12 }, // Text Color
      { wch: 15 }, // Font
      { wch: 10 }, // Corners
      { wch: 20 }, // Text 1
      { wch: 8 },  // Size 1
      { wch: 20 }, // Text 2
      { wch: 8 },  // Size 2
      { wch: 20 }, // Text 3
      { wch: 8 },  // Size 3
    ];

    XLSX.utils.book_append_sheet(workbook, worksheet, "Labels");
    const xlsxBuffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

    // Generate HTML for PDF conversion
    const totalLabels = labels.reduce((sum, l) => sum + l.quantity, 0);
    const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Nameplate Label Order - ${refId}</title>
  <style>
    body { font-family: Arial, sans-serif; padding: 40px; color: #333; }
    .header { border-bottom: 2px solid #10b981; padding-bottom: 20px; margin-bottom: 30px; }
    .header h1 { margin: 0; color: #1a365d; }
    .header p { margin: 5px 0; color: #666; }
    .summary { display: flex; gap: 40px; margin-bottom: 30px; }
    .summary-item { background: #f8fafc; padding: 15px 20px; border-radius: 8px; }
    .summary-item label { font-size: 12px; color: #64748b; display: block; }
    .summary-item value { font-size: 18px; font-weight: bold; color: #1a365d; }
    .labels-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 20px; }
    .label-card { border: 1px solid #e2e8f0; border-radius: 8px; padding: 15px; }
    .label-preview { width: 100%; height: 60px; border-radius: 4px; display: flex; align-items: center; justify-content: center; margin-bottom: 10px; font-size: 14px; }
    .label-details { font-size: 12px; color: #64748b; }
    .label-details p { margin: 4px 0; }
    .notes { margin-top: 30px; padding: 20px; background: #fffbeb; border-radius: 8px; }
    .notes h3 { margin: 0 0 10px; color: #92400e; }
  </style>
</head>
<body>
  <div class="header">
    <h1>Nameplate Label Order</h1>
    <p><strong>Reference ID:</strong> ${escapeHtml(refId)}</p>
    <p><strong>Date:</strong> ${formattedDate}</p>
    ${contactName ? `<p><strong>Contact:</strong> ${escapeHtml(contactName)}</p>` : ""}
    ${contactEmail ? `<p><strong>Email:</strong> ${escapeHtml(contactEmail)}</p>` : ""}
  </div>
  
  <div class="summary">
    <div class="summary-item">
      <label>Total Labels</label>
      <value>${totalLabels}</value>
    </div>
    <div class="summary-item">
      <label>Unique Designs</label>
      <value>${labels.length}</value>
    </div>
  </div>

  <div class="labels-grid">
    ${labels
      .map(
        (label) => `
      <div class="label-card">
        <div class="label-preview" style="background:${label.labelColor};color:${label.textColor};${label.labelColor.toUpperCase() === '#FFFFFF' ? 'border:1px solid #ddd;' : ''}">
          ${label.textLines
            .filter((l) => l.text)
            .map((l) => escapeHtml(l.text))
            .join(" | ")}
        </div>
        <div class="label-details">
          <p><strong>Size:</strong> ${label.width}" × ${label.height}"</p>
          <p><strong>Colors:</strong> ${getColorName(label.labelColor)} / ${getColorName(label.textColor)}</p>
          <p><strong>Font:</strong> ${label.font}</p>
          <p><strong>Corners:</strong> ${label.corners}</p>
          <p><strong>Quantity:</strong> ${label.quantity}</p>
        </div>
      </div>
    `
      )
      .join("")}
  </div>

  ${notes ? `<div class="notes"><h3>Notes</h3><p>${escapeHtml(notes)}</p></div>` : ""}
</body>
</html>`;

    // Upload to S3
    const xlsxKey = `nameplates/${refId}-${timestamp}.xlsx`;
    const htmlKey = `nameplates/${refId}-${timestamp}.html`;

    await s3.send(
      new PutObjectCommand({
        Bucket: BUCKET,
        Key: xlsxKey,
        Body: xlsxBuffer,
        ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      })
    );

    await s3.send(
      new PutObjectCommand({
        Bucket: BUCKET,
        Key: htmlKey,
        Body: html,
        ContentType: "text/html",
      })
    );

    const xlsxUrl = `https://${BUCKET}.s3.amazonaws.com/${xlsxKey}`;
    const htmlUrl = `https://${BUCKET}.s3.amazonaws.com/${htmlKey}`;

    // Send to Zapier
    if (ZAPIER_WEBHOOK) {
      await fetch(ZAPIER_WEBHOOK, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          refId,
          contactName,
          contactEmail,
          notes,
          timestamp,
          formattedDate,
          totalLabels,
          labelCount: labels.length,
          xlsxUrl,
          xlsxFileName: `${refId}-${timestamp}.xlsx`,
          htmlUrl,
          pdfFileName: `${refId}-${timestamp}.html`,
          labels: labels.map((l) => ({
            height: l.height,
            width: l.width,
            labelColor: getColorName(l.labelColor),
            textColor: getColorName(l.textColor),
            font: l.font,
            corners: l.corners,
            quantity: l.quantity,
            textLines: l.textLines.filter((t) => t.text).map((t) => ({
              text: t.text,
              fontSize: t.fontSize,
              x: t.x,
              y: t.y,
            })),
          })),
        }),
      });
    }

    return new Response(
      JSON.stringify({
        success: true,
        refId,
        xlsxUrl,
        htmlUrl,
      }),
      { status: 200, headers: { "Content-Type": "application/json" } }
    );
  } catch (error) {
    console.error("Submit order error:", error);
    return new Response(
      JSON.stringify({ error: "Failed to process order", details: String(error) }),
      { status: 500, headers: { "Content-Type": "application/json" } }
    );
  }
};
