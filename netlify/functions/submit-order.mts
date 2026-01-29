import type { Context, Config } from "@netlify/functions";
import * as XLSX from 'xlsx';
import { S3Client, PutObjectCommand } from '@aws-sdk/client-s3';

interface TextLine {
  text: string;
  fontSize: number;
  x: number;
  y: number;
}

interface LabelColor {
  name: string;
  bg: string;
  text: string;
}

interface Label {
  id: number;
  height: number;
  width: number;
  font: string;
  textLines: TextLine[];
  color: LabelColor;
  corners: string;
  quantity: number;
}

interface OrderData {
  refId: string;
  contactName: string;
  contactEmail: string;
  labels: Label[];
}

export default async (req: Request, context: Context) => {
  // Handle CORS preflight
  if (req.method === "OPTIONS") {
    return new Response(null, { 
      status: 204,
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "POST, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type"
      }
    });
  }

  if (req.method !== "POST") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), { 
      status: 405,
      headers: { "Content-Type": "application/json" }
    });
  }

  try {
    const data: OrderData = await req.json();
    const { refId, contactName, contactEmail, labels } = data;

    if (!refId || !labels || labels.length === 0) {
      return new Response(JSON.stringify({ error: "Missing required fields" }), { 
        status: 400,
        headers: { "Content-Type": "application/json" }
      });
    }

    // Initialize S3 client
    const s3Client = new S3Client({
      region: Netlify.env.get("MY_AWS_REGION") || "us-east-1",
      credentials: {
        accessKeyId: Netlify.env.get("MY_AWS_ACCESS_KEY_ID") || "",
        secretAccessKey: Netlify.env.get("MY_AWS_SECRET_ACCESS_KEY") || ""
      }
    });
    const bucketName = Netlify.env.get("S3_BUCKET") || "matrix-systems-labels";

    const timestamp = new Date().toISOString();
    const formattedDate = new Date().toLocaleString("en-US", {
      year: "numeric",
      month: "short",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit"
    });

    // ============================================
    // BUILD EXCEL FILE
    // ============================================
    const excelData: any[][] = [];
    
    // Row 1: Reference ID
    excelData.push([`Reference ID: ${refId}`]);
    
    // Row 2: Empty row for spacing
    excelData.push([]);
    
    // Row 3: Headers - Height, Width, Color, VAR1-6, VAR1-6 Sizes, Font, Corners
    const headers = [
      "Height", "Width", "Color",
      "VAR1", "VAR2", "VAR3", "VAR4", "VAR5", "VAR6",
      "VAR1 Size", "VAR2 Size", "VAR3 Size", "VAR4 Size", "VAR5 Size", "VAR6 Size",
      "Font", "Corners"
    ];
    excelData.push(headers);
    
    // Data rows - one row per label quantity
    let totalLabels = 0;
    labels.forEach((label: Label) => {
      // Pad textLines to 6 elements
      const lines = [...label.textLines];
      while (lines.length < 6) {
        lines.push({ text: '', fontSize: 0, x: 50, y: 50 });
      }
      
      const row = [
        label.height,
        label.width,
        label.color.name,
        lines[0]?.text || "",
        lines[1]?.text || "",
        lines[2]?.text || "",
        lines[3]?.text || "",
        lines[4]?.text || "",
        lines[5]?.text || "",
        lines[0]?.text ? lines[0].fontSize : "",
        lines[1]?.text ? lines[1].fontSize : "",
        lines[2]?.text ? lines[2].fontSize : "",
        lines[3]?.text ? lines[3].fontSize : "",
        lines[4]?.text ? lines[4].fontSize : "",
        lines[5]?.text ? lines[5].fontSize : "",
        label.font,
        label.corners
      ];
      
      // Repeat row for quantity
      for (let i = 0; i < label.quantity; i++) {
        excelData.push([...row]);
        totalLabels++;
      }
    });

    // Create workbook and worksheet
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(excelData);
    
    // Set column widths
    worksheet['!cols'] = [
      { wch: 8 },   // Height
      { wch: 8 },   // Width
      { wch: 14 },  // Color
      { wch: 20 },  // VAR1
      { wch: 20 },  // VAR2
      { wch: 20 },  // VAR3
      { wch: 20 },  // VAR4
      { wch: 20 },  // VAR5
      { wch: 20 },  // VAR6
      { wch: 10 },  // VAR1 Size
      { wch: 10 },  // VAR2 Size
      { wch: 10 },  // VAR3 Size
      { wch: 10 },  // VAR4 Size
      { wch: 10 },  // VAR5 Size
      { wch: 10 },  // VAR6 Size
      { wch: 12 },  // Font
      { wch: 10 }   // Corners
    ];

    XLSX.utils.book_append_sheet(workbook, worksheet, "Nameplate Labels");
    const xlsxBuffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

    // ============================================
    // BUILD HTML FOR PDF CONVERSION
    // ============================================
    const htmlContent = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Nameplate Labels - ${refId}</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { 
      font-family: Arial, sans-serif; 
      padding: 40px;
      background: #fff;
    }
    .header {
      margin-bottom: 30px;
      padding-bottom: 20px;
      border-bottom: 2px solid #1a365d;
    }
    .header h1 {
      color: #1a365d;
      font-size: 24px;
      margin-bottom: 8px;
    }
    .header-info {
      color: #64748b;
      font-size: 14px;
    }
    .header-info span {
      margin-right: 20px;
    }
    .summary {
      background: #f8fafc;
      padding: 16px 20px;
      border-radius: 8px;
      margin-bottom: 30px;
      display: flex;
      gap: 40px;
    }
    .summary-item {
      display: flex;
      flex-direction: column;
    }
    .summary-label {
      font-size: 12px;
      color: #64748b;
      text-transform: uppercase;
    }
    .summary-value {
      font-size: 20px;
      font-weight: 600;
      color: #1a365d;
    }
    .labels-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
      gap: 20px;
    }
    .label-card {
      border: 1px solid #e2e8f0;
      border-radius: 8px;
      overflow: hidden;
    }
    .label-preview {
      padding: 20px;
      display: flex;
      align-items: center;
      justify-content: center;
      min-height: 80px;
    }
    .label-visual {
      position: relative;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
    }
    .label-visual.rounded {
      border-radius: 6px;
    }
    .label-text {
      position: absolute;
      white-space: nowrap;
      transform: translate(-50%, -50%);
    }
    .label-details {
      padding: 16px;
      background: #f8fafc;
      border-top: 1px solid #e2e8f0;
    }
    .label-details-row {
      display: flex;
      justify-content: space-between;
      margin-bottom: 6px;
      font-size: 13px;
    }
    .label-details-row:last-child {
      margin-bottom: 0;
    }
    .detail-label {
      color: #64748b;
    }
    .detail-value {
      color: #1a365d;
      font-weight: 500;
    }
    .qty-badge {
      display: inline-block;
      background: #2563eb;
      color: white;
      padding: 2px 8px;
      border-radius: 4px;
      font-size: 12px;
      font-weight: 600;
    }
    .footer {
      margin-top: 40px;
      padding-top: 20px;
      border-top: 1px solid #e2e8f0;
      text-align: center;
      color: #94a3b8;
      font-size: 12px;
    }
  </style>
</head>
<body>
  <div class="header">
    <h1>Nameplate Label Quote Request</h1>
    <div class="header-info">
      <span><strong>Reference:</strong> ${refId}</span>
      <span><strong>Date:</strong> ${formattedDate}</span>
      ${contactName ? `<span><strong>Contact:</strong> ${contactName}</span>` : ''}
      ${contactEmail ? `<span><strong>Email:</strong> ${contactEmail}</span>` : ''}
    </div>
  </div>

  <div class="summary">
    <div class="summary-item">
      <span class="summary-label">Total Labels</span>
      <span class="summary-value">${totalLabels}</span>
    </div>
    <div class="summary-item">
      <span class="summary-label">Unique Designs</span>
      <span class="summary-value">${labels.length}</span>
    </div>
  </div>

  <div class="labels-grid">
    ${labels.map((label: Label) => {
      const scale = Math.min(240 / (label.width * 72), 60 / (label.height * 72));
      const width = label.width * 72 * scale;
      const height = label.height * 72 * scale;
      const borderRadius = label.corners === 'rounded' ? '6px' : '0';
      
      const textElements = label.textLines
        .filter(line => line.text)
        .map(line => {
          const fontSize = Math.max(8, line.fontSize * scale);
          return `<div class="label-text" style="left: ${line.x}%; top: ${line.y}%; font-family: ${label.font}, sans-serif; font-size: ${fontSize}px; color: ${label.color.text};">${escapeHtml(line.text)}</div>`;
        })
        .join('');
      
      const primaryText = label.textLines.find(l => l.text)?.text || 'Empty';
      
      return `
        <div class="label-card">
          <div class="label-preview">
            <div class="label-visual" style="width: ${width}px; height: ${height}px; background: ${label.color.bg}; border-radius: ${borderRadius}; ${label.color.bg === '#FFFFFF' ? 'border: 1px solid #e2e8f0;' : ''}">
              ${textElements}
            </div>
          </div>
          <div class="label-details">
            <div class="label-details-row">
              <span class="detail-label">Size</span>
              <span class="detail-value">${label.height}" × ${label.width}"</span>
            </div>
            <div class="label-details-row">
              <span class="detail-label">Color</span>
              <span class="detail-value">${label.color.name}</span>
            </div>
            <div class="label-details-row">
              <span class="detail-label">Font</span>
              <span class="detail-value">${label.font}</span>
            </div>
            <div class="label-details-row">
              <span class="detail-label">Corners</span>
              <span class="detail-value">${label.corners}</span>
            </div>
            <div class="label-details-row">
              <span class="detail-label">Quantity</span>
              <span class="qty-badge">×${label.quantity}</span>
            </div>
          </div>
        </div>
      `;
    }).join('')}
  </div>

  <div class="footer">
    Generated by Matrix Systems Nameplate Label Creator
  </div>
</body>
</html>`;

    // ============================================
    // UPLOAD TO S3
    // ============================================
    const xlsxKey = `nameplates/${refId}-${Date.now()}.xlsx`;
    const htmlKey = `nameplates/${refId}-${Date.now()}.html`;

    // Upload Excel
    await s3Client.send(new PutObjectCommand({
      Bucket: bucketName,
      Key: xlsxKey,
      Body: new Uint8Array(xlsxBuffer),
      ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }));

    // Upload HTML
    await s3Client.send(new PutObjectCommand({
      Bucket: bucketName,
      Key: htmlKey,
      Body: htmlContent,
      ContentType: "text/html"
    }));

    const xlsxUrl = `https://${bucketName}.s3.amazonaws.com/${xlsxKey}`;
    const htmlUrl = `https://${bucketName}.s3.amazonaws.com/${htmlKey}`;

    // ============================================
    // SEND TO ZAPIER WEBHOOK
    // ============================================
    const labelSummaries = labels.map((label: Label) => ({
      height: label.height,
      width: label.width,
      color: label.color.name,
      var1: label.textLines[0]?.text || "",
      var2: label.textLines[1]?.text || "",
      var3: label.textLines[2]?.text || "",
      var4: label.textLines[3]?.text || "",
      var5: label.textLines[4]?.text || "",
      var6: label.textLines[5]?.text || "",
      font: label.font,
      corners: label.corners,
      quantity: label.quantity
    }));

    const webhookUrl = Netlify.env.get("ZAPIER_WEBHOOK_URL") || "https://hooks.zapier.com/hooks/catch/24455310/ul65avn/";
    
    const webhookPayload = {
      refId,
      contactName: contactName || "",
      contactEmail: contactEmail || "",
      timestamp,
      formattedDate,
      totalLabels,
      labelCount: labels.length,
      xlsxUrl,
      xlsxFileName: `nameplates-${refId}.xlsx`,
      htmlUrl,
      pdfFileName: `nameplates-${refId}.pdf`,
      labels: labelSummaries
    };

    const webhookResponse = await fetch(webhookUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(webhookPayload)
    });

    if (!webhookResponse.ok) {
      console.error("Webhook failed:", await webhookResponse.text());
    }

    return new Response(JSON.stringify({ 
      success: true, 
      message: "Order submitted successfully",
      refId,
      totalLabels
    }), {
      status: 200,
      headers: { "Content-Type": "application/json" }
    });

  } catch (error) {
    console.error("Error processing order:", error);
    return new Response(JSON.stringify({ error: "Failed to process order" }), { 
      status: 500,
      headers: { "Content-Type": "application/json" }
    });
  }
};

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

export const config: Config = {
  path: "/api/submit-order"
};
