import { S3Client, PutObjectCommand } from "@aws-sdk/client-s3";
import * as XLSX from "xlsx";

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
  stickyBack: boolean;
  quantity: number;
}

interface RequestBody {
  refId: string;
  contactName: string;
  contactEmail: string;
  notes: string;
  labels: Label[];
}

const getColorName = (hex: string): string => {
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
};

const escapeHtml = (text: string): string => {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
};

export default async (req: Request) => {
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

    // Check environment variables
    const awsKey = process.env.MY_AWS_ACCESS_KEY_ID;
    const awsSecret = process.env.MY_AWS_SECRET_ACCESS_KEY;
    const awsRegion = process.env.MY_AWS_REGION || "us-east-1";
    const bucket = process.env.S3_BUCKET || "matrix-systems-labels";
    const zapierWebhook = process.env.ZAPIER_WEBHOOK_URL;

    if (!awsKey || !awsSecret) {
      console.error("Missing AWS credentials");
      return new Response(
        JSON.stringify({ error: "Server configuration error" }),
        { status: 500, headers: { "Content-Type": "application/json" } }
      );
    }

    const s3 = new S3Client({
      region: awsRegion,
      credentials: {
        accessKeyId: awsKey,
        secretAccessKey: awsSecret,
      },
    });

    const timestamp = Date.now();
    const formattedDate = new Date().toLocaleString("en-US", {
      timeZone: "America/Chicago",
      year: "numeric",
      month: "short",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    });

    // Find max text lines across all labels
    const maxTextLines = Math.max(...labels.map(l => l.textLines.filter(t => t.text).length), 1);

    // Build Excel headers dynamically
    const headers: string[] = ["ID #"];
    for (let i = 1; i <= maxTextLines; i++) {
      headers.push(`LINE ${i} TEXT`);
      headers.push(`LINE ${i} TEXT SIZE`);
    }
    headers.push("BACKGRND COLOR", "LETTER COLOR", "WIDTH (INCHES)", "HEIGHT (INCHES)", "CORNERS", "STICKY BACK", "COMMENTS");

    // Build Excel rows - one row per quantity
    const rows: (string | number)[][] = [headers];

    for (const label of labels) {
      for (let q = 0; q < label.quantity; q++) {
        const row: (string | number)[] = [refId];
        
        // Add text lines and sizes
        for (let i = 0; i < maxTextLines; i++) {
          const line = label.textLines[i];
          row.push(line?.text || "");
          row.push(line?.fontSize || "");
        }
        
        // Add remaining columns
        row.push(
          getColorName(label.labelColor),
          getColorName(label.textColor),
          label.width,
          label.height,
          label.corners.charAt(0).toUpperCase() + label.corners.slice(1),
          label.stickyBack ? "Yes" : "No",
          notes || ""
        );
        
        rows.push(row);
      }
    }

    // Create Excel workbook
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    
    // Set column widths
    const colWidths = [{ wch: 15 }]; // ID #
    for (let i = 0; i < maxTextLines; i++) {
      colWidths.push({ wch: 20 }); // TEXT
      colWidths.push({ wch: 12 }); // SIZE
    }
    colWidths.push(
      { wch: 14 }, // BACKGRND COLOR
      { wch: 14 }, // LETTER COLOR
      { wch: 14 }, // WIDTH
      { wch: 14 }, // HEIGHT
      { wch: 10 }, // CORNERS
      { wch: 12 }, // STICKY BACK
      { wch: 30 }  // COMMENTS
    );
    worksheet["!cols"] = colWidths;

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Labels");
    const xlsxBuffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

    // Calculate totals for PDF
    const totalLabels = labels.reduce((sum, l) => sum + l.quantity, 0);
    const totalPages = Math.ceil(labels.length / 10); // 10 labels per page

    // Generate HTML/PDF
    const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Nameplate Labels - ${escapeHtml(refId)}</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { font-family: Arial, sans-serif; padding: 40px; color: #333; background: #fff; }
    .header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 10px; }
    .header-left { }
    .header-right { text-align: right; color: #666; font-size: 13px; }
    .ref-id { font-size: 14px; font-weight: bold; color: #333; }
    .title { font-size: 28px; font-weight: bold; color: #000; margin: 5px 0 20px 0; }
    .page-info { font-size: 13px; color: #666; }
    
    .table { width: 100%; border-collapse: collapse; margin-top: 10px; }
    .table th { text-align: left; padding: 12px 15px; border-bottom: 2px solid #333; font-size: 13px; color: #666; font-weight: normal; }
    .table td { padding: 15px; border-bottom: 1px solid #eee; vertical-align: middle; }
    .table tr:last-child td { border-bottom: 2px solid #333; }
    
    .preview-cell { width: 120px; }
    .label-preview { 
      width: 100px; 
      height: 60px; 
      border-radius: 4px; 
      display: flex; 
      flex-direction: column;
      align-items: center; 
      justify-content: center; 
      font-size: 10px;
      overflow: hidden;
      border: 1px solid #ddd;
    }
    .label-preview.rounded { border-radius: 8px; }
    .label-preview.squared { border-radius: 0; }
    .label-preview-text { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 90px; }
    
    .size-name { }
    .label-name { font-size: 15px; font-weight: 600; color: #333; }
    .label-size { font-size: 13px; color: #666; margin-top: 3px; }
    
    .font-cell { font-size: 13px; color: #333; }
    .qty-cell { font-size: 15px; font-weight: bold; color: #333; }
    
    .contact-info { margin-top: 30px; padding: 20px; background: #f8f8f8; border-radius: 8px; }
    .contact-info h3 { font-size: 14px; margin-bottom: 10px; }
    .contact-info p { font-size: 13px; color: #666; margin: 5px 0; }
    
    .notes-section { margin-top: 20px; padding: 20px; background: #fffbeb; border-radius: 8px; }
    .notes-section h3 { font-size: 14px; margin-bottom: 10px; color: #92400e; }
    .notes-section p { font-size: 13px; color: #666; }

    .sticky-badge { display: inline-block; background: #dbeafe; color: #1d4ed8; font-size: 9px; padding: 2px 6px; border-radius: 4px; margin-top: 4px; }
  </style>
</head>
<body>
  <div class="header">
    <div class="header-left">
      <div class="ref-id">Reference ID: ${escapeHtml(refId)}</div>
      <div class="title">Saved Labels Summary</div>
    </div>
    <div class="header-right">
      <div>${formattedDate}</div>
      <div class="page-info">Page 1 of ${totalPages}</div>
    </div>
  </div>

  <table class="table">
    <thead>
      <tr>
        <th>Preview</th>
        <th>Size/Name</th>
        <th>Font</th>
        <th style="text-align: right;">Qty</th>
      </tr>
    </thead>
    <tbody>
      ${labels.map(label => {
        const textContent = label.textLines.filter(t => t.text).map(t => t.text);
        const primaryText = textContent[0] || "—";
        const previewLines = textContent.slice(0, 3);
        
        return `
          <tr>
            <td class="preview-cell">
              <div class="label-preview ${label.corners}" style="background:${label.labelColor};color:${label.textColor};${label.labelColor.toUpperCase() === '#FFFFFF' ? 'border:1px solid #ccc;' : ''}">
                ${previewLines.map(t => `<div class="label-preview-text">${escapeHtml(t)}</div>`).join("")}
              </div>
            </td>
            <td class="size-name">
              <div class="label-name">${escapeHtml(primaryText)}</div>
              <div class="label-size">${label.width}" × ${label.height}"</div>
              ${label.stickyBack ? '<div class="sticky-badge">STICKY BACK</div>' : ''}
            </td>
            <td class="font-cell">${label.font}</td>
            <td class="qty-cell" style="text-align: right;">×${label.quantity}</td>
          </tr>
        `;
      }).join("")}
    </tbody>
  </table>

  ${contactName || contactEmail ? `
    <div class="contact-info">
      <h3>Contact Information</h3>
      ${contactName ? `<p><strong>Name:</strong> ${escapeHtml(contactName)}</p>` : ''}
      ${contactEmail ? `<p><strong>Email:</strong> ${escapeHtml(contactEmail)}</p>` : ''}
    </div>
  ` : ''}

  ${notes ? `
    <div class="notes-section">
      <h3>Notes</h3>
      <p>${escapeHtml(notes)}</p>
    </div>
  ` : ''}
</body>
</html>`;

    // Upload to S3
    const xlsxKey = `nameplates/${refId}-${timestamp}.xlsx`;
    const htmlKey = `nameplates/${refId}-${timestamp}.html`;

    try {
      await s3.send(new PutObjectCommand({
        Bucket: bucket,
        Key: xlsxKey,
        Body: new Uint8Array(xlsxBuffer),
        ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }));

      await s3.send(new PutObjectCommand({
        Bucket: bucket,
        Key: htmlKey,
        Body: html,
        ContentType: "text/html",
      }));
    } catch (s3Error) {
      console.error("S3 upload error:", s3Error);
      return new Response(
        JSON.stringify({ error: "Failed to upload files", details: String(s3Error) }),
        { status: 500, headers: { "Content-Type": "application/json" } }
      );
    }

    const xlsxUrl = `https://${bucket}.s3.amazonaws.com/${xlsxKey}`;
    const htmlUrl = `https://${bucket}.s3.amazonaws.com/${htmlKey}`;

    // Send to Zapier
    if (zapierWebhook) {
      try {
        await fetch(zapierWebhook, {
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
            htmlUrl,
            labels: labels.map(l => ({
              height: l.height,
              width: l.width,
              labelColor: getColorName(l.labelColor),
              textColor: getColorName(l.textColor),
              font: l.font,
              corners: l.corners,
              stickyBack: l.stickyBack ? "Yes" : "No",
              quantity: l.quantity,
              textLines: l.textLines.filter(t => t.text).map(t => ({
                text: t.text,
                fontSize: t.fontSize,
              })),
            })),
          }),
        });
      } catch (zapierError) {
        console.error("Zapier webhook error:", zapierError);
        // Don't fail the request if Zapier fails
      }
    }

    return new Response(
      JSON.stringify({
        success: true,
        refId,
        xlsxUrl,
        htmlUrl,
        totalLabels,
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

export const config = {
  path: "/api/submit-order"
};
