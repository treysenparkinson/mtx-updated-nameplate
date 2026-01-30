import { S3Client, PutObjectCommand } from "@aws-sdk/client-s3";
import * as XLSX from "xlsx";

const getColorName = (hex: string): string => {
  const colors: Record<string, string> = {
    "#22c55e": "Green", "#ef4444": "Red", "#eab308": "Yellow",
    "#3b82f6": "Blue", "#1a1a1a": "Black", "#ffffff": "White",
    "#f97316": "Orange", "#6b7280": "Gray",
  };
  return colors[hex.toLowerCase()] || hex;
};

const escapeHtml = (text: string): string => {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
};

export default async (req: Request) => {
  if (req.method !== "POST") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), {
      status: 405,
      headers: { "Content-Type": "application/json" },
    });
  }

  try {
    console.log("Step 1: Parsing request body");
    const body = await req.json();
    const { refId, contactName, contactEmail, notes, labels } = body;

    if (!refId || !labels || labels.length === 0) {
      return new Response(JSON.stringify({ error: "Missing required fields" }), {
        status: 400,
        headers: { "Content-Type": "application/json" },
      });
    }

    console.log("Step 2: Checking env vars");
    const awsKey = process.env.MY_AWS_ACCESS_KEY_ID;
    const awsSecret = process.env.MY_AWS_SECRET_ACCESS_KEY;
    const awsRegion = process.env.MY_AWS_REGION || "us-east-1";
    const bucket = process.env.S3_BUCKET || "matrix-systems-labels";
    const zapierWebhook = process.env.ZAPIER_WEBHOOK_URL;

    console.log("ENV:", { hasKey: !!awsKey, hasSecret: !!awsSecret, region: awsRegion, bucket, hasZapier: !!zapierWebhook });

    if (!awsKey || !awsSecret) {
      return new Response(JSON.stringify({ error: "Missing AWS credentials" }), {
        status: 500,
        headers: { "Content-Type": "application/json" },
      });
    }

    console.log("Step 3: Creating S3 client");
    const s3 = new S3Client({
      region: awsRegion,
      credentials: { accessKeyId: awsKey, secretAccessKey: awsSecret },
    });

    const timestamp = Date.now();
    const formattedDate = new Date().toLocaleString("en-US", {
      timeZone: "America/Chicago",
      year: "numeric", month: "short", day: "numeric",
      hour: "2-digit", minute: "2-digit",
    });

    console.log("Step 4: Building Excel");
    // Find max text lines
    let maxTextLines = 1;
    for (const label of labels) {
      const count = label.textLines ? label.textLines.filter((t: any) => t.text).length : 0;
      if (count > maxTextLines) maxTextLines = count;
    }

    // Build headers
    const headers: string[] = ["ID #"];
    for (let i = 1; i <= maxTextLines; i++) {
      headers.push(`LINE ${i} TEXT`, `LINE ${i} TEXT SIZE`);
    }
    headers.push("BACKGRND COLOR", "LETTER COLOR", "WIDTH (INCHES)", "HEIGHT (INCHES)", "CORNERS", "STICKY BACK", "COMMENTS");

    // Build rows
    const rows: any[][] = [headers];
    for (const label of labels) {
      const qty = label.quantity || 1;
      for (let q = 0; q < qty; q++) {
        const row: any[] = [refId];
        for (let i = 0; i < maxTextLines; i++) {
          const line = label.textLines?.[i];
          row.push(line?.text || "", line?.fontSize || "");
        }
        row.push(
          getColorName(label.labelColor || "#22c55e"),
          getColorName(label.textColor || "#ffffff"),
          label.width || 7,
          label.height || 2,
          label.corners || "squared",
          label.stickyBack ? "Yes" : "No",
          notes || ""
        );
        rows.push(row);
      }
    }

    console.log("Step 5: Creating workbook");
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Labels");
    const xlsxBuffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

    console.log("Step 6: Building HTML");
    const totalLabels = labels.reduce((sum: number, l: any) => sum + (l.quantity || 1), 0);

    const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Nameplate Labels - ${escapeHtml(refId)}</title>
  <style>
    body { font-family: Arial, sans-serif; padding: 40px; color: #333; }
    .header { display: flex; justify-content: space-between; margin-bottom: 20px; }
    .ref-id { font-size: 14px; font-weight: bold; }
    .title { font-size: 28px; font-weight: bold; margin: 5px 0 20px 0; }
    .table { width: 100%; border-collapse: collapse; }
    .table th { text-align: left; padding: 12px; border-bottom: 2px solid #333; color: #666; }
    .table td { padding: 15px; border-bottom: 1px solid #eee; }
    .label-preview { width: 100px; height: 50px; display: flex; flex-direction: column; align-items: center; justify-content: center; font-size: 9px; border-radius: 4px; }
    .sticky-badge { background: #dbeafe; color: #1d4ed8; font-size: 9px; padding: 2px 6px; border-radius: 4px; margin-top: 4px; }
  </style>
</head>
<body>
  <div class="header">
    <div>
      <div class="ref-id">Reference ID: ${escapeHtml(refId)}</div>
      <div class="title">Saved Labels Summary</div>
    </div>
    <div style="text-align:right;color:#666;">${formattedDate}</div>
  </div>
  <table class="table">
    <tr><th>Preview</th><th>Size/Name</th><th>Font</th><th style="text-align:right;">Qty</th></tr>
    ${labels.map((label: any) => {
      const texts = (label.textLines || []).filter((t: any) => t.text).map((t: any) => t.text);
      const primary = texts[0] || "—";
      return `<tr>
        <td><div class="label-preview" style="background:${label.labelColor || '#22c55e'};color:${label.textColor || '#fff'};">${texts.slice(0, 2).map((t: string) => `<div>${escapeHtml(t)}</div>`).join("")}</div></td>
        <td><strong>${escapeHtml(primary)}</strong><br/>${label.width}" × ${label.height}"${label.stickyBack ? '<div class="sticky-badge">STICKY BACK</div>' : ''}</td>
        <td>${label.font || 'Calibri'}</td>
        <td style="text-align:right;font-weight:bold;">×${label.quantity || 1}</td>
      </tr>`;
    }).join("")}
  </table>
  ${notes ? `<div style="margin-top:20px;padding:15px;background:#fffbeb;border-radius:8px;"><strong>Notes:</strong> ${escapeHtml(notes)}</div>` : ''}
</body>
</html>`;

    console.log("Step 7: Uploading to S3");
    const xlsxKey = `nameplates/${refId}-${timestamp}.xlsx`;
    const htmlKey = `nameplates/${refId}-${timestamp}.html`;

    await s3.send(new PutObjectCommand({
      Bucket: bucket,
      Key: xlsxKey,
      Body: new Uint8Array(xlsxBuffer),
      ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }));
    console.log("Excel uploaded");

    await s3.send(new PutObjectCommand({
      Bucket: bucket,
      Key: htmlKey,
      Body: html,
      ContentType: "text/html",
    }));
    console.log("HTML uploaded");

    const xlsxUrl = `https://${bucket}.s3.amazonaws.com/${xlsxKey}`;
    const htmlUrl = `https://${bucket}.s3.amazonaws.com/${htmlKey}`;

    console.log("Step 8: Sending to Zapier");
    if (zapierWebhook) {
      await fetch(zapierWebhook, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          refId, contactName, contactEmail, notes, formattedDate,
          totalLabels, xlsxUrl, htmlUrl,
        }),
      });
      console.log("Zapier sent");
    }

    console.log("Step 9: Done!");
    return new Response(JSON.stringify({ success: true, refId, xlsxUrl, htmlUrl }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });

  } catch (error) {
    console.error("ERROR:", error);
    return new Response(JSON.stringify({ error: "Failed", details: String(error) }), {
      status: 500,
      headers: { "Content-Type": "application/json" },
    });
  }
};

export const config = { path: "/api/submit-order" };
