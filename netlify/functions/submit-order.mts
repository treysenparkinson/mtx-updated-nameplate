import { S3Client, PutObjectCommand } from "@aws-sdk/client-s3";
import * as XLSX from "xlsx";

export default async (req: Request) => {
  // Only allow POST
  if (req.method !== "POST") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), {
      status: 405,
      headers: { "Content-Type": "application/json" },
    });
  }

  try {
    const body = await req.json();
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
        JSON.stringify({ error: "Server configuration error - missing AWS credentials" }),
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
    });

    // Helper functions
    const getColorName = (hex: string): string => {
      const colors: Record<string, string> = {
        "#22c55e": "Green", "#ef4444": "Red", "#eab308": "Yellow",
        "#3b82f6": "Blue", "#1a1a1a": "Black", "#ffffff": "White",
        "#f97316": "Orange", "#6b7280": "Gray",
      };
      return colors[hex.toLowerCase()] || hex;
    };

    const escapeHtml = (text: string): string => {
      return text.replace(/&/g, "&amp;").replace(/</g, "&lt;")
        .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
    };

    // Generate Excel
    const rows: any[][] = [];
    rows.push([`Reference ID: ${refId}`]);
    rows.push([]);
    rows.push(["Height", "Width", "Label Color", "Text Color", "Font", "Corners", "Text 1", "Size 1", "Text 2", "Size 2", "Text 3", "Size 3"]);

    for (const label of labels) {
      for (let i = 0; i < label.quantity; i++) {
        const row: any[] = [
          label.height, label.width, getColorName(label.labelColor),
          getColorName(label.textColor), label.font, label.corners,
        ];
        for (let j = 0; j < 3; j++) {
          const line = label.textLines?.[j];
          row.push(line?.text || "");
          row.push(line?.fontSize || "");
        }
        rows.push(row);
      }
    }

    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Labels");
    const xlsxBuffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

    // Generate HTML
    const totalLabels = labels.reduce((sum: number, l: any) => sum + l.quantity, 0);
    const html = `<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>Order ${refId}</title>
<style>body{font-family:Arial,sans-serif;padding:40px}.header{border-bottom:2px solid #10b981;padding-bottom:20px;margin-bottom:30px}
.label-card{border:1px solid #ddd;padding:15px;margin:10px 0;border-radius:8px}</style></head>
<body><div class="header"><h1>Nameplate Label Order</h1><p><strong>Reference:</strong> ${escapeHtml(refId)}</p>
<p><strong>Date:</strong> ${formattedDate}</p>
${contactName ? `<p><strong>Contact:</strong> ${escapeHtml(contactName)}</p>` : ""}
${contactEmail ? `<p><strong>Email:</strong> ${escapeHtml(contactEmail)}</p>` : ""}
<p><strong>Total Labels:</strong> ${totalLabels}</p></div>
${labels.map((l: any) => `<div class="label-card">
<p><strong>Size:</strong> ${l.width}" × ${l.height}" | <strong>Colors:</strong> ${getColorName(l.labelColor)}/${getColorName(l.textColor)} | <strong>Qty:</strong> ${l.quantity}</p>
<p><strong>Text:</strong> ${l.textLines?.filter((t: any) => t.text).map((t: any) => escapeHtml(t.text)).join(", ") || "None"}</p>
</div>`).join("")}
${notes ? `<div style="margin-top:20px;padding:15px;background:#fffbeb;border-radius:8px"><strong>Notes:</strong> ${escapeHtml(notes)}</div>` : ""}
</body></html>`;

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
        JSON.stringify({ error: "Failed to upload files to S3", details: String(s3Error) }),
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
            refId, contactName, contactEmail, notes, timestamp, formattedDate,
            totalLabels, labelCount: labels.length, xlsxUrl, htmlUrl,
            labels: labels.map((l: any) => ({
              height: l.height, width: l.width,
              labelColor: getColorName(l.labelColor), textColor: getColorName(l.textColor),
              font: l.font, corners: l.corners, quantity: l.quantity,
              textLines: l.textLines?.filter((t: any) => t.text).map((t: any) => ({ text: t.text, fontSize: t.fontSize })),
            })),
          }),
        });
      } catch (zapierError) {
        console.error("Zapier webhook error:", zapierError);
        // Don't fail the request if Zapier fails
      }
    }

    return new Response(
      JSON.stringify({ success: true, refId, xlsxUrl, htmlUrl }),
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
