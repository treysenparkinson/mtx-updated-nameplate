import { S3Client, PutObjectCommand } from "@aws-sdk/client-s3";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";

const getColorName = (hex: string): string => {
  const colors: Record<string, string> = {
    "#22c55e": "Green", "#ef4444": "Red", "#eab308": "Yellow",
    "#3b82f6": "Blue", "#1a1a1a": "Black", "#ffffff": "White",
    "#f97316": "Orange", "#6b7280": "Gray",
  };
  return colors[hex.toLowerCase()] || hex;
};

const hexToRgb = (hex: string): { r: number; g: number; b: number } => {
  const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result ? {
    r: parseInt(result[1], 16),
    g: parseInt(result[2], 16),
    b: parseInt(result[3], 16)
  } : { r: 0, g: 0, b: 0 };
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

    if (!awsKey || !awsSecret) {
      console.log("Missing credentials:", { hasKey: !!awsKey, hasSecret: !!awsSecret });
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

    console.log("Step 5: Creating Excel workbook");
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Labels");
    const xlsxBuffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

    console.log("Step 6: Building PDF");
    const totalLabels = labels.reduce((sum: number, l: any) => sum + (l.quantity || 1), 0);
    
    // Create PDF
    const doc = new jsPDF();
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 20;
    let y = margin;

    // Header
    doc.setFontSize(11);
    doc.setFont("helvetica", "bold");
    doc.text(`Reference ID: ${refId}`, margin, y);
    doc.setFont("helvetica", "normal");
    doc.text(formattedDate, pageWidth - margin, y, { align: "right" });
    y += 10;

    doc.setFontSize(22);
    doc.setFont("helvetica", "bold");
    doc.text("Saved Labels Summary", margin, y);
    doc.setFontSize(10);
    doc.setFont("helvetica", "normal");
    doc.text(`Page 1 of 1`, pageWidth - margin, y, { align: "right" });
    y += 15;

    // Table header
    doc.setDrawColor(0);
    doc.setLineWidth(0.5);
    doc.line(margin, y, pageWidth - margin, y);
    y += 7;

    doc.setFontSize(10);
    doc.setTextColor(100);
    doc.text("Preview", margin, y);
    doc.text("Size/Name", margin + 45, y);
    doc.text("Font", margin + 115, y);
    doc.text("Qty", pageWidth - margin - 10, y, { align: "right" });
    y += 5;
    doc.line(margin, y, pageWidth - margin, y);
    y += 10;

    // Table rows
    doc.setTextColor(0);
    for (const label of labels) {
      const texts = (label.textLines || []).filter((t: any) => t.text).map((t: any) => t.text);
      const primary = texts[0] || "—";
      
      // Check if we need a new page
      if (y > pageHeight - 40) {
        doc.addPage();
        y = margin;
      }

      // Draw label preview box
      const boxWidth = 35;
      const boxHeight = 20;
      const labelColor = hexToRgb(label.labelColor || "#22c55e");
      const textColor = hexToRgb(label.textColor || "#ffffff");
      
      doc.setFillColor(labelColor.r, labelColor.g, labelColor.b);
      if (label.corners === "rounded") {
        doc.roundedRect(margin, y - 5, boxWidth, boxHeight, 3, 3, "F");
      } else {
        doc.rect(margin, y - 5, boxWidth, boxHeight, "F");
      }
      
      // Text inside preview box
      doc.setFontSize(6);
      doc.setTextColor(textColor.r, textColor.g, textColor.b);
      const previewTexts = texts.slice(0, 2);
      previewTexts.forEach((t: string, i: number) => {
        const truncated = t.length > 10 ? t.substring(0, 10) + "..." : t;
        doc.text(truncated, margin + boxWidth / 2, y + 2 + (i * 5), { align: "center" });
      });

      // Size/Name column
      doc.setTextColor(0);
      doc.setFontSize(11);
      doc.setFont("helvetica", "bold");
      const truncatedPrimary = primary.length > 25 ? primary.substring(0, 25) + "..." : primary;
      doc.text(truncatedPrimary, margin + 45, y);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      doc.setTextColor(100);
      doc.text(`${label.width}" × ${label.height}"`, margin + 45, y + 5);
      
      // Sticky back badge
      if (label.stickyBack) {
        doc.setFillColor(219, 234, 254);
        doc.roundedRect(margin + 45, y + 7, 25, 5, 1, 1, "F");
        doc.setFontSize(6);
        doc.setTextColor(29, 78, 216);
        doc.text("STICKY BACK", margin + 47, y + 10.5);
      }

      // Font column
      doc.setTextColor(0);
      doc.setFontSize(10);
      doc.text(label.font || "Calibri", margin + 115, y);

      // Qty column
      doc.setFont("helvetica", "bold");
      doc.text(`×${label.quantity || 1}`, pageWidth - margin - 10, y, { align: "right" });
      doc.setFont("helvetica", "normal");

      y += boxHeight + 10;

      // Divider line
      doc.setDrawColor(230);
      doc.setLineWidth(0.2);
      doc.line(margin, y - 5, pageWidth - margin, y - 5);
    }

    // Bottom line
    doc.setDrawColor(0);
    doc.setLineWidth(0.5);
    doc.line(margin, y - 5, pageWidth - margin, y - 5);

    // Contact info
    if (contactName || contactEmail) {
      y += 10;
      doc.setFillColor(248, 248, 248);
      doc.roundedRect(margin, y - 5, pageWidth - margin * 2, 25, 3, 3, "F");
      doc.setFontSize(10);
      doc.setFont("helvetica", "bold");
      doc.setTextColor(0);
      doc.text("Contact Information", margin + 5, y + 3);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      doc.setTextColor(100);
      if (contactName) doc.text(`Name: ${contactName}`, margin + 5, y + 10);
      if (contactEmail) doc.text(`Email: ${contactEmail}`, margin + 5, y + 16);
      y += 30;
    }

    // Notes
    if (notes) {
      y += 5;
      doc.setFillColor(255, 251, 235);
      const notesHeight = Math.max(20, Math.ceil(notes.length / 80) * 6 + 15);
      doc.roundedRect(margin, y - 5, pageWidth - margin * 2, notesHeight, 3, 3, "F");
      doc.setFontSize(10);
      doc.setFont("helvetica", "bold");
      doc.setTextColor(146, 64, 14);
      doc.text("Notes", margin + 5, y + 3);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      doc.setTextColor(100);
      const splitNotes = doc.splitTextToSize(notes, pageWidth - margin * 2 - 10);
      doc.text(splitNotes, margin + 5, y + 10);
    }

    // Get PDF as buffer
    const pdfBuffer = Buffer.from(doc.output("arraybuffer"));

    console.log("Step 7: Uploading to S3");
    const xlsxKey = `nameplates/${refId}-${timestamp}.xlsx`;
    const pdfKey = `nameplates/${refId}-${timestamp}.pdf`;

    await s3.send(new PutObjectCommand({
      Bucket: bucket,
      Key: xlsxKey,
      Body: new Uint8Array(xlsxBuffer),
      ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }));
    console.log("Excel uploaded");

    await s3.send(new PutObjectCommand({
      Bucket: bucket,
      Key: pdfKey,
      Body: pdfBuffer,
      ContentType: "application/pdf",
    }));
    console.log("PDF uploaded");

    const xlsxUrl = `https://${bucket}.s3.amazonaws.com/${xlsxKey}`;
    const pdfUrl = `https://${bucket}.s3.amazonaws.com/${pdfKey}`;

    console.log("Step 8: Sending to Zapier");
    if (zapierWebhook) {
      await fetch(zapierWebhook, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          refId, contactName, contactEmail, notes, formattedDate,
          totalLabels, xlsxUrl, pdfUrl,
        }),
      });
      console.log("Zapier sent");
    }

    console.log("Step 9: Done!");
    return new Response(JSON.stringify({ success: true, refId, xlsxUrl, pdfUrl }), {
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
