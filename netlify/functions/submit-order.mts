export default async (req: Request) => {
  if (req.method !== "POST") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), {
      status: 405,
      headers: { "Content-Type": "application/json" },
    });
  }

  try {
    const body = await req.json();
    const { refId, labels } = body;

    // Log environment variables (without exposing secrets)
    console.log("ENV CHECK:", {
      hasAwsKey: !!process.env.MY_AWS_ACCESS_KEY_ID,
      hasAwsSecret: !!process.env.MY_AWS_SECRET_ACCESS_KEY,
      region: process.env.MY_AWS_REGION,
      bucket: process.env.S3_BUCKET,
      hasZapier: !!process.env.ZAPIER_WEBHOOK_URL,
    });

    console.log("REQUEST:", { refId, labelCount: labels?.length });

    // Just test Zapier webhook first
    const zapierWebhook = process.env.ZAPIER_WEBHOOK_URL;
    
    if (zapierWebhook) {
      console.log("Sending to Zapier...");
      const zapierResponse = await fetch(zapierWebhook, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          refId,
          test: true,
          timestamp: Date.now(),
          labelCount: labels?.length || 0,
        }),
      });
      console.log("Zapier response status:", zapierResponse.status);
    } else {
      console.log("No Zapier webhook configured");
    }

    return new Response(
      JSON.stringify({ 
        success: true, 
        refId,
        message: "Test successful",
        envCheck: {
          hasAwsKey: !!process.env.MY_AWS_ACCESS_KEY_ID,
          hasAwsSecret: !!process.env.MY_AWS_SECRET_ACCESS_KEY,
          region: process.env.MY_AWS_REGION,
          bucket: process.env.S3_BUCKET,
          hasZapier: !!process.env.ZAPIER_WEBHOOK_URL,
        }
      }),
      { status: 200, headers: { "Content-Type": "application/json" } }
    );

  } catch (error) {
    console.error("Error:", error);
    return new Response(
      JSON.stringify({ error: "Failed", details: String(error) }),
      { status: 500, headers: { "Content-Type": "application/json" } }
    );
  }
};

export const config = {
  path: "/api/submit-order"
};
