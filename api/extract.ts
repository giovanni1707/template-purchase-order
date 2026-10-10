import { handleExtractRequest } from "../src/server/extractService.ts";

export const config = {
  api: {
    bodyParser: {
      sizeLimit: "50mb",
    },
  },
};

export default async function handler(req: any, res: any) {
  if (req.method !== "POST") {
    res.setHeader("Allow", ["POST"]);
    return res.status(405).json({ error: "Method not allowed. Use POST." });
  }

  try {
    const result = await handleExtractRequest(req.body);
    return res.status(result.status).json(result.data);
  } catch (err: any) {
    console.error("Vercel api/extract error:", err);
    return res.status(500).json({ error: "Serverless extraction error: " + (err?.message || "Unknown") });
  }
}
