import { RequestHandler } from "express";
import { z } from "zod";
import { AssessmentSubmission } from "@shared/api";

const AIReportSchema = z.object({
  environment: z.string().optional(),
  summary: z.string(),
  strengths: z.array(z.string()).default([]),
  gaps: z
    .array(
      z.object({
        name: z.string(),
        description: z.string().optional().default(""),
        businessImpact: z.string().optional().default(""),
        likelihood: z.enum(["Low", "Medium", "High"]).or(z.string()).optional().default("Medium"),
        impact: z.enum(["Low", "Medium", "High"]).or(z.string()).optional().default("Medium"),
        rating: z.enum(["Low", "Medium", "High"]).or(z.string()).optional().default("Medium"),
      })
    )
    .default([]),
  compliance: z.string().default(""),
  recommendations: z
    .array(
      z.union([
        z.string(),
        z.object({ action: z.string(), priority: z.string().optional().default("Medium") }),
      ])
    )
    .default([]),
  conclusion: z.string().default(""),
});

export const generateAssessmentReport: RequestHandler = async (req, res) => {
  const body = req.body as AssessmentSubmission | undefined;
  if (!body || !body.organization || !Array.isArray(body.responses)) {
    return res.status(400).json({ error: "Invalid submission payload" });
  }

  const apiKey = process.env.OPENAI_API_KEY || process.env.OPENAI_API_TOKEN || process.env.GPT_API_KEY;
  if (!apiKey) {
    return res.status(503).json({
      error: "AI not configured",
      message: "Set OPENAI_API_KEY environment variable to enable AI analysis.",
    });
  }

  // Build a compact dataset for the prompt
  const dataset = body.responses.map((r) => ({
    id: r.meta.id,
    nistFunction: r.meta.nistFunction,
    category: r.meta.category,
    control: r.meta.control,
    prompt: r.meta.prompt,
    response: r.response,
    maturity: r.maturity,
  }));

  // Compute quick stats to guide the model
  const byFunction: Record<string, { total: number; count: number }> = {};
  for (const it of dataset) {
    const key = it.nistFunction;
    if (!byFunction[key]) byFunction[key] = { total: 0, count: 0 };
    byFunction[key].total += it.maturity;
    byFunction[key].count += 1;
  }
  const functionAverages = Object.fromEntries(
    Object.entries(byFunction).map(([k, v]) => [k, v.count ? v.total / v.count : 0])
  );

  const systemPrompt =
    "You are a senior cloud security consultant creating an executive summary based on a Cloud Security Posture Assessment (CSPA). Use clear, concise language for executives. Map findings to business impact and NIST CSF. Output only valid JSON matching the response schema. Avoid markdown.";

  const userPrompt = `Based on the following questionnaire responses from a Cloud Security Posture Assessment, generate a professional Executive Summary Report. The report should include:

1. A high-level overview of the organization's cloud environment (e.g., Azure, AWS, GCP).
2. Key strengths in cloud security posture (e.g., identity management, data protection, threat detection).
3. Identified gaps or risks (e.g., misconfigurations, lack of logging, missing controls).
4. Compliance alignment with NIST CSF.
5. Strategic recommendations for remediation and improvement.
6. Add a risk matrix summary (likelihood, impact, overall rating for each primary risk).
7. A summary conclusion with business impact and next steps.

Organization: ${body.organization}
Contact: ${body.contactEmail}
Submitted At: ${body.submittedAt}
NIST Function Averages (0-4): ${JSON.stringify(functionAverages)}

Questionnaire Responses (array of items):\n${JSON.stringify(dataset)}\n\nRespond ONLY with JSON using this TypeScript type (no markdown):\n${AIReportSchema.toString()}`;

  try {
    const resp = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model: "gpt-4o-mini",
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: userPrompt },
        ],
        temperature: 0.2,
      }),
    });

    if (!resp.ok) {
      const text = await resp.text();
      return res.status(502).json({ error: "AI request failed", details: text });
    }

    const data = (await resp.json()) as any;
    const content = data?.choices?.[0]?.message?.content?.trim?.();
    if (!content) {
      return res.status(502).json({ error: "Empty AI response" });
    }

    // Try to locate JSON in content
    let jsonText = content;
    const firstBrace = content.indexOf("{");
    const lastBrace = content.lastIndexOf("}");
    if (firstBrace !== -1 && lastBrace !== -1) {
      jsonText = content.slice(firstBrace, lastBrace + 1);
    }

    let parsed: unknown;
    try {
      parsed = JSON.parse(jsonText);
    } catch (e) {
      return res.status(502).json({ error: "Failed to parse AI JSON", details: String(e) });
    }

    const safe = AIReportSchema.safeParse(parsed);
    if (!safe.success) {
      return res.status(502).json({ error: "AI JSON did not match schema", details: safe.error.flatten() });
    }

    const report = safe.data;

    return res.json({
      ok: true,
      report,
      metrics: {
        functionAverages,
        totalQuestions: body.responses.length,
      },
    });
  } catch (e: any) {
    return res.status(500).json({ error: "AI processing error", details: String(e?.message || e) });
  }
};
