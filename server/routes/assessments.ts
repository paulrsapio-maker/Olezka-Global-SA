import { RequestHandler } from "express";
import { z } from "zod";
import { AssessmentSubmission } from "@shared/api";
import { getPool } from "../db";

const ResponseItemSchema = z.object({
  meta: z.object({
    id: z.string(),
    nistFunction: z.enum([
      "GOVERN",
      "IDENTIFY",
      "PROTECT",
      "DETECT",
      "RESPOND",
      "RECOVER",
    ]),
    category: z.string(),
    control: z.string(),
    prompt: z.string(),
  }),
  response: z.string().min(1),
  maturity: z.union([
    z.literal(0),
    z.literal(1),
    z.literal(2),
    z.literal(3),
    z.literal(4),
  ]),
});

const SubmissionSchema = z.object({
  organization: z.string().min(2),
  contactEmail: z.string().email(),
  submittedAt: z.string(),
  responses: z.array(ResponseItemSchema).min(1),
});

export const submitAssessment: RequestHandler = async (req, res) => {
  const parsed = SubmissionSchema.safeParse(req.body);
  if (!parsed.success) {
    return res
      .status(400)
      .json({
        error: "Invalid submission payload",
        details: parsed.error.flatten(),
      });
  }

  const data = parsed.data as AssessmentSubmission;

  const pool = getPool();
  if (!pool) {
    return res.status(503).json({
      error: "Database not configured",
      message:
        "Set DATABASE_URL (e.g., Neon connection string) to enable persistence.",
    });
  }

  try {
    const client = await pool.connect();
    try {
      await client.query("BEGIN");
      const insertAssessment = `
        INSERT INTO assessments (organization, contact_email, submitted_at)
        VALUES ($1, $2, $3)
        RETURNING id
      `;
      const { rows } = await client.query(insertAssessment, [
        data.organization,
        data.contactEmail,
        data.submittedAt,
      ]);
      const assessmentId = rows[0].id as number;

      const insertResponse = `
        INSERT INTO assessment_responses (
          assessment_id, meta_id, nist_function, category, control, prompt, response, maturity
        ) VALUES ($1,$2,$3,$4,$5,$6,$7,$8)
      `;
      for (const item of data.responses) {
        await client.query(insertResponse, [
          assessmentId,
          item.meta.id,
          item.meta.nistFunction,
          item.meta.category,
          item.meta.control,
          item.meta.prompt,
          item.response,
          item.maturity,
        ]);
      }

      await client.query("COMMIT");
      return res.status(201).json({ ok: true, id: assessmentId });
    } catch (e) {
      await client.query("ROLLBACK");
      throw e;
    } finally {
      client.release();
    }
  } catch (e) {
    console.error("Failed to persist assessment:", e);
    return res.status(500).json({ error: "Failed to persist assessment" });
  }
};
