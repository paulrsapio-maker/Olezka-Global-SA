/**
 * Shared code between client and server
 * Useful to share types between client and server
 * and/or small pure JS functions that can be used on both client and server
 */

/**
 * Example response type for /api/demo
 */
export interface DemoResponse {
  message: string;
}

/**
 * Olezka Global Cloud Security Posture Assessment types
 */
export type NistFunction =
  | "GOVERN"
  | "IDENTIFY"
  | "PROTECT"
  | "DETECT"
  | "RESPOND"
  | "RECOVER";

export interface AssessmentQuestionMeta {
  id: string; // e.g., GO.SC-1
  nistFunction: NistFunction; // e.g., GOVERN
  category: string; // e.g., GO.SC: Strategic Context
  control: string; // e.g., GO.SC-1: The organization's role in the supply chain is understood.
  prompt: string; // Question shown to client
}

export interface AssessmentResponseItem {
  meta: AssessmentQuestionMeta;
  response: string; // Client's Response
  maturity: 0 | 1 | 2 | 3 | 4; // Maturity Rating (0-4)
}

export interface AssessmentSubmission {
  organization: string;
  contactEmail: string;
  submittedAt: string; // ISO string
  responses: AssessmentResponseItem[];
}
