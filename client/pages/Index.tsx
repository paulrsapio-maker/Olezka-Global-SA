import { useMemo } from "react";
import { useForm, useFieldArray } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { z } from "zod";
import {
  AssessmentQuestionMeta,
  AssessmentSubmission,
  NistFunction,
} from "@shared/api";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Label } from "@/components/ui/label";
import { Button } from "@/components/ui/button";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import {
  Accordion,
  AccordionContent,
  AccordionItem,
  AccordionTrigger,
} from "@/components/ui/accordion";
import { toast } from "sonner";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { Download } from "lucide-react";
import { Link } from "react-router-dom";

type FormResponse = {
  meta: AssessmentQuestionMeta;
  response: string;
  maturity: 0 | 1 | 2 | 3 | 4;
};

const questions: AssessmentQuestionMeta[] = [
  {
    id: "GO.SC-1",
    nistFunction: "GOVERN",
    category: "GO.SC: Strategic Context",
    control:
      "GO.SC-1: The organization's role in the supply chain is understood.",
    prompt:
      "How is cybersecurity risk managed at the leadership level? Is there a designated individual or committee responsible for overseeing cloud security?",
  },
  {
    id: "GO.RM-1",
    nistFunction: "GOVERN",
    category: "GO.RM: Risk Management",
    control: "GO.RM-1: Risk management processes are defined and documented.",
    prompt:
      "Do you have an established risk tolerance for your cloud environment?",
  },
  {
    id: "ID.AM-1",
    nistFunction: "IDENTIFY",
    category: "ID.AM: Asset Management",
    control: "ID.AM-1: Physical devices and systems are inventoried.",
    prompt:
      "Please provide an inventory of all Azure subscriptions and tenants.",
  },
  {
    id: "ID.AM-2",
    nistFunction: "IDENTIFY",
    category: "ID.AM: Asset Management",
    control: "ID.AM-2: Software platforms and applications are inventoried.",
    prompt:
      "Please provide a list of all mission-critical applications running on Azure and all Office 365 services in use.",
  },
  {
    id: "ID.AM-3",
    nistFunction: "IDENTIFY",
    category: "ID.AM: Asset Management",
    control: "ID.AM-3: Communication and data flows are mapped.",
    prompt:
      "Who is the business owner for the data in your SharePoint and OneDrive environment? Who decides what data is sensitive?",
  },
  {
    id: "ID.BE-3",
    nistFunction: "IDENTIFY",
    category: "ID.BE: Business Environment",
    control: "ID.BE-3: Priorities for organizational mission are established.",
    prompt:
      "If your core Azure-based student database were unavailable for 24 hours, what would the business impact be?",
  },
  {
    id: "PR.AC-4",
    nistFunction: "PROTECT",
    category: "PR.AC: Identity Management & Access Control",
    control: "PR.AC-4: MFA is used for all users.",
    prompt:
      "Is Multi-Factor Authentication (MFA) enabled for all user accounts, including administrators? Please specify any exceptions.",
  },
  {
    id: "PR.DS-1",
    nistFunction: "PROTECT",
    category: "PR.DS: Data Security",
    control: "PR.DS-1: Data is protected at rest and in transit.",
    prompt:
      "Are Azure Storage accounts, databases, and other data sources configured with encryption at rest? Is TLS/SSL enforced for all data in transit?",
  },
  {
    id: "PR.DS-4",
    nistFunction: "PROTECT",
    category: "PR.DS: Data Security",
    control: "PR.DS-4: Data Loss Prevention (DLP) is implemented.",
    prompt:
      "Do you have DLP policies configured in Office 365 for Exchange, SharePoint, and OneDrive?",
  },
  {
    id: "PR.PT-1",
    nistFunction: "PROTECT",
    category: "PR.PT: Protective Technology",
    control: "PR.PT-1: Auditable events are logged.",
    prompt:
      "Is auditing enabled for all critical security events in Azure and Office 365? Are these logs centralized?",
  },
  {
    id: "DE.CM-1",
    nistFunction: "DETECT",
    category: "DE.CM: Security Continuous Monitoring",
    control: "DE.CM-1: The network is monitored for malicious activity.",
    prompt:
      "What tools do you use for continuous monitoring of your Azure tenant?",
  },
  {
    id: "RS.RP-1",
    nistFunction: "RESPOND",
    category: "RS.RP: Response Planning",
    control: "RS.RP-1: A response plan is in place.",
    prompt:
      "Do you have a documented incident response plan specifically for a cloud security event involving your Azure or O365 tenant?",
  },
  {
    id: "RS.MI-1",
    nistFunction: "RESPOND",
    category: "RS.MI: Mitigation",
    control: "RS.MI-1: Mitigation activities are executed.",
    prompt:
      "What are your documented procedures for containing an incident, such as a compromised admin account or a data breach in SharePoint?",
  },
  {
    id: "RC.RP-1",
    nistFunction: "RECOVER",
    category: "RC.RP: Recovery Planning",
    control: "RC.RP-1: Recovery plan is executed.",
    prompt:
      "Is there a documented disaster recovery plan for your Azure workloads? How often is it tested?",
  },
  {
    id: "RC.IM-1",
    nistFunction: "RECOVER",
    category: "RC.IM: Improvements",
    control: "RC.IM-1: Lessons learned are incorporated.",
    prompt:
      "Is there a formal process for reviewing past security incidents to improve your security posture?",
  },
];

const schema = z.object({
  organization: z.string().min(2),
  contactEmail: z.string().email(),
  responses: z
    .array(
      z.object({
        response: z.string().min(1, "Please provide a response"),
        maturity: z.coerce.number().min(0).max(4),
      }),
    )
    .length(questions.length),
});

type FormValues = z.infer<typeof schema>;

function groupByFunction(items: AssessmentQuestionMeta[]) {
  return items.reduce(
    (acc, q) => {
      (acc[q.nistFunction] ||= []).push(q);
      return acc;
    },
    {} as Record<NistFunction, AssessmentQuestionMeta[]>,
  );
}

export default function Index() {
  const grouped = useMemo(() => groupByFunction(questions), []);
  const {
    control,
    register,
    handleSubmit,
    setValue,
    formState: { isSubmitting },
  } = useForm<FormValues>({
    resolver: zodResolver(schema),
    defaultValues: {
      organization: "",
      contactEmail: "",
      responses: questions.map(() => ({ response: "", maturity: 0 })),
    },
    mode: "onBlur",
  });

  const { fields } = useFieldArray({ name: "responses", control });

  function addRealData() {
    setValue("organization", "Contoso Education", { shouldDirty: true, shouldValidate: true });
    setValue("contactEmail", "secops@contoso.edu", { shouldDirty: true, shouldValidate: true });

    const sampleById: Record<string, string> = {
      "GO.SC-1": "Cybersecurity risk is governed by the IT Governance Committee chaired by the CIO. Cloud risk is on the board agenda twice per year with KPIs covering MFA adoption, privileged access reviews, and incident MTTR.",
      "GO.RM-1": "Risk tolerance is documented in the Enterprise Risk Register. For cloud services we target RTO ≤ 8 hours and RPO ≤ 4 hours for Tier-1 workloads; high-risk changes require CAB approval.",
      "ID.AM-1": "We operate two Azure AD tenants and three subscriptions (Prod, NonProd, Sandbox). Resources are tagged with Owner, DataClass, and Criticality; inventory is exported weekly to Log Analytics.",
      "ID.AM-2": "Mission-critical apps include SIS on Azure SQL, LMS on AKS, and O365 (Exchange, SharePoint, OneDrive, Teams). Business apps are tracked in a CMDB with dependency maps.",
      "ID.AM-3": "Data owners are assigned for each SharePoint site and OneDrive. Sensitive data is defined in the Data Classification Policy (Public, Internal, Confidential, Restricted).",
      "ID.BE-3": "If the Azure-based student database is unavailable for 24 hours, enrollment, grading, and attendance are disrupted; estimated impact ~$250k and reputational risk during exam periods.",
      "PR.AC-4": "MFA is enforced via Conditional Access for all users with number matching; break-glass accounts are stored in a safe and monitored. Admin roles are time-bound via PIM.",
      "PR.DS-1": "Encryption at rest uses Microsoft-managed keys; TLS 1.2+ enforced; private endpoints used for storage and databases; Defender for Cloud flags are remediated within 7 days.",
      "PR.DS-4": "Purview DLP policies protect SSN/PCI/PHI across Exchange, SharePoint, and OneDrive; policy tips are enabled with auto-quarantine for high risk.",
      "PR.PT-1": "Unified audit logging is enabled; Azure diagnostic settings forward to Log Analytics and Sentinel. Critical events (role changes, mailbox rules) are alerted in real-time.",
      "DE.CM-1": "Defender for Cloud and Microsoft Sentinel provide continuous monitoring; analytic rules cover impossible travel, OAuth consent grants, and key vault access anomalies.",
      "RS.RP-1": "An IR plan aligned to NIST is reviewed annually; communication templates and severity levels are defined for cloud incidents.",
      "RS.MI-1": "Containment runbooks exist for compromised admin accounts and SharePoint data exposure; playbooks auto-disable tokens, reset creds, and revoke sessions.",
      "RC.RP-1": "DR plans leverage Azure Backup and ASR; failover tests are performed semi-annually for Tier-1 services with documented results.",
      "RC.IM-1": "After-action reviews occur within 10 business days; lessons learned feed backlog items and policy updates tracked to completion.",
    };

    const maturityById: Record<string, 0 | 1 | 2 | 3 | 4> = {
      "GO.SC-1": 2,
      "GO.RM-1": 1,
      "ID.AM-1": 3,
      "ID.AM-2": 2,
      "ID.AM-3": 2,
      "ID.BE-3": 2,
      "PR.AC-4": 3,
      "PR.DS-1": 3,
      "PR.DS-4": 2,
      "PR.PT-1": 2,
      "DE.CM-1": 1,
      "RS.RP-1": 2,
      "RS.MI-1": 2,
      "RC.RP-1": 3,
      "RC.IM-1": 2,
    };

    questions.forEach((q, i) => {
      const txt = sampleById[q.id] || `Response for ${q.id}`;
      const m = maturityById[q.id] ?? 2;
      setValue(`responses.${i}.response` as const, txt, { shouldDirty: true, shouldValidate: true });
      setValue(`responses.${i}.maturity` as const, m, { shouldDirty: true, shouldValidate: true });
      const hidden = document.getElementById(`responses.${i}.maturity-hidden`) as HTMLInputElement | null;
      if (hidden) {
        hidden.value = String(m);
        hidden.dispatchEvent(new Event("input", { bubbles: true }));
        hidden.dispatchEvent(new Event("change", { bubbles: true }));
      }
    });
  }

  function hslToRgb(h: number, s: number, l: number): [number, number, number] {
    s /= 100;
    l /= 100;
    const k = (n: number) => (n + h / 30) % 12;
    const a = s * Math.min(l, 1 - l);
    const f = (n: number) =>
      l - a * Math.max(-1, Math.min(k(n) - 3, Math.min(9 - k(n), 1)));
    return [
      Math.round(255 * f(0)),
      Math.round(255 * f(8)),
      Math.round(255 * f(4)),
    ];
  }

  function getBrandRgb(): [number, number, number] {
    const root = getComputedStyle(document.documentElement);
    const hsl = root.getPropertyValue("--primary").trim();
    // value like: "167 92% 53%"
    const [h, s, l] = hsl
      .split(" ")
      .map((v, i) =>
        i === 0 ? parseFloat(v) : parseFloat(v.replace("%", "")),
      );
    return hslToRgb(h, s, l);
  }

  function drawPdfBackground(doc: jsPDF) {
    const width = doc.internal.pageSize.getWidth();
    const height = doc.internal.pageSize.getHeight();
    doc.setFillColor(0, 0, 0);
    doc.rect(0, 0, width, height, "F");
  }

  function loadImage(url: string): Promise<HTMLImageElement> {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.crossOrigin = "anonymous";
      img.decoding = "async";
      img.loading = "eager" as any;
      img.onload = () => resolve(img);
      img.onerror = (e) => reject(e);
      img.src = url;
    });
  }

  async function loadImageAsPngDataUrl(url: string): Promise<{ dataUrl: string; width: number; height: number }> {
    const img = await loadImage(url);
    const canvas = document.createElement("canvas");
    const scale = 2;
    canvas.width = img.naturalWidth * scale;
    canvas.height = img.naturalHeight * scale;
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("Canvas 2D context not available");
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = "high" as CanvasImageSmoothingQuality;
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
    const dataUrl = canvas.toDataURL("image/png");
    return { dataUrl, width: img.naturalWidth, height: img.naturalHeight };
  }

  function writeParagraph(doc: jsPDF, text: string, x: number, y: number, maxWidth: number, lineHeight = 14): number {
    const lines = doc.splitTextToSize(text || "", maxWidth) as string[];
    lines.forEach((line, i) => doc.text(line, x, y + i * lineHeight));
    return y + lines.length * lineHeight;
  }

  function drawFunctionDashboard(doc: jsPDF, averages: Record<string, number>, x: number, y: number, width: number, barHeight: number, color: [number, number, number]) {
    const entries = Object.entries(averages);
    const gap = 10;
    const maxScore = 4;
    for (let i = 0; i < entries.length; i++) {
      const [label, score] = entries[i];
      const pct = Math.max(0, Math.min(1, score / maxScore));
      const bw = width * pct;
      const by = y + i * (barHeight + gap);
      doc.setDrawColor(color[0], color[1], color[2]);
      doc.setTextColor(color[0], color[1], color[2]);
      doc.setFillColor(30, 30, 30);
      doc.roundedRect(x, by, width, barHeight, 3, 3, "F");
      doc.setFillColor(color[0], color[1], color[2]);
      // @ts-expect-error GState exists at runtime
      doc.setGState(new (doc as any).GState({ opacity: 0.35 }));
      doc.roundedRect(x, by, bw, barHeight, 3, 3, "F");
      // @ts-expect-error GState exists at runtime
      (doc as any).setGState(new (doc as any).GState({ opacity: 1 }));
      doc.setFontSize(10);
      doc.text(`${label} (${score.toFixed(2)}/4)`, x + 6, by + barHeight - 4);
    }
  }

  function classifyRisk(m: number): "Sustainable" | "Moderate" | "Severe" | "Critical" {
    if (m >= 4) return "Sustainable";
    if (m === 3) return "Moderate";
    if (m === 2) return "Severe";
    return "Critical";
  }

  function computeFunctionAveragesLocal(responses: { meta: { nistFunction: string }; maturity: number }[]) {
    const acc: Record<string, { total: number; count: number }> = {};
    for (const r of responses) {
      const k = r.meta.nistFunction;
      if (!acc[k]) acc[k] = { total: 0, count: 0 };
      acc[k].total += r.maturity;
      acc[k].count += 1;
    }
    const out: Record<string, number> = {};
    Object.entries(acc).forEach(([k, v]) => (out[k] = v.count ? v.total / v.count : 0));
    return out;
  }

  function drawComplianceTable(doc: jsPDF, avgs: Record<string, number>, pr: number, pg: number, pb: number, paintedPages: Set<number>) {
    const observation = (fn: string, avg: number) => {
      if (avg < 1.5) {
        switch (fn) {
          case "GOVERN": return "Minimal leadership engagement, risk tolerance undefined";
          case "IDENTIFY": return "Asset inventory incomplete, weak data flow mapping";
          case "PROTECT": return "Controls inconsistent, MFA/segmentation gaps";
          case "DETECT": return "No centralized monitoring or alerting";
          case "RESPOND": return "Basic IR plan, not cloud-specific";
          case "RECOVER": return "DR planning immature";
          default: return "Needs improvement";
        }
      }
      if (avg < 2.5) return "Developing capabilities with notable gaps";
      return "Solid capability with room to optimize";
    };
    const rows = Object.entries(avgs).map(([fn, avg]) => [fn, avg.toFixed(2), observation(fn, avg)]);
    autoTable(doc, {
      startY: (doc as any).lastAutoTable?.finalY ? (doc as any).lastAutoTable.finalY + 12 : 100,
      theme: "grid",
      headStyles: { fillColor: [0, 0, 0], textColor: [255, 255, 255] },
      bodyStyles: { cellPadding: 6, textColor: [255, 255, 255], fillColor: [0, 0, 0] },
      styles: { fontSize: 10, lineColor: [pr, pg, pb], lineWidth: 0.5, fillColor: [0, 0, 0] },
      head: [["Function", "Maturity Score", "Observations"]],
      body: rows,
      willDrawCell: () => {
        const page = (doc as any).internal.getCurrentPageInfo().pageNumber;
        if (!paintedPages.has(page)) {
          drawPdfBackground(doc);
          paintedPages.add(page);
        }
      },
    });
  }

  function computeRiskCounts(responses: { maturity: number }[]) {
    const levels = ["Sustainable", "Moderate", "Severe", "Critical"] as const;
    const inherent: Record<typeof levels[number], number> = { Sustainable: 0, Moderate: 0, Severe: 0, Critical: 0 };
    const residual: Record<typeof levels[number], number> = { Sustainable: 0, Moderate: 0, Severe: 0, Critical: 0 };
    for (const r of responses) {
      (inherent as any)[classifyRisk(r.maturity)] += 1;
      (residual as any)[classifyRisk(Math.min(4, r.maturity + 1))] += 1;
    }
    return { inherent, residual, levels: Array.from(levels) };
  }

  function drawRiskDashboards(doc: jsPDF, counts: { inherent: Record<string, number>; residual: Record<string, number>; levels: string[] }, x: number, startY: number, pr: number, pg: number, pb: number, paintedPages: Set<number>) {
    const total = Object.values(counts.inherent).reduce((a, b) => a + b, 0) || 1;
    const toRows = (obj: Record<string, number>) => counts.levels.map(l => [l, String(obj[l] || 0), `${((100 * (obj[l] || 0)) / total).toFixed(2)}%`]);

    doc.setFontSize(12);
    doc.setTextColor(pr, pg, pb);
    doc.text("Risk Dashboards", x, startY);

    autoTable(doc, {
      startY: startY + 10,
      theme: "grid",
      headStyles: { fillColor: [0, 0, 0], textColor: [255, 255, 255] },
      bodyStyles: { cellPadding: 6, textColor: [255, 255, 255], fillColor: [0, 0, 0] },
      styles: { fontSize: 10, lineColor: [pr, pg, pb], lineWidth: 0.5, fillColor: [0, 0, 0] },
      head: [["Risk Level", "Count", "Percentage"]],
      body: toRows(counts.inherent),
      willDrawCell: () => {
        const page = (doc as any).internal.getCurrentPageInfo().pageNumber;
        if (!paintedPages.has(page)) { drawPdfBackground(doc); paintedPages.add(page); }
      },
    });

    autoTable(doc, {
      startY: (doc as any).lastAutoTable.finalY + 12,
      theme: "grid",
      headStyles: { fillColor: [0, 0, 0], textColor: [255, 255, 255] },
      bodyStyles: { cellPadding: 6, textColor: [255, 255, 255], fillColor: [0, 0, 0] },
      styles: { fontSize: 10, lineColor: [pr, pg, pb], lineWidth: 0.5, fillColor: [0, 0, 0] },
      head: [["Risk Level", "Count", "Percentage"]],
      body: toRows(counts.residual),
      willDrawCell: () => {
        const page = (doc as any).internal.getCurrentPageInfo().pageNumber;
        if (!paintedPages.has(page)) { drawPdfBackground(doc); paintedPages.add(page); }
      },
    });
  }

  function drawRiskMatrixOverview(doc: jsPDF, x: number, startY: number, pr: number, pg: number, pb: number, paintedPages: Set<number>) {
    const probs = ["Low", "Medium", "Medium-High", "High", "Critical"];
    const impacts = ["Sustainable", "Moderate", "Severe", "Critical"];
    const severity = ["Sustainable", "Moderate", "Severe", "Critical"];
    const head = ["Probability → / Impact ↓", ...impacts];
    const body = probs.map((p, pi) => {
      const row = [p];
      for (let i = 0; i < impacts.length; i++) {
        const s = severity[Math.min(3, Math.max(i, pi - 1))];
        row.push(s);
      }
      return row;
    });
    doc.setFontSize(12);
    doc.setTextColor(pr, pg, pb);
    doc.text("Risk Matrix Overview", x, startY);
    autoTable(doc, {
      startY: startY + 10,
      theme: "grid",
      headStyles: { fillColor: [0, 0, 0], textColor: [255, 255, 255] },
      bodyStyles: { cellPadding: 6, textColor: [255, 255, 255], fillColor: [0, 0, 0] },
      styles: { fontSize: 10, lineColor: [pr, pg, pb], lineWidth: 0.5, fillColor: [0, 0, 0] },
      head: [head],
      body,
      willDrawCell: () => {
        const page = (doc as any).internal.getCurrentPageInfo().pageNumber;
        if (!paintedPages.has(page)) { drawPdfBackground(doc); paintedPages.add(page); }
      },
    });
  }

  async function onSubmit(values: FormValues) {
    const submission: AssessmentSubmission = {
      organization: values.organization,
      contactEmail: values.contactEmail,
      submittedAt: new Date().toISOString(),
      responses: values.responses.map((r, i) => ({
        meta: questions[i],
        response: r.response,
        maturity: r.maturity as 0 | 1 | 2 | 3 | 4,
      })),
    };

    const doc = new jsPDF({ unit: "pt", format: "a4" });
    const [pr, pg, pb] = getBrandRgb();

    // Track which pages have had backgrounds painted
    const paintedPages = new Set<number>();

    // Background (black)
    drawPdfBackground(doc);
    paintedPages.add((doc as any).internal.getCurrentPageInfo().pageNumber);

    // --- Title page ---
    try {
      const logoUrl = "https://cdn.builder.io/api/v1/image/assets%2F74452fbd65844fa092de7a3dcf4c1086%2Fdfa2f0d5c3b54a6583627c5690a0e221?format=webp&width=1600";
      const { dataUrl, width: iw, height: ih } = await loadImageAsPngDataUrl(logoUrl);
      const pw = doc.internal.pageSize.getWidth();
      const ph = doc.internal.pageSize.getHeight();
      const imgW = Math.min(300, pw - 120);
      const ratio = ih / iw;
      const imgH = imgW * ratio;
      const x = (pw - imgW) / 2;
      const y = ph * 0.25;
      // Draw high-quality PNG (preserves transparency, smoother scaling)
      doc.addImage(dataUrl, "PNG", x, y, imgW, imgH);
      // Title text centered
      doc.setTextColor(pr, pg, pb);
      doc.setFontSize(22);
      doc.text("Olezka Global", pw / 2, y + imgH + 48, { align: "center" });
      doc.setFontSize(14);
      doc.text(
        "Cloud Security Posture Assessment",
        pw / 2,
        y + imgH + 72,
        { align: "center" },
      );
    } catch {}

    // Start content on a new page
    doc.addPage();
    drawPdfBackground(doc);
    paintedPages.add((doc as any).internal.getCurrentPageInfo().pageNumber);

    // Title
    doc.setFontSize(20);
    doc.setTextColor(pr, pg, pb);
    doc.text("Olezka Global", 40, 50);
    doc.setFontSize(12);
    doc.setTextColor(pr, pg, pb);
    doc.text("Cloud Security Posture Assessment", 40, 70);

    // Org info box
    doc.setDrawColor(pr, pg, pb);
    doc.setFillColor(pr, pg, pb);
    (doc as any).setGState(new (doc as any).GState({ opacity: 0.12 }));
    doc.roundedRect(32, 85, 530, 48, 6, 6, "F");
    (doc as any).setGState(new (doc as any).GState({ opacity: 1 }));
    doc.setTextColor(255, 255, 255);
    doc.text(`Organization: ${submission.organization || "-"}`, 44, 110);
    doc.text(`Contact Email: ${submission.contactEmail || "-"}`, 300, 110);

    let cursorY = 150;

    // --- AI Executive Summary ---
    let aiRendered = false;
    try {
      const aiResp = await fetch("/api/assessments/generate-report", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(submission),
      });
      if (!aiResp.ok) {
        await aiResp.text();
        toast.error("AI analysis failed (quota or key). Using local summary.");
      } else {
        const { report, metrics } = (await aiResp.json()) as any;
        doc.setFontSize(16);
        doc.setTextColor(pr, pg, pb);
        doc.text("Executive Summary", 40, cursorY);
        doc.setFontSize(11);
        doc.setTextColor(255, 255, 255);
        cursorY = writeParagraph(doc, report.summary || "", 40, cursorY + 20, 520, 14);

        if (report.environment) {
          doc.setFontSize(12);
          doc.setTextColor(pr, pg, pb);
          doc.text("Environment", 40, cursorY + 18);
          doc.setTextColor(255, 255, 255);
          cursorY = writeParagraph(doc, String(report.environment), 40, cursorY + 34, 520);
        }

        if (Array.isArray(report.strengths) && report.strengths.length) {
          doc.setFontSize(12);
          doc.setTextColor(pr, pg, pb);
          doc.text("Key Strengths", 40, cursorY + 18);
          doc.setTextColor(255, 255, 255);
          let y = cursorY + 34;
          report.strengths.forEach((s: string) => { y = writeParagraph(doc, `• ${s}`, 44, y, 516); });
          cursorY = y;
        }

        if (Array.isArray(report.gaps) && report.gaps.length) {
          doc.setFontSize(12);
          doc.setTextColor(pr, pg, pb);
          doc.text("Identified Gaps & Risks", 40, cursorY + 18);
          doc.setTextColor(255, 255, 255);
          let y = cursorY + 34;
          report.gaps.slice(0, 6).forEach((r: any) => { y = writeParagraph(doc, `• ${r.name} — ${r.description || ""}`, 44, y, 516); });
          cursorY = y;
        }

        doc.addPage();
        drawPdfBackground(doc);
        paintedPages.add((doc as any).internal.getCurrentPageInfo().pageNumber);
        cursorY = 60;

        doc.setFontSize(12);
        doc.setTextColor(pr, pg, pb);
        doc.text("Compliance Alignment (NIST CSF)", 40, cursorY);
        doc.setTextColor(255, 255, 255);
        cursorY = writeParagraph(doc, report.compliance || "", 40, cursorY + 16, 520);
        drawComplianceTable(doc, metrics?.functionAverages || {}, pr, pg, pb, paintedPages);
        cursorY = (doc as any).lastAutoTable.finalY + 10;

        if (Array.isArray(report.recommendations) && report.recommendations.length) {
          doc.setFontSize(12);
          doc.setTextColor(pr, pg, pb);
          doc.text("Strategic Recommendations", 40, cursorY + 18);
          doc.setTextColor(255, 255, 255);
          let y = cursorY + 34;
          report.recommendations.forEach((rec: any) => {
            const item = typeof rec === "string" ? rec : rec.action;
            y = writeParagraph(doc, `• ${item}`, 44, y, 516);
          });
          cursorY = y;
        }

        if (Array.isArray(report.gaps) && report.gaps.length) {
          doc.setFontSize(12);
          doc.setTextColor(pr, pg, pb);
          doc.text("Risk Matrix", 40, cursorY + 24);
          autoTable(doc, {
            startY: cursorY + 30,
            theme: "grid",
            headStyles: { fillColor: [0, 0, 0], textColor: [255, 255, 255] },
            bodyStyles: { cellPadding: 6, textColor: [255, 255, 255], fillColor: [0, 0, 0] },
            styles: { fontSize: 10, lineColor: [pr, pg, pb], lineWidth: 0.5, fillColor: [0, 0, 0] },
            head: [["Risk", "Likelihood", "Impact", "Rating"]],
            body: report.gaps.map((g: any) => [g.name, g.likelihood || "Medium", g.impact || "Medium", g.rating || "Medium"]),
            willDrawCell: () => {
              const page = (doc as any).internal.getCurrentPageInfo().pageNumber;
              if (!paintedPages.has(page)) { drawPdfBackground(doc); paintedPages.add(page); }
            },
          });
          cursorY = (doc as any).lastAutoTable.finalY + 10;
        }

        doc.setFontSize(12);
        doc.setTextColor(pr, pg, pb);
        doc.text("Dashboard (NIST Function Averages)", 40, cursorY + 24);
        drawFunctionDashboard(doc, metrics?.functionAverages || {}, 40, cursorY + 34, 520, 14, [pr, pg, pb]);
        cursorY = cursorY + 34 + (Object.keys(metrics?.functionAverages || {}).length || 6) * 24;

        drawRiskMatrixOverview(doc, 40, cursorY + 24, pr, pg, pb, paintedPages);
        cursorY = (doc as any).lastAutoTable.finalY + 10;

        const counts = computeRiskCounts(submission.responses);
        drawRiskDashboards(doc, counts, 40, cursorY + 24, pr, pg, pb, paintedPages);
        cursorY = (doc as any).lastAutoTable.finalY + 10;

        doc.addPage();
        drawPdfBackground(doc);
        paintedPages.add((doc as any).internal.getCurrentPageInfo().pageNumber);
        cursorY = 60;
        doc.setFontSize(12);
        doc.setTextColor(pr, pg, pb);
        doc.text("Conclusion & Next Steps", 40, cursorY);
        doc.setTextColor(255, 255, 255);
        cursorY = writeParagraph(doc, report.conclusion || "", 40, cursorY + 16, 520);
        aiRendered = true;
      }
    } catch (e) {
      toast.error("AI unavailable; generating local summary.");
    }

    if (!aiRendered) {
      const avgs = computeFunctionAveragesLocal(submission.responses);
      const overall = Object.values(avgs).reduce((a, b) => a + b, 0) / (Object.keys(avgs).length || 1);
      const env = "Azure + Microsoft 365";
      const strengths = Object.entries(avgs).filter(([, v]) => v >= 2.5).map(([k]) => `${k}: maturing capability`);
      const gaps = Object.entries(avgs).filter(([, v]) => v < 1.5).map(([k]) => ({ name: `${k} capability gap`, likelihood: "Medium", impact: "High", rating: "Severe" }));
      const recommendations = [
        "Close MFA and conditional access gaps; enforce admin protections",
        "Centralize logging to Sentinel and tune detections",
        "Harden data protection with DLP coverage and private endpoints",
        "Formalize cloud-specific IR and test DR scenarios",
      ];
      const summary = `This assessment summarizes your current cloud security posture. Average maturity is ${overall.toFixed(2)} of 4. Stronger areas include ${strengths.map(s=>s.split(":")[0]).join(", ") || "none"}. Lower-maturity areas should be prioritized to reduce exposure and align with NIST CSF.`;
      const conclusion = `Improving low-maturity functions will reduce residual risk and support business continuity. Prioritize the recommendations over the next 90 days and track progress via KPIs.`;

      doc.setFontSize(16);
      doc.setTextColor(pr, pg, pb);
      doc.text("Executive Summary", 40, cursorY);
      doc.setFontSize(11);
      doc.setTextColor(255, 255, 255);
      cursorY = writeParagraph(doc, summary, 40, cursorY + 20, 520, 14);

      doc.setFontSize(12);
      doc.setTextColor(pr, pg, pb);
      doc.text("Environment", 40, cursorY + 18);
      doc.setTextColor(255, 255, 255);
      cursorY = writeParagraph(doc, env, 40, cursorY + 34, 520);

      if (strengths.length) {
        doc.setFontSize(12);
        doc.setTextColor(pr, pg, pb);
        doc.text("Key Strengths", 40, cursorY + 18);
        doc.setTextColor(255, 255, 255);
        let y = cursorY + 34;
        strengths.forEach((s: string) => { y = writeParagraph(doc, `• ${s}`, 44, y, 516); });
        cursorY = y;
      }

      if (gaps.length) {
        doc.setFontSize(12);
        doc.setTextColor(pr, pg, pb);
        doc.text("Identified Gaps & Risks", 40, cursorY + 18);
        doc.setTextColor(255, 255, 255);
        let y = cursorY + 34;
        gaps.slice(0, 6).forEach((r: any) => { y = writeParagraph(doc, `• ${r.name}`, 44, y, 516); });
        cursorY = y;
      }

      doc.addPage();
      drawPdfBackground(doc);
      paintedPages.add((doc as any).internal.getCurrentPageInfo().pageNumber);
      cursorY = 60;

      doc.setFontSize(12);
      doc.setTextColor(pr, pg, pb);
      doc.text("Compliance Alignment (NIST CSF)", 40, cursorY);
      doc.setTextColor(255, 255, 255);
      cursorY = writeParagraph(doc, "The table below summarizes average maturity by NIST function with brief observations.", 40, cursorY + 16, 520);
      drawComplianceTable(doc, avgs, pr, pg, pb, paintedPages);
      cursorY = (doc as any).lastAutoTable.finalY + 10;

      doc.setFontSize(12);
      doc.setTextColor(pr, pg, pb);
      doc.text("Strategic Recommendations", 40, cursorY + 18);
      doc.setTextColor(255, 255, 255);
      let y2 = cursorY + 34;
      recommendations.forEach((rec: string) => { y2 = writeParagraph(doc, `• ${rec}`, 44, y2, 516); });
      cursorY = y2;

      drawRiskMatrixOverview(doc, 40, cursorY + 24, pr, pg, pb, paintedPages);
      cursorY = (doc as any).lastAutoTable.finalY + 10;

      const counts = computeRiskCounts(submission.responses);
      drawRiskDashboards(doc, counts, 40, cursorY + 24, pr, pg, pb, paintedPages);
      cursorY = (doc as any).lastAutoTable.finalY + 10;

      doc.addPage();
      drawPdfBackground(doc);
      paintedPages.add((doc as any).internal.getCurrentPageInfo().pageNumber);
      cursorY = 60;
      doc.setFontSize(12);
      doc.setTextColor(pr, pg, pb);
      doc.text("Conclusion & Next Steps", 40, cursorY);
      doc.setTextColor(255, 255, 255);
      cursorY = writeParagraph(doc, conclusion, 40, cursorY + 16, 520);
    }

    const functions: NistFunction[] = [
      "GOVERN",
      "IDENTIFY",
      "PROTECT",
      "DETECT",
      "RESPOND",
      "RECOVER",
    ];

    for (const fn of functions) {
      const items = submission.responses.filter(
        (r) => r.meta.nistFunction === fn,
      );
      if (!items.length) continue;

      // Section header
      doc.setFillColor(pr, pg, pb);
      (doc as any).setGState(new (doc as any).GState({ opacity: 0.15 }));
      doc.rect(32, cursorY - 18, 530, 28, "F");
      (doc as any).setGState(new (doc as any).GState({ opacity: 1 }));
      doc.setTextColor(pr, pg, pb);
      doc.setFontSize(13);
      doc.text(fn, 44, cursorY);

      // Table for this section
      autoTable(doc, {
        startY: cursorY + 14,
        theme: "grid",
        headStyles: { fillColor: [0, 0, 0], textColor: [255, 255, 255] },
        bodyStyles: { cellPadding: 6, textColor: [255, 255, 255], fillColor: [0, 0, 0] },
        styles: { fontSize: 10, lineColor: [pr, pg, pb], lineWidth: 0.5, fillColor: [0, 0, 0] },
        alternateRowStyles: { fillColor: [0, 0, 0] },
        columnStyles: {
          0: { cellWidth: 70 },
          1: { cellWidth: 180 },
          2: { cellWidth: 220 },
          3: { cellWidth: 60, halign: "center" },
        },
        head: [["Control", "Prompt", "Response", "Maturity"]],
        body: items.map((it) => [
          it.meta.id,
          it.meta.prompt,
          it.response || "",
          String(it.maturity),
        ]),
        willDrawCell: (data: any) => {
          // Paint page background BEFORE any cell on a new page
          const page = (doc as any).internal.getCurrentPageInfo().pageNumber;
          if (!paintedPages.has(page)) {
            drawPdfBackground(doc);
            paintedPages.add(page);
          }
          if (data.section === 'head') {
            doc.setFillColor(pr, pg, pb);
            (doc as any).setGState(new (doc as any).GState({ opacity: 0.15 }));
            doc.rect(data.cell.x, data.cell.y, data.cell.width, data.cell.height, 'F');
            (doc as any).setGState(new (doc as any).GState({ opacity: 1 }));
          }
        },
      });

      cursorY = (doc as any).lastAutoTable.finalY + 28;
      if (cursorY > 720) {
        doc.addPage();
        drawPdfBackground(doc);
        paintedPages.add((doc as any).internal.getCurrentPageInfo().pageNumber);
        cursorY = 60;
      }
    }

    // --- Summary page ---
    const totalQuestions = submission.responses.length;
    const totalScore = submission.responses.reduce((acc, r) => acc + r.maturity, 0);
    const maxScore = totalQuestions * 4;
    const avg = totalQuestions ? totalScore / totalQuestions : 0;
    const pct = Math.round((avg / 4) * 100);

    doc.addPage();
    drawPdfBackground(doc);
    paintedPages.add((doc as any).internal.getCurrentPageInfo().pageNumber);

    doc.setFontSize(18);
    doc.setTextColor(pr, pg, pb);
    doc.text("Overall Maturity Score", 40, 60);

    // Summary box
    doc.setDrawColor(pr, pg, pb);
    doc.setFillColor(pr, pg, pb);
    // transparent teal card
    // @ts-expect-error - GState exists at runtime
    doc.setGState(new (doc as any).GState({ opacity: 0.12 }));
    doc.roundedRect(32, 80, 530, 120, 6, 6, "F");
    // @ts-expect-error - GState exists at runtime
    (doc as any).setGState(new (doc as any).GState({ opacity: 1 }));
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(12);
    doc.text(`Total Points: ${totalScore} / ${maxScore}`, 44, 110);
    doc.text(`Average Rating (0-4): ${avg.toFixed(2)}`, 44, 136);
    doc.text(`Overall Percentage: ${pct}%`, 44, 162);

    doc.save("Olezka-Assessment.pdf");
    toast.success("PDF downloaded");
  }

  return (
    <div className="min-h-screen relative">
      <header className="sticky top-0 z-40 border-b bg-background/70 backdrop-blur">
        <div className="container mx-auto flex items-center justify-between py-5">
          <Link
            to="/"
            className="flex items-center gap-3"
            aria-label="Olezka Global home"
          >
            <img
              src="https://cdn.builder.io/api/v1/image/assets%2F74452fbd65844fa092de7a3dcf4c1086%2Fe5b839fd3d154ffbb8519dfbf58af002?format=webp&width=800"
              alt="Olezka Global"
              className="h-9 md:h-10 w-auto object-contain"
              loading="eager"
              decoding="async"
            />
            <span className="text-lg font-semibold tracking-tight">
              Olezka Global
            </span>
          </Link>
          <div className="text-xs text-muted-foreground">
            Cloud Security Posture Assessment
          </div>
        </div>
      </header>

      <main className="container mx-auto px-4 py-10">
        <div className="mb-10 grid gap-4 lg:grid-cols-3">
          <div className="lg:col-span-2">
            <h1 className="text-3xl md:text-4xl font-extrabold tracking-tight">
              Cloud Security Posture Assessment
            </h1>
            <p className="mt-3 text-muted-foreground max-w-2xl">
              Complete this questionnaire to evaluate your Azure and Microsoft
              365 security posture. Responses are organized by NIST functions
              and controls. A maturity rating of 0-4 is required for each item.
            </p>
          </div>
          <Card className="bg-gradient-to-br from-primary/10 to-transparent border-primary/30">
            <CardHeader>
              <CardTitle>About Your Organization</CardTitle>
              <CardDescription>
                Provide basic details so we can associate your assessment.
              </CardDescription>
            </CardHeader>
            <CardContent className="grid gap-4">
              <div className="grid gap-2">
                <Label htmlFor="organization">Organization</Label>
                <Input
                  id="organization"
                  placeholder="Company or Institution Name"
                  {...register("organization")}
                />
              </div>
              <div className="grid gap-2">
                <Label htmlFor="email">Contact Email</Label>
                <Input
                  id="email"
                  placeholder="you@company.com"
                  {...register("contactEmail")}
                />
              </div>
            </CardContent>
          </Card>
        </div>

        <form onSubmit={handleSubmit(onSubmit)} className="grid gap-8">
          <Accordion type="multiple" className="grid gap-4">
            {(Object.keys(grouped) as NistFunction[]).map((fn) => (
              <AccordionItem
                key={fn}
                value={fn}
                className="rounded-md border border-border/60"
              >
                <AccordionTrigger className="px-4">
                  <div className="flex items-center gap-3">
                    <span className="inline-flex h-6 w-6 items-center justify-center rounded-sm bg-primary/20 text-primary text-xs font-semibold">
                      {fn[0]}
                    </span>
                    <span className="text-left">{fn}</span>
                  </div>
                </AccordionTrigger>
                <AccordionContent className="px-4">
                  <div className="grid gap-6">
                    {grouped[fn].map((q) => {
                      const idx = questions.findIndex((qq) => qq.id === q.id);
                      return (
                        <Card key={q.id} className="border-border/60">
                          <CardHeader className="pb-3">
                            <CardTitle className="text-base flex items-center justify-between">
                              <span>
                                {q.id} • {q.category}
                              </span>
                              <span className="text-xs text-muted-foreground font-normal">
                                {q.control}
                              </span>
                            </CardTitle>
                            <CardDescription className="text-sm">
                              {q.prompt}
                            </CardDescription>
                          </CardHeader>
                          <CardContent className="grid gap-4">
                            <div className="grid gap-2">
                              <Label htmlFor={`responses.${idx}.response`}>
                                Client's Response
                              </Label>
                              <Textarea
                                id={`responses.${idx}.response`}
                                placeholder="Provide details, references, and current state"
                                {...register(
                                  `responses.${idx}.response` as const,
                                )}
                              />
                            </div>
                            <div className="grid gap-2 max-w-xs">
                              <Label>Maturity Rating (0-4)</Label>
                              <Select
                                defaultValue={`${fields[idx]?.maturity ?? 0}`}
                                onValueChange={(v) => {
                                  const num = Number(v) as 0 | 1 | 2 | 3 | 4;
                                  setValue(`responses.${idx}.maturity` as const, num, { shouldDirty: true, shouldValidate: true });
                                  const input = document.getElementById(
                                    `responses.${idx}.maturity-hidden`,
                                  ) as HTMLInputElement | null;
                                  if (input) {
                                    input.value = String(num);
                                    input.dispatchEvent(new Event("input", { bubbles: true }));
                                    input.dispatchEvent(new Event("change", { bubbles: true }));
                                  }
                                }}
                              >
                                <SelectTrigger>
                                  <SelectValue placeholder="Select rating" />
                                </SelectTrigger>
                                <SelectContent>
                                  <SelectItem value="0">0 - None</SelectItem>
                                  <SelectItem value="1">1 - Initial</SelectItem>
                                  <SelectItem value="2">
                                    2 - Developing
                                  </SelectItem>
                                  <SelectItem value="3">3 - Defined</SelectItem>
                                  <SelectItem value="4">4 - Managed</SelectItem>
                                </SelectContent>
                              </Select>
                              <input
                                id={`responses.${idx}.maturity-hidden`}
                                type="hidden"
                                {...register(
                                  `responses.${idx}.maturity` as const,
                                  {
                                    valueAsNumber: true,
                                  },
                                )}
                              />
                            </div>
                          </CardContent>
                        </Card>
                      );
                    })}
                  </div>
                </AccordionContent>
              </AccordionItem>
            ))}
          </Accordion>

          <div className="flex items-center justify-end gap-3">
            <Button type="button" variant="outline" onClick={addRealData}>
              Add Real Data
            </Button>
            <Button type="submit" disabled={isSubmitting} className="gap-2">
              <Download className="h-4 w-4" /> Download Assessment
            </Button>
          </div>
        </form>
      </main>

      <footer className="border-t mt-10">
        <div className="container mx-auto px-4 py-6 text-xs text-muted-foreground flex items-center justify-between">
          <span className="text-primary">© {new Date().getFullYear()} Olezka Global</span>
          <a href="https://www.olezkaglobal.com" className="text-primary hover:underline" target="_blank" rel="noopener noreferrer">www.olezkaglobal.com</a>
          <span className="text-primary">Secure by design</span>
        </div>
      </footer>
    </div>
  );
}
