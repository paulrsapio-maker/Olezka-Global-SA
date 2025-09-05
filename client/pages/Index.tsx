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

  function addTestData() {
    setValue("organization", "Acme University");
    setValue("contactEmail", "security@acme.edu");
    questions.forEach((q, i) => {
      setValue(`responses.${i}.response` as const, `Sample response for ${q.id}: controls are documented and reviewed quarterly.`);
      const m = (i % 5) as 0 | 1 | 2 | 3 | 4;
      setValue(`responses.${i}.maturity` as const, m);
      const hidden = document.getElementById(`responses.${i}.maturity-hidden`) as HTMLInputElement | null;
      if (hidden) hidden.value = String(m);
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
    // Render at 2x for crisper downscaling in PDF
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
                                  // React Hook Form manual set using hidden input binding
                                  const input = document.getElementById(
                                    `responses.${idx}.maturity-hidden`,
                                  ) as HTMLInputElement | null;
                                  if (input) input.value = v;
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
            <Button type="button" variant="outline" onClick={addTestData}>
              Add Test Data
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
