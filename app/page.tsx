"use client";

import { useState } from "react";
import {
  Rocket,
  Loader2,
  ClipboardList,
  Download,
  Presentation,
  FileText,
  Sheet,
  TrendingUp,
  PackageCheck,
  CheckCircle2,
  Ticket,
  X,
} from "lucide-react";

const SOURCE_SHEETS = [
  {
    name: "Support Tickets",
    url: "https://docs.google.com/spreadsheets/d/1jEntEdMkQOG3W_oQ5TOTmTfQAaPS1yeeANSZhGCe5SM/export?format=csv&gid=0",
  },
  {
    name: "Call Transcripts (Gong)",
    url: "https://docs.google.com/spreadsheets/d/1oM89T5WCEMSByF6QPoQ7QH0_zbOUGTsAKYpniL2QMLk/export?format=csv&gid=0",
  },
  {
    name: "CRM & Revenue",
    url: "https://docs.google.com/spreadsheets/d/1DXPFu-d1V_3Ys7gsJQbITL5lZZLnNPqFzXiYlN7sBhA/export?format=csv&gid=0",
  },
  {
    name: "Data Analytics",
    url: "https://docs.google.com/spreadsheets/d/1X5deyLNK3rD-_xG7bW-AS9-1cCOiUg29l9rBWytbrps/export?format=csv&gid=0",
  },
];

const pendingTickets = [
  {
    company_id: "Comp_037",
    workflow: "dashboard_load",
    feedback: "The dashboard is completely down and our finance team cannot close the month.",
  },
  {
    company_id: "Comp_034",
    workflow: "user_onboarding",
    feedback: "The NetSuite sync during onboarding fails intermittently with error 504 for some new users.",
  },
  {
    company_id: "Comp_012",
    workflow: "card_transaction_auth",
    feedback: "Card approvals sometimes take 8 to 10 seconds, which frustrates our drivers at the pump.",
  },
  {
    company_id: "Comp_012",
    workflow: "card_transaction_auth",
    feedback: "It would be great to export card transactions in a custom CSV format.",
  },
  {
    company_id: "Comp_026",
    workflow: "card_transaction_auth",
    feedback: "How do I change the spending limit for a single card?",
  },
];

interface TriageApiResponse {
  feature?: string;
  priority: string;
  status: string;
  rationale: string;
  core_workflow_tracked?: string;
  company_name?: string;
  user_name?: string;
  user_id?: string;
  mrr?: string | number;
  subscription_tier?: string;
  renewal_date?: string;
  csm_health_score?: string | number;
  rice_reach?: number;
  rice_impact?: number;
  rice_confidence?: number;
  rice_effort?: number;
  rice_score?: number;
  calibration_factor?: number;
  outcomes_used?: number;
  issue_type?: string;
  prediction_id?: string;
}

interface TriageRow extends TriageApiResponse {
  feature: string;
  feedback: string;
  workflow: string;
  company_id: string;
}

type ColumnKey =
  | "feedback"
  | "core_workflow"
  | "priority"
  | "status"
  | "rice"
  | "rationale"
  | "company_id";

const COLUMNS: { key: ColumnKey; label: string }[] = [
  { key: "feedback", label: "Customer Request" },
  { key: "core_workflow", label: "Core Workflow Tracked / Analytics" },
  { key: "priority", label: "Priority" },
  { key: "status", label: "Status" },
  { key: "rice", label: "RICE Score" },
  { key: "rationale", label: "Rationale" },
  { key: "company_id", label: "Company ID" },
];

function getCellValue(row: TriageRow, key: ColumnKey): string {
  switch (key) {
    case "feedback":
      return row.feedback;
    case "core_workflow":
      return row.core_workflow_tracked || row.workflow;
    case "priority":
      return row.priority;
    case "status":
      return row.status;
    case "rice":
      return row.rice_score === undefined
        ? "—"
        : `${row.rice_score} (R ${row.rice_reach}, I ${row.rice_impact}, C ${row.rice_confidence}, E ${row.rice_effort})`;
    case "rationale":
      return row.rationale;
    case "company_id":
      return row.company_id;
    default:
      return "—";
  }
}

function PriorityBadge({ priority }: { priority: string }) {
  const normalized = priority.toUpperCase();
  const styles: Record<string, string> = {
    P0: "bg-red-100 text-red-700 border-red-200",
    P1: "bg-orange-100 text-orange-700 border-orange-200",
    P2: "bg-yellow-100 text-yellow-700 border-yellow-200",
    P3: "bg-gray-100 text-gray-700 border-gray-200",
    "N/A": "bg-yellow-50 text-yellow-700 border-yellow-200 border-dashed",
  };
  const classes = styles[normalized] ?? "bg-gray-100 text-gray-700 border-gray-200";
  return (
    <span
      className={`inline-flex items-center whitespace-nowrap rounded-full border px-2.5 py-0.5 text-xs font-semibold ${classes}`}
    >
      {priority}
    </span>
  );
}

function StatusBadge({ status }: { status: string }) {
  const normalized = status.toUpperCase();
  const styles: Record<string, string> = {
    "TO BE PICKED": "bg-green-100 text-green-700 border-green-200",
    NOISE: "bg-gray-100 text-gray-600 border-gray-200",
    "NO DATA": "bg-yellow-50 text-yellow-700 border-yellow-200 border-dashed",
  };
  const classes =
    styles[normalized] ?? "bg-blue-100 text-blue-700 border-blue-200";
  return (
    <span
      className={`inline-flex items-center whitespace-nowrap rounded-full border px-2.5 py-0.5 text-xs font-semibold ${classes}`}
    >
      {status}
    </span>
  );
}

declare global {
  interface Window {
    PptxGenJS?: new () => any;
  }
}

let pptxGenJSLoader: Promise<new () => any> | null = null;

function loadPptxGenJS(): Promise<new () => any> {
  if (window.PptxGenJS) {
    return Promise.resolve(window.PptxGenJS);
  }
  if (!pptxGenJSLoader) {
    pptxGenJSLoader = new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = "/libs/pptxgen.bundle.js";
      script.onload = () => {
        if (window.PptxGenJS) {
          resolve(window.PptxGenJS);
        } else {
          reject(new Error("PptxGenJS failed to load"));
        }
      };
      script.onerror = () => reject(new Error("Failed to load PptxGenJS script"));
      document.head.appendChild(script);
    });
  }
  return pptxGenJSLoader;
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function escapeCsvValue(value: string): string {
  if (/[",\n]/.test(value)) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
}

const ROWS_PER_PAGE = 5;

export default function Home() {
  const [results, setResults] = useState<TriageRow[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [exporting, setExporting] = useState<string | null>(null);
  const [page, setPage] = useState(0);
  const [shipped, setShipped] = useState<Record<string, string>>({});
  const [showJira, setShowJira] = useState(false);

  const handleShip = async (row: TriageRow) => {
    if (!row.prediction_id) return;
    const id = row.prediction_id;
    setExporting(`ship-${id}`);
    try {
      const response = await fetch("/api/ship", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ prediction_id: id }),
      });
      if (!response.ok) throw new Error("Could not mark as shipped");
      const data = await response.json();
      setShipped((prev) => ({ ...prev, [id]: data.shipped_date ?? "today" }));
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not mark as shipped");
    } finally {
      setExporting(null);
    }
  };

  const handleAnalyze = async () => {
    setIsLoading(true);
    setError(null);
    setResults([]);
    setPage(0);

    for (const ticket of pendingTickets) {
      try {
        const response = await fetch("/api/triage", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            company_id: ticket.company_id,
            workflow: ticket.workflow,
            feedback: ticket.feedback,
          }),
        });

        if (!response.ok) {
          throw new Error(`Request failed for ${ticket.company_id}`);
        }

        const raw = await response.text();

        if (!raw) {
          setResults((prev) => [
            ...prev,
            {
              feature: ticket.workflow,
              feedback: ticket.feedback,
              workflow: ticket.workflow,
              company_id: ticket.company_id,
              priority: "N/A",
              status: "NO DATA",
              rationale: `No triage data returned for ${ticket.company_id}. The backend workflow could not find a company profile to score this ticket against.`,
            },
          ]);
          continue;
        }

        const data: TriageApiResponse = JSON.parse(raw);
        setResults((prev) => [
          ...prev,
          {
            ...data,
            feature: data.feature || ticket.workflow,
            feedback: ticket.feedback,
            workflow: ticket.workflow,
            company_id: ticket.company_id,
          },
        ]);
      } catch (err) {
        setError(
          err instanceof Error ? err.message : "Failed to analyze tickets."
        );
      }
    }

    setIsLoading(false);
  };

  const handleDownloadCsv = () => {
    const header = COLUMNS.map((c) => escapeCsvValue(c.label)).join(",");
    const rows = results.map((row) =>
      COLUMNS.map((c) => escapeCsvValue(getCellValue(row, c.key))).join(",")
    );
    const csv = [header, ...rows].join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    downloadBlob(blob, "voc-triage-results.csv");
  };

  const handleCreatePptx = async () => {
    setExporting("pptx");
    try {
      const PptxGenJS = await loadPptxGenJS();
      const pptx = new PptxGenJS();

      const title = pptx.addSlide();
      title.addText("VoC Triage Engine", {
        x: 0.5,
        y: 1.8,
        w: 9,
        h: 1,
        fontSize: 36,
        bold: true,
        color: "111827",
      });
      title.addText("Stakeholder Review — Customer Feedback Triage", {
        x: 0.5,
        y: 2.7,
        w: 9,
        h: 0.6,
        fontSize: 18,
        color: "6B7280",
      });
      title.addText(new Date().toLocaleDateString(), {
        x: 0.5,
        y: 3.3,
        w: 9,
        h: 0.4,
        fontSize: 12,
        color: "9CA3AF",
      });

      for (const row of results) {
        const slide = pptx.addSlide();
        slide.addText(row.feature, {
          x: 0.4,
          y: 0.3,
          w: 9.2,
          h: 0.6,
          fontSize: 24,
          bold: true,
          color: "111827",
        });
        slide.addText(`${row.priority}  •  ${row.status}`, {
          x: 0.4,
          y: 0.9,
          w: 9.2,
          h: 0.4,
          fontSize: 14,
          bold: true,
          color: "DC2626",
        });
        slide.addText(
          [
            { text: "Issue: ", options: { bold: true } },
            { text: `${row.feedback}\n\n` },
            { text: "Rationale: ", options: { bold: true } },
            { text: `${row.rationale}\n\n` },
            { text: "Company: ", options: { bold: true } },
            {
              text: `${row.company_name ?? "—"} (${row.company_id})  |  MRR: ${
                row.mrr ?? "—"
              }  |  Tier: ${row.subscription_tier ?? "—"}  |  Health: ${
                row.csm_health_score ?? "—"
              }`,
            },
          ],
          {
            x: 0.4,
            y: 1.5,
            w: 9.2,
            h: 4.5,
            fontSize: 13,
            color: "374151",
            valign: "top",
          }
        );
      }

      await pptx.writeFile({ fileName: "voc-triage-stakeholder-review.pptx" });
    } finally {
      setExporting(null);
    }
  };

  const handleCreatePrd = async (row: TriageRow, index: number) => {
    setExporting(`prd-${index}`);
    try {
      const { Document, Packer, Paragraph, HeadingLevel, TextRun } =
        await import("docx");

      const doc = new Document({
        sections: [
          {
            children: [
              new Paragraph({
                text: `PRD: ${row.feature}`,
                heading: HeadingLevel.TITLE,
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: `Priority ${row.priority}  •  ${row.status}`,
                    bold: true,
                  }),
                ],
              }),
              new Paragraph({ text: "" }),

              new Paragraph({
                text: "Problem Statement",
                heading: HeadingLevel.HEADING_1,
              }),
              new Paragraph({ text: row.feedback }),
              new Paragraph({ text: "" }),

              new Paragraph({
                text: "Business Rationale",
                heading: HeadingLevel.HEADING_1,
              }),
              new Paragraph({ text: row.rationale }),
              new Paragraph({ text: "" }),

              new Paragraph({
                text: "Customer Context",
                heading: HeadingLevel.HEADING_1,
              }),
              new Paragraph({
                text: `Company: ${row.company_name ?? "—"} (${row.company_id})`,
              }),
              new Paragraph({ text: `User: ${row.user_name ?? "—"} (${row.user_id ?? "—"})` }),
              new Paragraph({ text: `MRR: ${row.mrr ?? "—"}` }),
              new Paragraph({
                text: `Subscription Tier: ${row.subscription_tier ?? "—"}`,
              }),
              new Paragraph({ text: `Renewal Date: ${row.renewal_date ?? "—"}` }),
              new Paragraph({
                text: `CSM Health Score: ${row.csm_health_score ?? "—"}`,
              }),
              new Paragraph({ text: "" }),

              new Paragraph({
                text: "Proposed Solution",
                heading: HeadingLevel.HEADING_1,
              }),
              new Paragraph({ text: "[ Product manager to fill in ]" }),
              new Paragraph({ text: "" }),

              new Paragraph({
                text: "Success Metrics",
                heading: HeadingLevel.HEADING_1,
              }),
              new Paragraph({ text: "[ Product manager to fill in ]" }),
              new Paragraph({ text: "" }),

              new Paragraph({
                text: "Open Questions",
                heading: HeadingLevel.HEADING_1,
              }),
              new Paragraph({ text: "[ Product manager to fill in ]" }),
            ],
          },
        ],
      });

      const blob = await Packer.toBlob(doc);
      const safeName = row.feature.replace(/[^a-z0-9]+/gi, "-").toLowerCase();
      downloadBlob(blob, `PRD-${safeName}.docx`);
    } finally {
      setExporting(null);
    }
  };

  const hasResults = results.length > 0;
  const learning = results.find((r) => r.calibration_factor !== undefined);
  const totalPages = Math.max(1, Math.ceil(results.length / ROWS_PER_PAGE));
  const currentPage = Math.min(page, totalPages - 1);
  const pagedResults = results.slice(
    currentPage * ROWS_PER_PAGE,
    currentPage * ROWS_PER_PAGE + ROWS_PER_PAGE
  );

  return (
    <main className="min-h-screen p-8">
      <div className="mx-auto max-w-[1600px]">
        <header className="mb-6 flex flex-wrap items-center justify-between gap-4 border-b border-gray-200 pb-6">
          <div className="flex items-center gap-3">
            <div className="flex h-10 w-10 items-center justify-center rounded-lg bg-gray-900 text-white">
              <ClipboardList size={20} />
            </div>
            <div>
              <h1 className="text-xl font-bold text-gray-900">
                NextBuild
              </h1>
            </div>
          </div>
        </header>

        <section className="mb-8 rounded-xl border border-gray-200 bg-white p-8 text-center">
          <h2 className="mb-4 inline-block rounded-lg bg-indigo-50 px-3 py-1 text-sm font-extrabold text-indigo-700 sm:text-base">
            Stop Triaging. Start Roadmapping.
          </h2>

          <p className="mx-auto max-w-3xl text-sm leading-relaxed text-gray-600 sm:text-base">
            Requirements hit PMs from every direction: customer tickets,
            leadership demands, market research, and competitor moves.
            Saying &quot;yes&quot; to everything creates massive feature
            bloat, but missing the right signal costs PMs critical revenue.
          </p>

          <p className="mx-auto mt-4 max-w-3xl text-sm leading-relaxed text-gray-600 sm:text-base">
            Right now, PMs waste over 15 hours a week manually
            cross-referencing these endless funnels with CRM Revenue and
            Data Analytics. <span className="font-bold text-gray-900">NextBuild</span>{" "}
            automates this. We ingest the noise, weigh it against your live
            business data, and isolate the P0 features that actually drive
            business impact, prevent high-value churn, and are worth your
            engineering time.
          </p>

          <div className="mx-auto mt-6 inline-flex w-fit flex-wrap items-center justify-center gap-x-6 gap-y-2 rounded-lg border bg-gray-50/50 px-5 py-3">
            <div className="text-center">
              <p className="whitespace-nowrap text-sm font-bold text-gray-900">
                4 sources → 1 score
              </p>
              <p className="mt-0.5 whitespace-nowrap text-[11px] font-medium text-gray-500">
                Tickets, calls, CRM revenue, product analytics
              </p>
            </div>
            <div className="hidden h-8 w-px bg-gray-200 sm:block" />
            <div className="text-center">
              <p className="whitespace-nowrap text-sm font-bold text-gray-900">
                ~12 sec per ticket
              </p>
              <p className="mt-0.5 whitespace-nowrap text-[11px] font-medium text-gray-500">
                vs. hours of manual cross-referencing
              </p>
            </div>
            <div className="hidden h-8 w-px bg-gray-200 sm:block" />
            <div className="text-center">
              <p className="whitespace-nowrap text-sm font-bold text-gray-900">
                100% auditable
              </p>
              <p className="mt-0.5 whitespace-nowrap text-[11px] font-medium text-gray-500">
                Every priority shows its Reach, Impact, Confidence, Effort
              </p>
            </div>
          </div>

          <p className="mt-8 text-xs font-semibold uppercase tracking-wide text-gray-500">
            How it works: a roadmap that learns from every fix
          </p>
          <div className="mx-auto mt-3 grid max-w-4xl grid-cols-1 gap-3 text-left sm:grid-cols-4">
            {[
              {
                n: "1",
                title: "Ingest",
                text: "Tickets, calls, CRM revenue and product analytics in one view.",
              },
              {
                n: "2",
                title: "Score",
                text: "AI classifies each request and RICE math ranks it P0 to P3.",
              },
              {
                n: "3",
                title: "Ship",
                text: "Draft the PRD, build the executive deck, mark the fix as shipped.",
              },
              {
                n: "4",
                title: "Learn",
                text: "Predicted vs actual impact calibrates the next roadmap.",
              },
            ].map((step) => (
              <div
                key={step.n}
                className={`rounded-lg border p-3 ${
                  step.n === "4"
                    ? "border-indigo-200 bg-indigo-50"
                    : "border-gray-200 bg-white"
                }`}
              >
                <div className="flex items-center gap-2">
                  <span
                    className={`flex h-5 w-5 items-center justify-center rounded-full text-[11px] font-bold text-white ${
                      step.n === "4" ? "bg-indigo-600" : "bg-gray-900"
                    }`}
                  >
                    {step.n}
                  </span>
                  <span className="text-sm font-bold text-gray-900">
                    {step.title}
                  </span>
                </div>
                <p className="mt-1.5 text-xs leading-relaxed text-gray-600">
                  {step.text}
                </p>
              </div>
            ))}
          </div>
        </section>

        <div className="mb-6 rounded-lg border border-gray-200 bg-white p-4">
          <p className="mb-3 text-xs font-semibold uppercase tracking-wide text-gray-500">
            Download Source Sheets (for cross-verification)
          </p>
          <div className="flex flex-wrap gap-3">
            {SOURCE_SHEETS.map((sheet) => (
              <a
                key={sheet.name}
                href={sheet.url}
                className="flex items-center gap-2 rounded-lg border border-gray-300 bg-white px-3.5 py-2 text-sm font-semibold text-gray-700 transition hover:bg-gray-50"
              >
                <Sheet size={16} />
                {sheet.name}
              </a>
            ))}
          </div>
        </div>

        <div className="mb-6 flex flex-col items-center gap-2">
          <button
            onClick={handleAnalyze}
            disabled={isLoading}
            className="flex items-center gap-2 rounded-lg bg-indigo-600 px-5 py-2.5 text-sm font-semibold text-white transition hover:bg-indigo-700 disabled:cursor-not-allowed disabled:opacity-60"
          >
            {isLoading ? (
              <Loader2 size={16} className="animate-spin" />
            ) : (
              <Rocket size={16} />
            )}
            {isLoading ? "Generating..." : "Generate Data-Backed Roadmap"}
          </button>
          <p className="whitespace-nowrap text-xs font-medium text-gray-500">
            Turn endless requirements into a revenue-proven sprint in 3
            seconds.
          </p>
        </div>

        {error && (
          <div className="mb-4 rounded-lg border border-red-200 bg-red-50 p-3 text-sm text-red-700">
            {error}
          </div>
        )}

        {learning && (
          <div className="mb-4 flex items-start gap-3 rounded-lg border border-indigo-200 bg-indigo-50 p-4 text-sm text-indigo-900">
            <TrendingUp size={18} className="mt-0.5 shrink-0" />
            <p>
              <span className="font-semibold">Learning from outcomes: </span>
              {(learning.outcomes_used ?? 0) > 0
                ? `Impact estimates are calibrated x${learning.calibration_factor} using ${learning.outcomes_used} shipped outcome(s), where the predicted impact was compared with the actual change in MRR and customer health.`
                : "No shipped outcomes yet. Impact estimates stay uncalibrated (x1) until the first fix ships and is measured."}
            </p>
          </div>
        )}

        <div className="overflow-x-auto rounded-lg border border-gray-200 bg-white shadow-sm">
          <table className="w-full text-left text-sm">
            <thead className="bg-gray-50 text-xs uppercase tracking-wide text-gray-500">
              <tr>
                {COLUMNS.map((col) => (
                  <th
                    key={col.key}
                    className="whitespace-nowrap p-4 font-semibold"
                  >
                    {col.label}
                  </th>
                ))}
                <th className="whitespace-nowrap p-4 font-semibold">
                  Action
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {results.length === 0 ? (
                <tr>
                  <td
                    colSpan={COLUMNS.length + 1}
                    className="p-10 text-center text-sm text-gray-400"
                  >
                    {isLoading
                      ? "Analyzing tickets, results will appear here..."
                      : "No tickets analyzed yet. Click \"Generate Data-Backed Roadmap\" to begin."}
                  </td>
                </tr>
              ) : (
                pagedResults.map((row, localIndex) => {
                  const index = currentPage * ROWS_PER_PAGE + localIndex;
                  const canCreatePrd = row.status.toUpperCase() === "TO BE PICKED";
                  return (
                    <tr key={index} className="hover:bg-gray-50">
                      {COLUMNS.map((col) => {
                        if (col.key === "priority") {
                          return (
                            <td key={col.key} className="p-4">
                              <PriorityBadge priority={row.priority} />
                            </td>
                          );
                        }
                        if (col.key === "status") {
                          return (
                            <td key={col.key} className="p-4">
                              <StatusBadge status={row.status} />
                            </td>
                          );
                        }
                        if (col.key === "rice") {
                          return (
                            <td key={col.key} className="p-4">
                              <div className="font-semibold text-gray-900">
                                {row.rice_score ?? "—"}
                              </div>
                              {row.rice_score !== undefined && (
                                <div className="mt-0.5 whitespace-nowrap text-xs text-gray-500">
                                  R {row.rice_reach} · I {row.rice_impact} · C{" "}
                                  {row.rice_confidence} · E {row.rice_effort}
                                </div>
                              )}
                            </td>
                          );
                        }
                        const isWide = col.key === "rationale" || col.key === "feedback";
                        return (
                          <td
                            key={col.key}
                            className={`p-4 ${
                              isWide
                                ? "min-w-[280px] text-gray-600"
                                : "whitespace-nowrap text-gray-700"
                            }`}
                          >
                            {getCellValue(row, col.key)}
                          </td>
                        );
                      })}
                      <td className="p-4">
                        <div className="flex flex-col items-start gap-2">
                        <button
                          onClick={() => handleCreatePrd(row, index)}
                          disabled={!canCreatePrd || exporting === `prd-${index}`}
                          title={
                            canCreatePrd
                              ? "Create a PRD draft for this ticket"
                              : "PRD creation is only available for tickets marked 'To Be Picked'"
                          }
                          className="flex items-center gap-1.5 whitespace-nowrap rounded-lg border border-indigo-200 bg-indigo-50 px-2.5 py-1.5 text-xs font-semibold text-indigo-700 transition hover:bg-indigo-100 disabled:cursor-not-allowed disabled:border-gray-300 disabled:bg-white disabled:text-gray-700 disabled:opacity-40"
                        >
                          {exporting === `prd-${index}` ? (
                            <Loader2 size={14} className="animate-spin" />
                          ) : (
                            <FileText size={14} />
                          )}
                          Create PRD
                        </button>
                        {row.prediction_id && shipped[row.prediction_id] ? (
                          <span className="flex items-center gap-1.5 whitespace-nowrap rounded-lg border border-green-200 bg-green-50 px-2.5 py-1.5 text-xs font-semibold text-green-700">
                            <CheckCircle2 size={14} />
                            Shipped {shipped[row.prediction_id]}
                          </span>
                        ) : (
                          <button
                            onClick={() => handleShip(row)}
                            disabled={
                              !canCreatePrd ||
                              !row.prediction_id ||
                              exporting === `ship-${row.prediction_id}`
                            }
                            title={
                              canCreatePrd
                                ? "Record that this fix has shipped so the system can learn from the outcome"
                                : "Only tickets marked 'To Be Picked' can be shipped"
                            }
                            className="flex items-center gap-1.5 whitespace-nowrap rounded-lg border border-green-200 bg-white px-2.5 py-1.5 text-xs font-semibold text-green-700 transition hover:bg-green-50 disabled:cursor-not-allowed disabled:border-gray-300 disabled:text-gray-700 disabled:opacity-40"
                          >
                            {exporting === `ship-${row.prediction_id}` ? (
                              <Loader2 size={14} className="animate-spin" />
                            ) : (
                              <PackageCheck size={14} />
                            )}
                            Mark as shipped
                          </button>
                        )}
                        <button
                          onClick={() => setShowJira(true)}
                          disabled={!canCreatePrd}
                          title="Create a Jira ticket and sync its status automatically"
                          className="flex items-center gap-1.5 whitespace-nowrap rounded-lg border border-blue-200 bg-white px-2.5 py-1.5 text-xs font-semibold text-blue-700 transition hover:bg-blue-50 disabled:cursor-not-allowed disabled:border-gray-300 disabled:text-gray-700 disabled:opacity-40"
                        >
                          <Ticket size={14} />
                          Sync with Jira
                          <span className="rounded bg-blue-100 px-1 py-0.5 text-[10px] font-bold uppercase text-blue-700">
                            Soon
                          </span>
                        </button>
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>

        {hasResults && (
          <div className="mt-3 flex items-center justify-between">
            <p className="text-xs text-gray-500">
              Page {currentPage + 1} of {totalPages}
            </p>
            <div className="flex items-center gap-2">
              <button
                onClick={() => setPage((p) => Math.max(0, p - 1))}
                disabled={currentPage === 0}
                className="rounded-lg border border-gray-300 bg-white px-3 py-1.5 text-xs font-semibold text-gray-700 transition hover:bg-gray-50 disabled:cursor-not-allowed disabled:opacity-50"
              >
                Previous
              </button>
              <button
                onClick={() =>
                  setPage((p) => Math.min(totalPages - 1, p + 1))
                }
                disabled={currentPage >= totalPages - 1}
                className="rounded-lg border border-gray-300 bg-white px-3 py-1.5 text-xs font-semibold text-gray-700 transition hover:bg-gray-50 disabled:cursor-not-allowed disabled:opacity-50"
              >
                Next
              </button>
            </div>
          </div>
        )}

        <div className="mt-6 flex flex-wrap items-center gap-3">
          <button
            onClick={handleDownloadCsv}
            disabled={!hasResults}
            className="flex items-center gap-2 rounded-lg border border-gray-300 bg-white px-3.5 py-2 text-sm font-semibold text-gray-700 transition hover:bg-gray-50 disabled:cursor-not-allowed disabled:opacity-50"
          >
            <Download size={16} />
            Download Roadmap
          </button>

          <button
            onClick={handleCreatePptx}
            disabled={!hasResults || exporting === "pptx"}
            className="flex items-center gap-2 rounded-lg border border-gray-300 bg-white px-3.5 py-2 text-sm font-semibold text-gray-700 transition hover:bg-gray-50 disabled:cursor-not-allowed disabled:opacity-50"
          >
            {exporting === "pptx" ? (
              <Loader2 size={16} className="animate-spin" />
            ) : (
              <Presentation size={16} />
            )}
            Generate Executive Deck
          </button>
        </div>
      </div>

      {showJira && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-gray-900/40 p-4"
          onClick={() => setShowJira(false)}
        >
          <div
            className="w-full max-w-lg rounded-xl bg-white p-6 shadow-xl"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-start justify-between gap-4">
              <div>
                <h3 className="text-lg font-bold text-gray-900">
                  Jira sync
                  <span className="ml-2 rounded bg-blue-100 px-1.5 py-0.5 align-middle text-[10px] font-bold uppercase text-blue-700">
                    Coming soon
                  </span>
                </h3>
                <p className="mt-1 text-sm text-gray-600">
                  Today the PM clicks Mark as shipped. With Jira connected, that
                  manual step disappears.
                </p>
              </div>
              <button
                onClick={() => setShowJira(false)}
                className="rounded p-1 text-gray-500 hover:bg-gray-100"
                aria-label="Close"
              >
                <X size={18} />
              </button>
            </div>

            <ol className="mt-4 space-y-3 text-sm text-gray-700">
              <li className="flex gap-3">
                <span className="flex h-5 w-5 shrink-0 items-center justify-center rounded-full bg-gray-900 text-[11px] font-bold text-white">
                  1
                </span>
                Create a Jira ticket from any To Be Picked row, with the PRD
                attached.
              </li>
              <li className="flex gap-3">
                <span className="flex h-5 w-5 shrink-0 items-center justify-center rounded-full bg-gray-900 text-[11px] font-bold text-white">
                  2
                </span>
                When the ticket moves to Done, NextBuild records the ship date
                automatically.
              </li>
              <li className="flex gap-3">
                <span className="flex h-5 w-5 shrink-0 items-center justify-center rounded-full bg-indigo-600 text-[11px] font-bold text-white">
                  3
                </span>
                The outcome check runs on a schedule, compares predicted and
                actual impact, and recalibrates the next roadmap.
              </li>
            </ol>

            <p className="mt-5 rounded-lg bg-gray-50 p-3 text-xs text-gray-600">
              Not connected in this demo. Use Mark as shipped to simulate the
              Done status.
            </p>
          </div>
        </div>
      )}
    </main>
  );
}
