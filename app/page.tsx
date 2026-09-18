"use client";

import { useState } from "react";
import {
  Sparkles,
  Loader2,
  ClipboardList,
  Download,
  Presentation,
  FileText,
  Sheet,
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
    company_id: "comp_901",
    workflow: "report_data_export",
    feedback: "NetSuite sync timed out with error 504.",
  },
  {
    company_id: "comp_101",
    workflow: "receipt_capture_ocr",
    feedback: "Change the upload button to pink.",
  },
  {
    company_id: "comp_808",
    workflow: "card_transaction_auth",
    feedback: "Fuel cards are declining at the terminal.",
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
}

interface TriageRow extends TriageApiResponse {
  feature: string;
  feedback: string;
  workflow: string;
  company_id: string;
}

const WEBHOOK_URL =
  "https://vasaviprasannag.app.n8n.cloud/webhook/09dff8cf-b5ec-4c68-bd4d-79b06023f775";

type ColumnKey =
  | "feedback"
  | "core_workflow"
  | "priority"
  | "status"
  | "rationale"
  | "company_name"
  | "company_id"
  | "user_name"
  | "user_id"
  | "mrr"
  | "subscription_tier"
  | "renewal_date"
  | "csm_health_score";

const COLUMNS: { key: ColumnKey; label: string }[] = [
  { key: "feedback", label: "User Issue / Complaint / Transcript Snippet" },
  { key: "core_workflow", label: "Core Workflow Tracked / Analytics" },
  { key: "priority", label: "Priority" },
  { key: "status", label: "Status" },
  { key: "rationale", label: "Rationale" },
  { key: "company_name", label: "Company Name" },
  { key: "company_id", label: "Company ID" },
  { key: "user_name", label: "User Name" },
  { key: "user_id", label: "User ID" },
  { key: "mrr", label: "MRR" },
  { key: "subscription_tier", label: "Subscription Tier" },
  { key: "renewal_date", label: "Renewal Date" },
  { key: "csm_health_score", label: "CSM Health Score" },
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
    case "rationale":
      return row.rationale;
    case "company_name":
      return row.company_name ?? "—";
    case "company_id":
      return row.company_id;
    case "user_name":
      return row.user_name ?? "—";
    case "user_id":
      return row.user_id ?? "—";
    case "mrr":
      return row.mrr !== undefined && row.mrr !== null ? String(row.mrr) : "—";
    case "subscription_tier":
      return row.subscription_tier ?? "—";
    case "renewal_date":
      return row.renewal_date ?? "—";
    case "csm_health_score":
      return row.csm_health_score !== undefined && row.csm_health_score !== null
        ? String(row.csm_health_score)
        : "—";
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
      script.src = "/vendor/pptxgen.bundle.js";
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

export default function Home() {
  const [results, setResults] = useState<TriageRow[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [exporting, setExporting] = useState<string | null>(null);

  const handleAnalyze = async () => {
    setIsLoading(true);
    setError(null);
    setResults([]);

    for (const ticket of pendingTickets) {
      try {
        const response = await fetch(WEBHOOK_URL, {
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
                VoC Triage Engine
              </h1>
            </div>
          </div>

          <button
            onClick={handleAnalyze}
            disabled={isLoading}
            className="flex items-center gap-2 rounded-lg bg-gray-900 px-4 py-2.5 text-sm font-semibold text-white transition hover:bg-gray-700 disabled:cursor-not-allowed disabled:opacity-60"
          >
            {isLoading ? (
              <Loader2 size={16} className="animate-spin" />
            ) : (
              <Sparkles size={16} />
            )}
            {isLoading ? "Analyzing..." : "Analyze Pending Tickets"}
          </button>
        </header>

        <div className="mb-4 flex flex-wrap items-center gap-3">
          <button
            onClick={handleDownloadCsv}
            disabled={!hasResults}
            className="flex items-center gap-2 rounded-lg border border-gray-300 bg-white px-3.5 py-2 text-sm font-semibold text-gray-700 transition hover:bg-gray-50 disabled:cursor-not-allowed disabled:opacity-50"
          >
            <Download size={16} />
            Download as CSV
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
            Create PowerPoint Deck for{" "}
            <span className="font-bold text-amber-600">Stakeholder Review</span>
          </button>
        </div>

        <div className="mb-6 rounded-lg border border-gray-200 bg-white p-4">
          <p className="mb-3 text-xs font-semibold uppercase tracking-wide text-gray-500">
            Download Source Sheets (for cross-verification)
          </p>
          <div className="flex flex-wrap gap-3">
            {SOURCE_SHEETS.map((sheet) => (
              <a
                key={sheet.name}
                href={sheet.url}
                target="_blank"
                rel="noopener noreferrer"
                className="flex items-center gap-2 rounded-lg border border-gray-300 bg-white px-3.5 py-2 text-sm font-semibold text-gray-700 transition hover:bg-gray-50"
              >
                <Sheet size={16} />
                {sheet.name}
              </a>
            ))}
          </div>
        </div>

        {error && (
          <div className="mb-4 rounded-lg border border-red-200 bg-red-50 p-3 text-sm text-red-700">
            {error}
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
                  Actions
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
                      : "No tickets analyzed yet. Click \"Analyze Pending Tickets\" to begin."}
                  </td>
                </tr>
              ) : (
                results.map((row, index) => {
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
                        const isWide =
                          col.key === "feedback" || col.key === "rationale";
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
                        <button
                          onClick={() => handleCreatePrd(row, index)}
                          disabled={!canCreatePrd || exporting === `prd-${index}`}
                          title={
                            canCreatePrd
                              ? "Create a PRD draft for this ticket"
                              : "PRD creation is only available for tickets marked 'To Be Picked'"
                          }
                          className="flex items-center gap-1.5 whitespace-nowrap rounded-lg border border-gray-300 bg-white px-2.5 py-1.5 text-xs font-semibold text-gray-700 transition hover:bg-gray-50 disabled:cursor-not-allowed disabled:opacity-40"
                        >
                          {exporting === `prd-${index}` ? (
                            <Loader2 size={14} className="animate-spin" />
                          ) : (
                            <FileText size={14} />
                          )}
                          Create PRD
                        </button>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>
    </main>
  );
}
