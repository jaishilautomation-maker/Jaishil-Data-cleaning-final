import { createFileRoute } from "@tanstack/react-router";
import { useState } from "react";
import { Database, Sparkles, CheckCircle2, Download, Loader2, ShieldCheck } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group";
import { Label } from "@/components/ui/label";
import {
  Accordion,
  AccordionContent,
  AccordionItem,
  AccordionTrigger,
} from "@/components/ui/accordion";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Badge } from "@/components/ui/badge";
import { toast } from "sonner";
import { FileDropzone } from "@/components/FileDropzone";
import {
  cleanFile,
  downloadCleanedFile,
  PROFILE_LABELS,
  type CleanResult,
  type Profile,
} from "@/lib/cleansing";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Jaishil Data Cleansing Portal" },
      {
        name: "description",
        content:
          "Secure enterprise portal for cleansing stock inventory and debtors reports for Jaishil Sulphur & Chemical Industries.",
      },
      { property: "og:title", content: "Jaishil Data Cleansing Portal" },
      {
        property: "og:description",
        content: "Enterprise data cleansing for stock and debtors reports.",
      },
    ],
  }),
  component: Index,
});

function Index() {
  const [file, setFile] = useState<File | null>(null);
  const [profile, setProfile] = useState<Profile>("stock");
  const [processing, setProcessing] = useState(false);
  const [result, setResult] = useState<CleanResult | null>(null);

  const handleProcess = async () => {
    if (!file) {
      toast.error("Please upload a file first");
      return;
    }
    setProcessing(true);
    setResult(null);
    try {
      const res = await cleanFile(file, profile);
      setResult(res);
      toast.success(`Processed — ${res.removed.length} rows removed`);
    } catch (e) {
      console.error(e);
      toast.error("Failed to process file. Please check the format.");
    } finally {
      setProcessing(false);
    }
  };

  const handleDownload = () => {
    if (result && file) downloadCleanedFile(result, file.name);
  };

  const reset = () => {
    setFile(null);
    setResult(null);
  };

  return (
    <div className="min-h-screen bg-background">
      {/* Header */}
      <header className="border-b border-border bg-surface">
        <div className="mx-auto flex max-w-6xl items-center justify-between px-4 py-4 sm:px-6">
          <div className="flex items-center gap-3">
            <div className="flex h-10 w-10 items-center justify-center rounded-md bg-primary text-primary-foreground">
              <Database className="h-5 w-5" />
            </div>
            <div>
              <h1 className="text-base font-semibold text-foreground sm:text-lg">
                Jaishil Data Cleansing Portal
              </h1>
              <p className="text-xs text-muted-foreground">
                Sulphur &amp; Chemical Industries
              </p>
            </div>
          </div>
          <Badge variant="secondary" className="hidden gap-1.5 sm:flex">
            <ShieldCheck className="h-3.5 w-3.5" /> Secure Workspace
          </Badge>
        </div>
      </header>

      <main className="mx-auto max-w-6xl px-4 py-8 sm:px-6 sm:py-12">
        <div className="mb-8">
          <h2 className="text-2xl font-semibold tracking-tight text-foreground sm:text-3xl">
            Clean your reports in seconds
          </h2>
          <p className="mt-2 text-sm text-muted-foreground sm:text-base">
            Upload an Excel report, select the data profile, and download a cleaned file with a full change log.
          </p>
        </div>

        <div className="grid gap-6 lg:grid-cols-5">
          {/* Left: workflow */}
          <div className="space-y-6 lg:col-span-3">
            <Card className="shadow-[var(--shadow-card)]">
              <CardHeader>
                <CardTitle className="flex items-center gap-2 text-base">
                  <span className="flex h-6 w-6 items-center justify-center rounded-full bg-primary text-xs font-semibold text-primary-foreground">
                    1
                  </span>
                  Select data profile
                </CardTitle>
                <CardDescription>Choose what kind of report you're cleaning.</CardDescription>
              </CardHeader>
              <CardContent>
                <RadioGroup
                  value={profile}
                  onValueChange={(v) => setProfile(v as Profile)}
                  className="grid gap-3 sm:grid-cols-3"
                >
                  {(Object.keys(PROFILE_LABELS) as Profile[]).map((p) => (
                    <label
                      key={p}
                      htmlFor={p}
                      className={`flex cursor-pointer flex-col rounded-lg border p-4 transition-colors ${
                        profile === p
                          ? "border-primary bg-accent"
                          : "border-border bg-surface hover:border-primary/40"
                      }`}
                    >
                      <div className="flex items-center gap-2">
                        <RadioGroupItem value={p} id={p} />
                        <span className="text-sm font-semibold text-foreground">
                          {PROFILE_LABELS[p].title}
                        </span>
                      </div>
                      <span className="mt-2 text-xs text-muted-foreground">
                        {PROFILE_LABELS[p].desc}
                      </span>
                    </label>
                  ))}
                </RadioGroup>
              </CardContent>
            </Card>

            <Card className="shadow-[var(--shadow-card)]">
              <CardHeader>
                <CardTitle className="flex items-center gap-2 text-base">
                  <span className="flex h-6 w-6 items-center justify-center rounded-full bg-primary text-xs font-semibold text-primary-foreground">
                    2
                  </span>
                  Upload your file
                </CardTitle>
                <CardDescription>Excel files only (.xlsx, .xls).</CardDescription>
              </CardHeader>
              <CardContent>
                <FileDropzone file={file} onFile={(f) => { setFile(f); setResult(null); }} />
              </CardContent>
            </Card>

            <Button
              size="lg"
              className="w-full bg-primary text-primary-foreground hover:bg-[var(--color-primary-hover)]"
              disabled={!file || processing}
              onClick={handleProcess}
            >
              {processing ? (
                <><Loader2 className="mr-2 h-4 w-4 animate-spin" /> Processing…</>
              ) : (
                <><Sparkles className="mr-2 h-4 w-4" /> Process &amp; Clean Data</>
              )}
            </Button>
          </div>

          {/* Right: result */}
          <div className="lg:col-span-2">
            {result ? (
              <Card className="border-success/30 shadow-[var(--shadow-elevated)]">
                <CardHeader>
                  <div className="mb-2 flex h-10 w-10 items-center justify-center rounded-full bg-success/15 text-success">
                    <CheckCircle2 className="h-5 w-5" />
                  </div>
                  <CardTitle className="text-base">Cleansing complete</CardTitle>
                  <CardDescription>Your file is ready to download.</CardDescription>
                </CardHeader>
                <CardContent className="space-y-4">
                  <div className="grid grid-cols-3 gap-2 rounded-lg border border-border bg-surface-elevated p-3 text-center">
                    <Stat label="Original" value={result.originalCount} />
                    <Stat label="Removed" value={result.removed.length} tone="destructive" />
                    <Stat label="Cleaned" value={result.cleanedRows.length} tone="success" />
                  </div>

                  <Button
                    className="w-full bg-primary text-primary-foreground hover:bg-[var(--color-primary-hover)]"
                    onClick={handleDownload}
                  >
                    <Download className="mr-2 h-4 w-4" /> Download Cleaned File
                  </Button>

                  {result.removed.length > 0 && (
                    <Accordion type="single" collapsible>
                      <AccordionItem value="log" className="border-border">
                        <AccordionTrigger className="text-sm">
                          View change log ({result.removed.length} rows)
                        </AccordionTrigger>
                        <AccordionContent>
                          <div className="max-h-80 overflow-auto rounded-md border border-border">
                            <Table>
                              <TableHeader>
                                <TableRow>
                                  <TableHead className="w-16">Row</TableHead>
                                  <TableHead>Reason</TableHead>
                                </TableRow>
                              </TableHeader>
                              <TableBody>
                                {result.removed.slice(0, 200).map((r, i) => (
                                  <TableRow key={i}>
                                    <TableCell className="font-mono text-xs">{r.rowNumber}</TableCell>
                                    <TableCell className="text-xs">{r.reason}</TableCell>
                                  </TableRow>
                                ))}
                              </TableBody>
                            </Table>
                            {result.removed.length > 200 && (
                              <p className="border-t border-border p-2 text-center text-xs text-muted-foreground">
                                + {result.removed.length - 200} more in downloaded file
                              </p>
                            )}
                          </div>
                        </AccordionContent>
                      </AccordionItem>
                    </Accordion>
                  )}

                  <Button variant="outline" className="w-full" onClick={reset}>
                    Process another file
                  </Button>
                </CardContent>
              </Card>
            ) : (
              <Card className="border-dashed bg-surface-elevated">
                <CardContent className="flex flex-col items-center justify-center py-12 text-center">
                  <div className="mb-3 flex h-12 w-12 items-center justify-center rounded-full bg-muted text-muted-foreground">
                    <Sparkles className="h-5 w-5" />
                  </div>
                  <p className="text-sm font-medium text-foreground">Results appear here</p>
                  <p className="mt-1 text-xs text-muted-foreground">
                    Upload a file and click Process to see your cleaned report.
                  </p>
                </CardContent>
              </Card>
            )}
          </div>
        </div>

        <footer className="mt-12 border-t border-border pt-6 text-center text-xs text-muted-foreground">
          © {new Date().getFullYear()} Jaishil Sulphur &amp; Chemical Industries · Internal Use Only
        </footer>
      </main>
    </div>
  );
}

function Stat({ label, value, tone }: { label: string; value: number; tone?: "destructive" | "success" }) {
  const color =
    tone === "destructive"
      ? "text-destructive"
      : tone === "success"
      ? "text-success"
      : "text-foreground";
  return (
    <div>
      <p className={`text-xl font-semibold tabular-nums ${color}`}>{value.toLocaleString()}</p>
      <p className="text-[10px] uppercase tracking-wide text-muted-foreground">{label}</p>
    </div>
  );
}
