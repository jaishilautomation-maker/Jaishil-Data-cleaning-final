import { useDropzone } from "react-dropzone";
import { Upload, FileSpreadsheet, X } from "lucide-react";
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";

interface Props {
  file: File | null;
  onFile: (file: File | null) => void;
}

export function FileDropzone({ file, onFile }: Props) {
  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    accept: {
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
      "application/vnd.ms-excel": [".xls"],
    },
    multiple: false,
    onDrop: (files) => files[0] && onFile(files[0]),
  });

  if (file) {
    return (
      <div className="flex items-center justify-between rounded-lg border border-border bg-surface p-4 shadow-[var(--shadow-card)]">
        <div className="flex items-center gap-3">
          <div className="flex h-10 w-10 items-center justify-center rounded-md bg-accent text-primary">
            <FileSpreadsheet className="h-5 w-5" />
          </div>
          <div>
            <p className="text-sm font-medium text-foreground">{file.name}</p>
            <p className="text-xs text-muted-foreground">
              {(file.size / 1024).toFixed(1)} KB
            </p>
          </div>
        </div>
        <Button variant="ghost" size="icon" onClick={() => onFile(null)} aria-label="Remove file">
          <X className="h-4 w-4" />
        </Button>
      </div>
    );
  }

  return (
    <div
      {...getRootProps()}
      className={cn(
        "flex cursor-pointer flex-col items-center justify-center rounded-lg border-2 border-dashed bg-surface px-6 py-12 text-center transition-colors",
        isDragActive
          ? "border-primary bg-accent"
          : "border-border hover:border-primary/50 hover:bg-surface-elevated",
      )}
    >
      <input {...getInputProps()} />
      <div className="mb-3 flex h-12 w-12 items-center justify-center rounded-full bg-accent text-primary">
        <Upload className="h-6 w-6" />
      </div>
      <p className="text-sm font-medium text-foreground">
        {isDragActive ? "Drop your file here" : "Drag & drop Excel file here"}
      </p>
      <p className="mt-1 text-xs text-muted-foreground">
        or click to browse · .xlsx, .xls
      </p>
    </div>
  );
}
