/*
 * "Save a copy first" backup.
 *
 * The original VBA wrote `<name>_original.<ext>` next to the file. The Office.js sandbox
 * can't write to arbitrary disk paths, so instead we read the current document bytes via
 * getFileAsync and hand the user a download of `<name>_original.docx`. This is the
 * closest available equivalent and pairs with single-Ctrl+Z undo as the safety net.
 */

/* global Office, Blob, BlobPart, URL, document */

/** Best-effort base file name (without extension) from the document URL. */
export function documentBaseName(): string {
  return baseNameFromUrl();
}

/** Best-effort full file name (with extension) from the document URL. */
export function documentFileName(): string {
  try {
    const url = (Office.context.document as unknown as { url?: string }).url || "";
    const raw = decodeURIComponent(url.split(/[\\/]/).pop() || "");
    if (raw) return raw;
  } catch {
    // ignore
  }
  return "specification.docx";
}

function baseNameFromUrl(): string {
  try {
    const url = (Office.context.document as unknown as { url?: string }).url || "";
    const raw = decodeURIComponent(url.split(/[\\/]/).pop() || "");
    const dot = raw.lastIndexOf(".");
    const base = dot > 0 ? raw.slice(0, dot) : raw;
    if (base) return base;
  } catch {
    // ignore
  }
  return "specification";
}

/** True if the document has been saved to a location (has a URL/path). */
export function documentIsSaved(): boolean {
  try {
    return Boolean((Office.context.document as unknown as { url?: string }).url);
  } catch {
    return false;
  }
}

/**
 * Read the whole document and trigger a download of a `_original` copy.
 * Resolves once the download has started; rejects on failure.
 */
export function downloadBackup(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(
      Office.FileType.Compressed,
      { sliceSize: 65536 },
      (fileResult) => {
        if (fileResult.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(fileResult.error?.message || "Could not read the document."));
          return;
        }

        const file = fileResult.value;
        const sliceCount = file.sliceCount;
        const slices: Uint8Array[] = new Array(sliceCount);
        let received = 0;

        const finish = () => {
          file.closeAsync(() => {
            const blob = new Blob(slices as BlobPart[], {
              type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            });
            const url = URL.createObjectURL(blob);
            const fileName = baseNameFromUrl() + "_original.docx";
            const a = document.createElement("a");
            a.href = url;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            resolve(fileName);
          });
        };

        const getSlice = (index: number) => {
          file.getSliceAsync(index, (sliceResult) => {
            if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
              file.closeAsync(() => {});
              reject(
                new Error(sliceResult.error?.message || "Could not read a part of the document.")
              );
              return;
            }
            const data = sliceResult.value.data as number[] | Uint8Array;
            slices[sliceResult.value.index] =
              data instanceof Uint8Array ? data : new Uint8Array(data);
            received++;
            if (received === sliceCount) {
              finish();
            } else {
              getSlice(index + 1);
            }
          });
        };

        getSlice(0);
      }
    );
  });
}
