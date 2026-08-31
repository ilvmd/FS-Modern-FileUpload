import { IControlConfig, IPendingFile } from "../types/Interfaces";
import { ValidationService } from "./ValidationService";

/**
 * FileService
 * -----------
 * Owns the in-memory collection of files staged by the user - whether on a
 * Create form (record does not exist yet) or an existing record before an
 * additional save. Nothing in this service talks to Dataverse; see
 * AnnotationService for persistence once a record id is available.
 *
 * NOTE ON LIFETIME: browser File objects cannot be serialized into the
 * bound Multiline Text column (or anywhere else) - they only exist for the
 * lifetime of this control instance. If the control is destroyed and
 * re-initialized before the record is saved (e.g. the user navigates away
 * from the form tab), staged files are lost. The control persists file
 * *metadata* to the bound field precisely so it can detect and warn about
 * this situation - see index.ts `restoreStateFromBoundField`.
 */
export class FileService {
    /** Map keyed by client-generated id -> staged file wrapper */
    private pendingFiles: Map<string, IPendingFile> = new Map();

    /**
     * Validates and stages a set of files (from an <input type="file">
     * change event or a drop event). Files that fail validation are
     * reported in `errors` and are NOT added to the pending collection.
     */
    public async addFiles(uploadedCount: number, fileList: FileList | File[], config: IControlConfig): Promise<{ added: IPendingFile[]; errors: string[] }> {
        const incoming = Array.from(fileList);
        const errors: string[] = [];
        const added: IPendingFile[] = [];

        // Enforce max count / single-file constraints up front, before doing
        // any (comparatively expensive) per-file preview generation.
        const countCheck = ValidationService.validateFileCount(uploadedCount, this.pendingFiles.size, incoming.length, config);
        if (!countCheck.isValid) {
            return { added: [], errors: countCheck.errors };
        }

        for (const file of incoming) {
            const validation = ValidationService.validateFile(file, config);
            if (!validation.isValid) {
                errors.push(...validation.errors);
                continue;
            }

            const pending: IPendingFile = {
                id: FileService.generateId(),
                file,
                fileName: file.name,
                fileSize: file.size,
                mimeType: file.type || "application/octet-stream"
            };

            // Best-effort thumbnail/preview - never blocks staging the file.
            try {
                pending.previewDataUrl = await FileService.generatePreview(file);
            } catch {
                pending.previewDataUrl = undefined;
            }

            this.pendingFiles.set(pending.id, pending);
            added.push(pending);
        }

        return { added, errors };
    }

    /** Removes a single staged file by its client-generated id */
    public removeFile(id: string): void {
        this.pendingFiles.delete(id);
    }

    /** Clears all staged files (called after they have been persisted as Notes) */
    public clear(): void {
        this.pendingFiles.clear();
    }

    /** Returns all currently staged files, insertion order preserved */
    public getPendingFiles(): IPendingFile[] {
        return Array.from(this.pendingFiles.values());
    }

    public getCount(): number {
        return this.pendingFiles.size;
    }

    /**
     * Converts a File to a base64 string WITHOUT the "data:mime;base64,"
     * prefix, as required by the Dataverse Annotation.documentbody field.
     */
    public static fileToBase64(file: File): Promise<string> {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => {
                const result = reader.result as string;
                const commaIndex = result.indexOf(",");
                resolve(commaIndex >= 0 ? result.substring(commaIndex + 1) : result);
            };
            reader.onerror = () => reject(reader.error ?? new Error("Failed to read file."));
            reader.readAsDataURL(file);
        });
    }

    /**
     * Generates a preview data URL for supported types:
     *   - images: the image itself (the UI layer renders it scaled down)
     *   - PDFs: rendered by the UI layer inside an <embed>/<iframe>
     *   - text: rendered by the UI layer as a small text snippet
     * Returns undefined for any other type (a generic file icon is used).
     */
    private static generatePreview(file: File): Promise<string | undefined> {
        const isPreviewable = file.type.startsWith("image/") || file.type === "application/pdf" || file.type.startsWith("text/");

        if (!isPreviewable) {
            return Promise.resolve(undefined);
        }

        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result as string);
            reader.onerror = () => reject(reader.error ?? new Error("Failed to generate preview."));
            reader.readAsDataURL(file);
        });
    }

    /** Generates a reasonably unique client-side id (crypto.randomUUID with a safe fallback) */
    public static generateId(): string {
        const cryptoObj = (window as unknown as { crypto?: Crypto }).crypto;
        if (cryptoObj && typeof cryptoObj.randomUUID === "function") {
            return cryptoObj.randomUUID();
        }
        // Fallback RFC4122-v4-ish generator for older embedded WebViews
        return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
            const r = (Math.random() * 16) | 0;
            const v = c === "x" ? r : (r & 0x3) | 0x8;
            return v.toString(16);
        });
    }
}
