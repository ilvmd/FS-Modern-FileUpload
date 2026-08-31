import { IControlConfig, IValidationResult } from "../types/Interfaces";

/**
 * ValidationService
 * -----------------
 * Pure, stateless validation logic for files selected/dropped by the user.
 * Contains no DOM access and no Dataverse access, so it is trivially
 * unit-testable in isolation from the rest of the control.
 */
export class ValidationService {
    /**
     * Validates a single File against the supplied configuration.
     * Checks, in order: extension allow-list, MIME type allow-list, max size.
     * All applicable rules are evaluated (not short-circuited) so the user
     * sees every problem with a file at once.
     */
    public static validateFile(file: File, config: IControlConfig): IValidationResult {
        const errors: string[] = [];

        // --- Extension check ---------------------------------------------
        const extension = ValidationService.getExtension(file.name).toLowerCase();
        if (config.allowedExtensions.length > 0 && !config.allowedExtensions.includes(extension)) {
            errors.push(
                `"${file.name}": file type "${extension || "unknown"}" is not allowed. Allowed types: ${config.allowedExtensions.join(", ")}`
            );
        }

        // --- MIME type check ------------------------------------------------
        // Some browsers/OS combinations report an empty file.type for certain
        // extensions (e.g. .csv on some platforms) - only enforce the check
        // when the browser actually reported a MIME type, to avoid false
        // rejections; the extension check above still applies regardless.
        if (config.allowedMimeTypes.length > 0 && file.type && !config.allowedMimeTypes.includes(file.type)) {
            errors.push(`"${file.name}": MIME type "${file.type}" is not allowed.`);
        }

        // --- Size check -------------------------------------------------
        const maxBytes = config.maxFileSizeKB * 1024;
        if (maxBytes > 0 && file.size > maxBytes) {
            errors.push(
                `"${file.name}" exceeds the maximum allowed size of ${ValidationService.formatBytes(maxBytes)} ` +
                    `(selected file is ${ValidationService.formatBytes(file.size)}).`
            );
        }

        return { isValid: errors.length === 0, errors };
    }

    /**
     * Validates the total pending file count (already-staged + incoming)
     * against the configured maximum, and against the single-file rule
     * when multiple files are disabled.
     */
    public static validateFileCount(uploadedCount: number, currentCount: number, incomingCount: number, config: IControlConfig): IValidationResult {
        const errors: string[] = [];

        if (!config.allowMultipleFiles && (uploadedCount + currentCount + incomingCount) > 1) {
            errors.push("Only a single file can be attached to this field.");
            return { isValid: false, errors };
        }

        if (config.maxFileCount > 0 && (uploadedCount + currentCount + incomingCount) > config.maxFileCount) {
            errors.push(
                `You can attach a maximum of ${config.maxFileCount} file(s). ` +
                    `You currently have ${currentCount} staged/attached.`
            );
        }

        return { isValid: errors.length === 0, errors };
    }

    /** Extracts the extension including the leading dot, e.g. "a.b.pdf" -> ".pdf" */
    public static getExtension(fileName: string): string {
        const idx = fileName.lastIndexOf(".");
        return idx >= 0 ? fileName.substring(idx) : "";
    }

    /** Formats a byte count into a human-readable string (B/KB/MB/GB) */
    public static formatBytes(bytes: number): string {
        if (!bytes || bytes <= 0) {
            return "0 B";
        }
        const units = ["B", "KB", "MB", "GB"];
        const exponent = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
        const value = bytes / Math.pow(1024, exponent);
        return `${value.toFixed(exponent === 0 ? 0 : 1)} ${units[exponent]}`;
    }

    /**
     * Parses a comma-separated configuration string (extensions or MIME
     * types) into a clean, lower-cased array with empty entries removed.
     */
    public static parseCsvList(value: string | null | undefined): string[] {
        if (!value) {
            return [];
        }
        return value
            .split(",")
            .map((v) => v.trim().toLowerCase())
            .filter((v) => v.length > 0);
    }
}
