/**
 * Interfaces.ts
 * -------------
 * Shared, framework-agnostic type definitions used across the control and
 * all of its services. Kept dependency-free so services stay easy to
 * unit-test in isolation.
 */

/**
 * Sentinel annotationid used when a "note" shown in the UI actually
 * represents a Dataverse File-column value rather than a real Annotation
 * record (single-file mode with documentFileFieldLogicalName set - see
 * index.ts loadFromDocumentField()). Download/Delete handlers branch on
 * this value to route to the file-column APIs instead of AnnotationService.
 */
export const FILE_FIELD_NOTE_ID = "__file-field-value__";

/**
 * A file that has been selected/dropped by the user but not yet persisted
 * to Dataverse. Lives entirely in browser memory (the File object cannot
 * be serialized into the bound column) until the parent record is saved.
 */
export interface IPendingFile {
    /** Client-generated unique id, used as a stable DOM/list key */
    id: string;
    /** The native browser File handle (binary content) */
    file: File;
    /** Original file name including extension */
    fileName: string;
    /** File size in bytes */
    fileSize: number;
    /** MIME type as reported by the browser (may be empty for some types) */
    mimeType: string;
    /** Data URL used to render a thumbnail/preview; undefined if not previewable */
    previewDataUrl?: string;
}

/**
 * An Annotation (Note) record already persisted in Dataverse and
 * associated with the current entity record.
 */
export interface IAnnotationRecord {
    annotationid: string;
    filename: string;
    mimetype: string;
    filesize: number;
    subject?: string;
    createdon?: string;
    /**
     * documentbody is intentionally excluded from the list query (large
     * base64 payloads would slow down the list). It is fetched on demand
     * only when the user clicks "Download".
     */
    documentbody?: string;
}

/** Outcome of validating one or more files against the configured rules */
export interface IValidationResult {
    isValid: boolean;
    errors: string[];
}

/** How the control's "trigger" (as opposed to the file list) is rendered. */
export type DisplayMode = "Dropzone" | "Button";

/** Strongly typed configuration derived from the manifest's input properties */
export interface IControlConfig {
    documentFileFieldLogicalName: string;
    allowMultipleFiles: boolean;
    allowedExtensions: string[];
    allowedMimeTypes: string[];
    maxFileSizeKB: number;
    maxFileCount: number;

    /**
     * Manual override for "which field is this control instance on" - see
     * index.ts `resolveSourceFieldTag()`. Optional; auto-detected from the
     * bound field's metadata in the vast majority of cases.
     */
    // currentFieldLogicalName: string;

    /** Dropzone (current large drag & drop card) or Button (compact Tailwind button). */
    displayMode: DisplayMode;
    /** Label shown on the trigger button before any file is staged/attached (Button mode only). */
    buttonIdleLabel: string;

    /** Hide this PCF control's own field on the form once upload/note-creation succeeds. */
    hideControlAfterUpload: boolean;
    /** Show the bound Dataverse File column's field on the form once upload succeeds. */
    showDocumentFieldAfterUpload: boolean;
    /**
     * Single-file mode with documentFileFieldLogicalName set only: after the
     * file is uploaded to the File column, delete the Annotation (Note) that
     * would otherwise also be created for it. When true, the Note is never
     * even created (avoids uploading the same file twice); when false, the
     * Note is kept alongside the File column value as an audit-trail copy
     * (previous/default behavior). Preview always reads from the File
     * column in this mode either way - see index.ts loadFromDocumentField().
     */
    deleteNoteAfterFileFieldUpload: boolean;
    /** Master switch for every Global-Xrm-dependent feature (hide/show/confirm dialogs/client URL). */
    // enableXrmIntegration: boolean;
    /** Show the "Download" action on already-attached files. */
    allowDownload: boolean;
    /** Show the "Delete"/"Remove" action on files (pending or attached). */
    allowDelete: boolean;
    /** Ask for confirmation (native Xrm dialog, falling back to window.confirm) before deleting an attached file. */
    confirmBeforeDelete: boolean;
    /** Render image/PDF/text thumbnails in file cards instead of always using the generic type icon. */
    showPreviewThumbnails: boolean;
    /** Show the size/MIME-type meta line under each file name. */
    showFileMeta: boolean;
    /** Fully read-only: no new uploads, no remove/delete actions - the list becomes purely informational. */
    readOnly: boolean;
    /** Optional accent color (hex) used for the dropzone/button/links/badges instead of the default blue. */
    accentColor: string;
    /** Optional accent color (hex) used for the Label dropzone/button/links/badges instead of the default blue. */
    labalColor: string;

    replacePCFWithSelectedFiles: boolean;
    allowPreview: boolean;
}

/** Minimal shape of the JSON persisted into the bound Multiline Text column */
export interface IPersistedState {
    /** Metadata only - never the binary content - so the field stays small */
    pendingFileMeta: {
        id: string;
        fileName: string;
        fileSize: number;
        mimeType: string;
    }[];
    /**
     * Guards against duplicate Note creation if updateView fires more than
     * once after save (e.g. platform re-renders trigger a second pass).
     */
    notesCreatedForSession: boolean;
    /**
     * Count of Notes already persisted in Dataverse for this record/field.
     * Kept in sync purely so the bound value stays non-empty (and thus
     * satisfies a "Business Required" flag on this column) whenever real
     * files exist, even though the binaries themselves never live here.
     */
    attachedFileCount?: number;
}

/** Result of an attempt to create Notes for a batch of pending files */
export interface IBatchUploadResult {
    succeeded: string[];
    failed: { fileName: string; error: string }[];
}