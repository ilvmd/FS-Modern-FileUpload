import { DisplayMode, IAnnotationRecord, IControlConfig, IPendingFile } from "../types/Interfaces";
import { ValidationService } from "./ValidationService";
import * as XLSX from "xlsx";


/** Callback contract the UI layer needs from its controller (index.ts). */
export interface IUICallbacks {
    onFilesSelected: (files: FileList) => void;
    onRemovePendingFile: (id: string) => void;
    onDownloadNote: (annotationId: string, fileName: string, mimeType: string) => void;
    onPreviewNote: (annotationId: string, fileName: string, mimeType: string) => void;
    onDeleteNote: (annotationId: string) => void;
    onDeleteFromFileField: () => void;
}

/**
 * UIService
 * ---------
 * Owns all DOM construction and updates for the control. Contains no
 * validation logic and no Dataverse access - it is purely presentational,
 * driven entirely by state handed to it by index.ts.
 *
 * Visual styling is Tailwind CSS v3 utility classes (compiled ahead of
 * time - see tailwind.config.js / css/tailwind.src.css / `npm run
 * build:css`), every one written as a complete literal string so
 * Tailwind's build-time content scanner can find it. All generated
 * utilities are namespaced behind a `tw-` prefix and Tailwind's global
 * "preflight" reset is disabled, so this stylesheet can safely live
 * alongside the rest of a Dataverse form without affecting sibling
 * controls (see tailwind.config.js for details).
 *
 * Structural/query hooks (for JS, not for styling) use plain `data-fuc-*`
 * attributes so styling and behavior concerns stay cleanly separated.
 */
export class UIService {
    private readonly container: HTMLDivElement;
    private readonly callbacks: IUICallbacks;
    private config: IControlConfig;

    private rootEl!: HTMLDivElement;

    private previewBlobUrl?: string;
    private previewKeydownHandler?: (
        event: KeyboardEvent
    ) => void;

    /** Whichever trigger (dropzone card or button) is currently mounted. */
    private triggerEl!: HTMLElement;
    private triggerLabelEl: HTMLElement | null = null;
    private triggerSubtextEl: HTMLElement | null = null;
    private currentDisplayMode: DisplayMode | undefined;

    private fileInputEl!: HTMLInputElement;
    private errorBannerEl!: HTMLDivElement;
    private pendingSectionEl!: HTMLDivElement;
    private pendingListEl!: HTMLDivElement;
    private notesSectionEl!: HTMLDivElement;
    private notesListEl!: HTMLDivElement;
    private loadingOverlayEl!: HTMLDivElement;
    private loadingTextEl!: HTMLDivElement;

    /** Lives for the whole control instance - only torn down in destroy(). */
    private readonly rootAbortController = new AbortController();
    /** Re-created every time the trigger element is (re)built (mode switch, config refresh). */
    private triggerAbortController = new AbortController();

    private lastPendingFiles: IPendingFile[] = [];
    private lastNotes: IAnnotationRecord[] = [];

    public constructor(container: HTMLDivElement, callbacks: IUICallbacks, config: IControlConfig) {
        this.container = container;
        this.callbacks = callbacks;
        this.config = config;
        this.buildBaseLayout();
    }

    /** Applies updated configuration (e.g. if manifest inputs change) to already-built DOM. */
    public updateConfig(config: IControlConfig): void {
        this.config = config;
        this.applyAccentColor();
        this.applyLabelColor();

        this.fileInputEl.multiple = config.allowMultipleFiles;
        this.fileInputEl.accept = config.allowedExtensions.join(",");

        if (config.displayMode !== this.currentDisplayMode) {
            this.rebuildTrigger();
        } else if (this.triggerSubtextEl) {
            this.triggerSubtextEl.textContent = this.buildHintText();
        }

        this.refreshTriggerState();
    }

    /** Removes every event listener this service attached (call from index.ts destroy()). */
    public destroy(): void {
        this.triggerAbortController.abort();
        this.rootAbortController.abort();
    }

    // =======================================================================
    // Base layout (built once)
    // =======================================================================

    private buildBaseLayout(): void {
        this.rootEl = this.el("div", "fuc-root tw-relative tw-flex tw-w-full tw-flex-col tw-gap-3 tw-font-sans tw-text-sm tw-text-slate-800");
        this.applyAccentColor();
        this.applyLabelColor();

        this.fileInputEl = this.buildHiddenFileInput();
        this.triggerEl = this.buildTrigger();

        this.errorBannerEl = this.el("div", "tw-hidden");

        this.pendingSectionEl = this.el("div", "tw-hidden tw-flex tw-flex-col tw-gap-1.5");
        const pendingTitle = this.el("div", "tw-text-[11px] tw-font-semibold tw-uppercase tw-tracking-wide tw-text-slate-500");
        pendingTitle.textContent = "Files to upload";
        this.pendingListEl = this.el("div", "tw-flex tw-flex-col tw-gap-1.5");
        this.pendingSectionEl.appendChild(pendingTitle);
        this.pendingSectionEl.appendChild(this.pendingListEl);

        this.pendingListEl.style.border = "1px solid #DDD";
        this.pendingListEl.style.borderRadius = "5px";
        this.pendingListEl.style.gap = "0";
        this.pendingListEl.style.overflow = "hidden";

        this.notesSectionEl = this.el("div", "tw-hidden tw-flex tw-flex-col tw-gap-1.5");
        const notesTitle = this.el("div", "tw-text-[11px] tw-font-semibold tw-uppercase tw-tracking-wide tw-text-slate-500");
        notesTitle.textContent = "Attached files";
        this.notesListEl = this.el("div", "tw-flex tw-flex-col tw-gap-1.5");
        this.notesSectionEl.appendChild(notesTitle);
        this.notesSectionEl.appendChild(this.notesListEl);

        this.notesListEl.style.border = "1px solid #DDD";
        this.notesListEl.style.borderRadius = "5px";
        this.notesListEl.style.gap = "0";
        this.notesListEl.style.overflow = "hidden";

        this.loadingOverlayEl = this.el(
            "div",
            "tw-hidden tw-absolute tw-inset-0 tw-z-10 tw-flex tw-flex-col tw-items-center tw-justify-center tw-gap-2 tw-rounded-xl tw-bg-white/85 tw-backdrop-blur-sm"
        );
        const spinner = this.el("div", "tw-h-6 tw-w-6 tw-animate-spin tw-rounded-full tw-border-[3px] tw-border-slate-200 tw-border-t-[var(--fuc-accent)]");
        this.loadingTextEl = this.el("div", "tw-text-xs tw-font-medium tw-text-slate-600");
        this.loadingTextEl.textContent = "Working...";
        this.loadingOverlayEl.appendChild(spinner);
        this.loadingOverlayEl.appendChild(this.loadingTextEl);

        this.rootEl.appendChild(this.triggerEl);
        this.rootEl.appendChild(this.fileInputEl);
        this.rootEl.appendChild(this.errorBannerEl);
        this.rootEl.appendChild(this.pendingSectionEl);
        this.rootEl.appendChild(this.notesSectionEl);
        this.rootEl.appendChild(this.loadingOverlayEl);

        this.container.appendChild(this.rootEl);
    }

    /** Converts the configured accent color into CSS custom properties consumed by Tailwind arbitrary-value classes. */
    private applyAccentColor(): void {
        const hex = /^#[0-9a-fA-F]{6}$/.test(this.config.accentColor) ? this.config.accentColor : "#0078d4";
        this.rootEl.style.setProperty("--fuc-accent", hex);
        this.rootEl.style.setProperty("--fuc-accent-soft", UIService.hexToRgba(hex, 0.08));
        this.rootEl.style.setProperty("--fuc-accent-softer", UIService.hexToRgba(hex, 0.16));
    }

    /** Converts the configured accent color into CSS custom properties consumed by Tailwind arbitrary-value classes. */
    private applyLabelColor(): void {
        const hex = /^#[0-9a-fA-F]{6}$/.test(this.config.labalColor) ? this.config.labalColor : "#000000";
        this.rootEl.style.setProperty("--fuc-labal", hex);
        this.rootEl.style.setProperty("--fuc-labal-soft", UIService.hexToRgba(hex, 0.08));
        this.rootEl.style.setProperty("--fuc-labal-softer", UIService.hexToRgba(hex, 0.16));
    }

    private static hexToRgba(hex: string, alpha: number): string {
        const r = parseInt(hex.slice(1, 3), 16);
        const g = parseInt(hex.slice(3, 5), 16);
        const b = parseInt(hex.slice(5, 7), 16);
        return `rgba(${r}, ${g}, ${b}, ${alpha})`;
    }

    // =======================================================================
    // Trigger (Dropzone card OR compact Button) - built based on displayMode
    // =======================================================================

    private rebuildTrigger(): void {
        const next = this.buildTrigger();
        this.triggerEl.replaceWith(next);
        this.triggerEl = next;
        this.refreshTriggerState();
    }

    private buildTrigger(): HTMLElement {
        this.currentDisplayMode = this.config.displayMode;
        this.triggerAbortController = new AbortController();

        const trigger = this.config.displayMode === "Button" ? this.buildButtonTrigger() : this.buildDropzoneTrigger();

        if (!this.config.readOnly) {
            this.attachDragAndDrop(trigger);
            trigger.addEventListener("click", () => {
                if (trigger.getAttribute("aria-disabled") !== "true") {
                    this.fileInputEl.click();
                }
            }, { signal: this.triggerAbortController.signal });
        } else {
            // trigger.classList.add("tw-hidden");
            trigger.setAttribute("aria-disabled", String(true));
            trigger.setAttribute("aria-disabled", String(true));
            trigger.classList.add("tw-opacity-50");
            trigger.classList.add("tw-cursor-not-allowed");
        }

        if (this.config.displayMode === "Button" && trigger instanceof HTMLButtonElement) {
            trigger.disabled = this.config.readOnly;
        }

        return trigger;
    }

    private buildDropzoneTrigger(): HTMLDivElement {
        const zone = this.el(
            "div",
            "tw-group tw-flex tw-flex-col tw-items-center tw-justify-center tw-gap-1.5 tw-rounded-xl tw-border-2 tw-border-dashed tw-border-slate-300 tw-bg-slate-50 tw-px-6 tw-py-8 tw-text-center tw-transition-colors tw-duration-150 tw-cursor-pointer hover:tw-border-[var(--fuc-accent)] hover:tw-bg-[var(--fuc-accent-soft)] focus:tw-outline-none focus-visible:tw-ring-2 focus-visible:tw-ring-[var(--fuc-accent)] focus-visible:tw-ring-offset-2"
        );
        zone.tabIndex = 0;
        zone.setAttribute("role", "button");
        zone.setAttribute("aria-label", "Upload files. Click, or drag and drop files here.");
        zone.setAttribute("data-fuc", "dropzone");

        const icon = this.el("div", "tw-text-[var(--fuc-accent)] tw-transition-transform tw-duration-150 group-hover:tw-scale-105");
        icon.innerHTML = UIService.iconUpload();

        const text = this.el("div", "tw-text-sm tw-text-slate-700");
        text.innerHTML = `<span class="tw-font-semibold tw-text-[var(--fuc-accent)]">Click to browse</span> or drag and drop files here`;

        const subText = this.el("div", "tw-text-xs tw-text-slate-500");
        subText.setAttribute("data-fuc", "trigger-subtext");
        subText.textContent = this.buildHintText();

        zone.appendChild(icon);
        zone.appendChild(text);
        zone.appendChild(subText);

        zone.addEventListener("keydown", (e: KeyboardEvent) => {
            if ((e.key === "Enter" || e.key === " ") && zone.getAttribute("aria-disabled") !== "true") {
                e.preventDefault();
                this.fileInputEl.click();
            }
        }, { signal: this.triggerAbortController.signal });

        this.triggerLabelEl = text;
        this.triggerSubtextEl = subText;
        return zone;
    }

    private buildButtonTrigger(): HTMLButtonElement {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.style.background = "var(--fuc-accent)";
        btn.style.color = "var(--fuc-labal)";
        btn.style.border = "none";
        btn.style.fontSize = "13px";
        btn.style.fontWeight = "500";
        btn.className = "tw-inline-flex tw-w-fit tw-max-w-full tw-items-center tw-gap-2 tw-rounded-lg tw-bg-[var(--fuc-accent)] tw-px-4 tw-py-2.5 tw-text-sm tw-font-semibold tw-text-white tw-shadow-sm tw-transition-colors tw-duration-150 hover:tw-brightness-110 focus:tw-outline-none focus-visible:tw-ring-2 focus-visible:tw-ring-[var(--fuc-accent)] focus-visible:tw-ring-offset-2 disabled:tw-cursor-not-allowed disabled:tw-bg-slate-300 disabled:tw-text-slate-500";
        btn.setAttribute("data-fuc", "button-trigger");
        btn.setAttribute("aria-label", "Upload files");

        const icon = this.el("span", "tw-flex tw-shrink-0 tw-items-center");
        icon.innerHTML = UIService.iconUpload(18);

        const label = this.el("span", "tw-truncate tw-max-w-[220px]");
        label.textContent = this.config.buttonIdleLabel || "Upload File";

        btn.appendChild(icon);
        btn.appendChild(label);

        this.triggerLabelEl = label;
        this.triggerSubtextEl = null;
        return btn;
    }

    private attachDragAndDrop(trigger: HTMLElement): void {
        const signal = this.triggerAbortController.signal;

        trigger.addEventListener("dragover", (e: DragEvent) => {
            e.preventDefault();
            e.stopPropagation();
            if (trigger.getAttribute("aria-disabled") === "true") return;
            trigger.classList.add("tw-border-[var(--fuc-accent)]", "tw-bg-[var(--fuc-accent-softer)]", "tw-border-solid");
        }, { signal });

        trigger.addEventListener("dragleave", (e: DragEvent) => {
            e.preventDefault();
            e.stopPropagation();
            trigger.classList.remove("tw-border-[var(--fuc-accent)]", "tw-bg-[var(--fuc-accent-softer)]", "tw-border-solid");
        }, { signal });

        trigger.addEventListener("drop", (e: DragEvent) => {
            e.preventDefault();
            e.stopPropagation();
            trigger.classList.remove("tw-border-[var(--fuc-accent)]", "tw-bg-[var(--fuc-accent-softer)]", "tw-border-solid");
            if (trigger.getAttribute("aria-disabled") === "true") return;
            const files = e.dataTransfer?.files;
            if (files && files.length > 0) {
                this.callbacks.onFilesSelected(files);
            }
        }, { signal });
    }

    private buildHiddenFileInput(): HTMLInputElement {
        const input = document.createElement("input");
        input.type = "file";
        input.className = "tw-hidden";
        input.multiple = this.config.allowMultipleFiles;
        input.accept = this.config.allowedExtensions.join(",");
        input.addEventListener("change", () => {
            if (input.files && input.files.length > 0) {
                this.callbacks.onFilesSelected(input.files);
            }
            // Reset so selecting the exact same file again still fires "change"
            input.value = "";
        }, { signal: this.rootAbortController.signal });
        return input;
    }

    private buildHintText(): string {
        const sizeText = ValidationService.formatBytes(this.config.maxFileSizeKB * 1024);
        const typesText = this.config.allowedExtensions.length > 0 ? this.config.allowedExtensions.join(", ") : "any type";
        const countText = this.config.allowMultipleFiles ? `up to ${this.config.maxFileCount} files` : "1 file";
        return `${typesText} · max ${sizeText} per file · ${countText}`;
    }

    /** Reflects staged/attached counts onto the trigger: label text (Button mode) + disabled state at capacity. */
    private refreshTriggerState(): void {
        const totalCount = this.lastPendingFiles.length + this.lastNotes.length;
        const atCapacity = !this.config.allowMultipleFiles ? totalCount >= 1 : this.config.maxFileCount > 0 && totalCount >= this.config.maxFileCount;

        if (this.config.displayMode === "Button" && this.triggerLabelEl) {
            if (totalCount === 0) {
                this.triggerLabelEl.textContent = this.config.buttonIdleLabel || "Upload File";
            } else if (totalCount === 1) {
                const only = this.lastPendingFiles[0]?.fileName ?? this.lastNotes[0]?.filename ?? "1 file";
                this.triggerLabelEl.textContent = only;
                this.triggerLabelEl.title = only;
            } else {
                this.triggerLabelEl.textContent = `${totalCount} files selected`;
            }
        }

        const disable = !this.config.readOnly && atCapacity;
        this.triggerEl.toggleAttribute("aria-disabled", disable);
        this.triggerEl.setAttribute("aria-disabled", String(disable));
        this.triggerEl.classList.toggle("tw-opacity-50", disable);
        this.triggerEl.classList.toggle("tw-cursor-not-allowed", disable);

        // replace
        const replacePCF = atCapacity && this.config.replacePCFWithSelectedFiles
        this.triggerEl.classList.toggle("tw-hidden", replacePCF);

        if (this.triggerSubtextEl) {
            this.triggerSubtextEl.textContent = disable
                ? "Maximum number of files reached."
                : this.buildHintText();
        }
        if (this.config.displayMode === "Button" && this.triggerEl instanceof HTMLButtonElement) {
            this.triggerEl.disabled = disable;
        }
    }

    // =======================================================================
    // Error banner
    // =======================================================================

    /** Displays one or more validation/operation error messages in a dismissible banner. */
    public showErrors(errors: string[]): void {
        if (errors.length === 0) {
            this.clearErrors();
            return;
        }
        this.errorBannerEl.innerHTML = "";
        this.errorBannerEl.className = "tw-flex tw-items-start tw-gap-2 tw-rounded-lg tw-border tw-border-rose-200 tw-bg-rose-50 tw-p-3 tw-text-rose-700";

        const icon = this.el("div", "tw-mt-0.5 tw-shrink-0");
        icon.innerHTML = UIService.iconError();

        const list = this.el("div", "tw-flex tw-flex-1 tw-flex-col tw-gap-1 tw-text-xs tw-leading-relaxed tw-ml-1");
        errors.forEach((msg) => {
            const line = this.el("div", "");
            line.textContent = msg;
            list.appendChild(line);
        });

        const dismiss = this.el(
            "button",
            "tw-flex tw-h-6 tw-w-6 tw-shrink-0 tw-items-center tw-justify-center tw-rounded-full tw-text-rose-700 tw-transition-colors hover:tw-bg-rose-100"
        );
        dismiss.type = "button";
        dismiss.setAttribute("aria-label", "Dismiss");
        dismiss.innerHTML = UIService.iconClose();
        dismiss.addEventListener("click", () => this.clearErrors());

        this.errorBannerEl.appendChild(icon);
        this.errorBannerEl.appendChild(list);
        this.errorBannerEl.appendChild(dismiss);
    }

    public clearErrors(): void {
        this.errorBannerEl.className = "tw-hidden";
        this.errorBannerEl.innerHTML = "";
    }

    // =======================================================================
    // Loading overlay
    // =======================================================================

    public showLoading(message = "Working..."): void {
        this.loadingTextEl.textContent = message;
        this.loadingOverlayEl.classList.remove("tw-hidden");
    }

    public hideLoading(): void {
        this.loadingOverlayEl.classList.add("tw-hidden");
    }

    // =======================================================================
    // Pending files list (staged, not yet saved)
    // =======================================================================

    public renderPendingFiles(files: IPendingFile[]): void {
        this.lastPendingFiles = files;
        this.pendingListEl.innerHTML = "";
        this.pendingSectionEl.classList.toggle("tw-hidden", files.length === 0);
        this.pendingSectionEl.classList.toggle("tw-flex", files.length > 0);

        const canRemove = !this.config.readOnly && this.config.allowDelete;

        files.forEach((pf, index) => {
            const card = this.buildFileCard({
                fileName: pf.fileName,
                fileSize: pf.fileSize,
                mimeType: pf.mimeType,
                previewDataUrl: pf.previewDataUrl,
                statusLabel: "Pending upload",
                statusClass: "tw-bg-amber-100 tw-text-amber-700",
                isLast: index == files.length - 1
            });

            if (canRemove) {
                const removeBtn = this.buildIconButton(UIService.iconTrash(), "Remove file", true);
                removeBtn.addEventListener("click", () => this.callbacks.onRemovePendingFile(pf.id));

                const actions = this.el("div", "tw-flex tw-shrink-0 tw-gap-1");
                actions.appendChild(removeBtn);
                card.appendChild(actions);
            }

            this.pendingListEl.appendChild(card);
        });

        this.refreshTriggerState();
    }

    // =======================================================================
    // Existing notes list (already persisted on the record)
    // =======================================================================
    // NOTE: single-file mode with a Document File Field configured also
    // renders through this same method - index.ts builds a single synthetic
    // IAnnotationRecord (id = FILE_FIELD_NOTE_ID) from the File column's
    // metadata and passes it in here, rather than this file reaching into
    // XrmHelper itself. Keeps UIService presentation-only (see class doc)
    // and avoids maintaining two near-identical card-rendering code paths.
    // =======================================================================

    public renderNotes(notes: IAnnotationRecord[]): void {
        this.lastNotes = notes;
        this.notesListEl.innerHTML = "";
        this.notesSectionEl.classList.toggle("tw-hidden", notes.length === 0);
        this.notesSectionEl.classList.toggle("tw-flex", notes.length > 0);

        const canDownload = this.config.allowDownload;
        const canDelete = !this.config.readOnly && this.config.allowDelete;

        notes.forEach((note, index) => {
            const card = this.buildFileCard({
                fileName: note.filename,
                fileSize: note.filesize,
                mimeType: note.mimetype,
                statusLabel: "Attached",
                statusClass: "tw-bg-emerald-100 tw-text-emerald-700",
                isLast: index == notes.length - 1
            });

            if (canDownload || canDelete) {
                const actions = this.el("div", "tw-flex tw-shrink-0 tw-gap-1");

                if (this.config.allowPreview) {
                    const previewBtn = this.buildIconButton(UIService.iconPreview(), "preview file", false);
                    previewBtn.addEventListener("click", () => this.callbacks.onPreviewNote(note.annotationid, note.filename, note.mimetype));
                    actions.appendChild(previewBtn);
                }

                if (canDownload) {
                    const downloadBtn = this.buildIconButton(UIService.iconDownload(), "Download file", false);
                    downloadBtn.addEventListener("click", () => this.callbacks.onDownloadNote(note.annotationid, note.filename, note.mimetype));
                    actions.appendChild(downloadBtn);
                }

                if (canDelete) {
                    const deleteBtn = this.buildIconButton(UIService.iconTrash(), "Delete file", true);
                    deleteBtn.addEventListener("click", () => this.callbacks.onDeleteNote(note.annotationid));
                    actions.appendChild(deleteBtn);
                }

                card.appendChild(actions);
            }

            this.notesListEl.appendChild(card);
        });

        this.refreshTriggerState();
    }

    // =======================================================================
    // Shared file-card builder (used by both pending & notes lists to avoid
    // duplicated markup/logic)
    // =======================================================================

    private buildFileCard(opts: { fileName: string; fileSize: number; mimeType: string; previewDataUrl?: string; statusLabel: string; statusClass: string; isLast: boolean }): HTMLDivElement {
        const card = this.el("div", "tw-flex tw-items-center tw-gap-3 tw-rounded-lg tw-border tw-border-slate-200 tw-bg-white tw-p-2.5 tw-shadow-sm tw-transition-shadow hover:tw-shadow-md");

        card.style.boxShadow = "none";
        card.style.borderRadius = "0";
        if (opts.isLast == false) {
            card.style.borderBottom = "1px solid #ddd";
        }

        const thumb = this.el("div", "tw-flex tw-h-9 tw-w-9 tw-shrink-0 tw-items-center tw-justify-center tw-overflow-hidden tw-rounded-md tw-bg-slate-50 tw-text-slate-400");
        if (this.config.showPreviewThumbnails && opts.previewDataUrl && opts.mimeType.startsWith("image/")) {
            const img = document.createElement("img");
            img.className = "tw-h-full tw-w-full tw-object-cover";
            img.src = opts.previewDataUrl;
            img.alt = opts.fileName;
            thumb.appendChild(img);
        } else {
            thumb.innerHTML = UIService.iconForMimeType(opts.mimeType);
        }

        const info = this.el("div", "tw-flex tw-min-w-0 tw-flex-1 tw-flex-col tw-gap-0.5");
        const nameRow = this.el("div", "tw-flex tw-min-w-0 tw-items-center tw-gap-2");
        const name = this.el("span", "tw-truncate tw-text-[13px] tw-font-semibold tw-text-slate-800");
        name.textContent = opts.fileName;
        name.title = opts.fileName;

        const badge = this.el("span", `tw-shrink-0 tw-rounded-full tw-px-2 tw-py-0.5 tw-text-[10px] tw-font-semibold tw-uppercase tw-tracking-wide ${opts.statusClass}`);
        badge.textContent = opts.statusLabel;
        nameRow.appendChild(name);
        nameRow.appendChild(badge);

        info.appendChild(nameRow);

        if (this.config.showFileMeta) {
            const metaRow = this.el("div", "tw-truncate tw-text-xs tw-text-slate-500");
            metaRow.textContent = `${ValidationService.formatBytes(opts.fileSize)} · ${opts.mimeType || "unknown type"}`;
            info.appendChild(metaRow);
        }

        card.appendChild(thumb);
        card.appendChild(info);
        return card;
    }

    private buildIconButton(svg: string, ariaLabel: string, danger: boolean): HTMLButtonElement {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = danger
            ? "tw-flex tw-h-8 tw-w-8 tw-items-center tw-justify-center tw-rounded-md tw-text-slate-500 tw-transition-colors hover:tw-bg-rose-50 hover:tw-text-rose-600"
            : "tw-flex tw-h-8 tw-w-8 tw-items-center tw-justify-center tw-rounded-md tw-text-slate-500 tw-transition-colors hover:tw-bg-[var(--fuc-accent-soft)] hover:tw-text-[var(--fuc-accent)]";
        btn.setAttribute("aria-label", ariaLabel);
        btn.title = ariaLabel;
        btn.innerHTML = svg;
        return btn;
    }

    private getPreviewType(
        fileName: string,
        mimeType: string
    ): "image" | "pdf" | "spreadsheet" | "text" | "audio" | "video" | "office" | "unsupported" {

        const extension = fileName
            .split(".")
            .pop()
            ?.toLowerCase() ?? "";

        // Images
        if (
            mimeType.startsWith("image/") ||
            ["png", "jpg", "jpeg", "gif", "webp", "bmp", "svg"].includes(extension)
        ) {
            return "image";
        }

        // PDF
        if (
            mimeType === "application/pdf" ||
            extension === "pdf"
        ) {
            return "pdf";
        }

        // Excel / Spreadsheet
        if (
            [
                "xls",
                "xlsx",
                "xlsm",
                "xlsb",
                "csv"
            ].includes(extension)
        ) {
            return "spreadsheet";
        }

        // Text
        if (
            mimeType.startsWith("text/") ||
            [
                "txt",
                "json",
                "xml",
                "html",
                "css",
                "js",
                "ts",
                "log"
            ].includes(extension)
        ) {
            return "text";
        }

        // Audio
        if (mimeType.startsWith("audio/")) {
            return "audio";
        }

        // Video
        if (mimeType.startsWith("video/")) {
            return "video";
        }

        // Other Office documents
        if (
            [
                "doc",
                "docx",
                "ppt",
                "pptx"
            ].includes(extension)
        ) {
            return "office";
        }

        return "unsupported";
    }


    public showFilePreview(
        fileName: string,
        mimeType: string,
        base64: string
    ): void {

        this.closeFilePreview();

        const modal = document.createElement("div");

        modal.id = "fuc-file-preview-modal";

        modal.className = [
            "tw-fixed",
            "tw-inset-0",
            "tw-z-[99999]",
            "tw-flex",
            "tw-items-center",
            "tw-justify-center",
            "tw-bg-black/60",
            "tw-p-4",
            "tw-model"
        ].join(" ");

        const blob = this.base64ToBlob(base64, mimeType);
        const url = URL.createObjectURL(blob);

        const previewType = this.getPreviewType(fileName, mimeType);

        const safeFileName = this.escapeHtml(fileName);
        const safeMimeType = this.escapeHtml(mimeType);

        let previewHtml = "";

        switch (previewType) {

            case "image":

                previewHtml = `
                    <div class="tw-flex tw-h-full tw-w-full tw-items-center tw-justify-center tw-overflow-auto tw-p-6">
                        <img
                            src="${url}"
                            alt="${safeFileName}"
                            class="tw-max-h-full tw-max-w-full tw-object-contain"
                        />
                    </div>
                `;

                break;
            case "pdf":
                previewHtml = `
                    <iframe
                        src="${url}"
                        class="tw-h-full tw-w-full tw-border-0"
                        title="${safeFileName}">
                    </iframe>
                `;
                break;
            case "spreadsheet":
                previewHtml = this.renderSpreadsheet(base64);
                break;

            case "text":
                previewHtml = `
                    <div class="tw-h-full tw-w-full tw-overflow-auto tw-bg-white tw-p-6">
                        <pre class="tw-whitespace-pre-wrap tw-break-words tw-font-mono tw-text-sm">
                            ${this.escapeHtml(this.base64ToText(base64))}
                        </pre>
                    </div>`;
                break;

            case "audio":
                previewHtml = `
                    <div class="tw-flex tw-h-full tw-items-center tw-justify-center">
                        <audio
                            controls
                            src="${url}">
                        </audio>
                    </div>`;

                break;

            case "video":
                previewHtml = `
                    <div class="tw-flex tw-h-full tw-items-center tw-justify-center tw-bg-black">
                        <video
                            controls
                            class="tw-max-h-full tw-max-w-full"
                            src="${url}">
                        </video>
                    </div>`;

                break;

            case "office":

                previewHtml = `
                    <div class="tw-flex tw-h-full tw-flex-col tw-items-center tw-justify-center">
                        <div class="tw-mb-4 tw-text-5xl">📄</div>
                        <div class="tw-text-lg tw-font-semibold">
                            Preview not available
                        </div>
                        <div class="tw-mt-1 tw-text-sm tw-text-slate-500">
                            ${safeFileName}
                        </div>
                        <a
                            href="${url}"
                            download="${safeFileName}"
                            class="tw-mt-4 tw-rounded-lg tw-bg-blue-600 tw-px-4 tw-py-2 tw-text-white"
                        >
                            Download
                        </a>
                    </div>`;
                break;


            default:
                previewHtml = `
                    <div class="tw-flex tw-h-full tw-flex-col tw-items-center tw-justify-center">
                        <div class="tw-text-lg tw-font-semibold">
                            Preview not available
                        </div>

                        <a href="${url}" download="${safeFileName}" class="tw-mt-4 tw-rounded-lg tw-bg-blue-600 tw-px-4 tw-py-2 tw-text-white">
                            Download
                        </a>
                    </div>`;

                break;
        }

        modal.innerHTML = `
        <div class="tw-flex tw-h-[90vh] tw-w-[95vw] tw-flex-col tw-overflow-hidden tw-rounded-xl tw-bg-white tw-shadow-2xl tw-model-layout">

            <!-- Header -->
            <div class="tw-flex tw-shrink-0 tw-items-center tw-justify-between tw-border-b tw-border-slate-200 tw-bg-white tw-px-5 tw-py-3 tw-model-header">

                <div class="tw-flex tw-min-w-0 tw-items-center tw-gap-3">

                    <div class="tw-flex tw-h-9 tw-w-9 tw-shrink-0 tw-items-center tw-justify-center tw-rounded-lg tw-bg-slate-100 tw-text-slate-600">
                        📄
                    </div>

                    <div class="tw-min-w-0">

                        <div
                            class="tw-truncate tw-font-semibold tw-text-slate-800"
                            title="${safeFileName}"
                        >
                            ${safeFileName}
                        </div>

                        <div class="tw-text-xs tw-text-slate-500">
                            ${safeMimeType}
                        </div>

                    </div>

                </div>

                <button
                    type="button"
                    data-preview-close
                    class="tw-ml-4 tw-flex tw-h-9 tw-w-9 tw-shrink-0 tw-items-center tw-justify-center tw-rounded-lg tw-text-xl tw-text-slate-500 tw-transition hover:tw-bg-slate-100 hover:tw-text-slate-800 focus:tw-outline-none"
                    aria-label="Close preview"
                >
                    &times;
                </button>

            </div>

            <!-- Preview -->
            <div class="tw-relative tw-min-h-0 tw-flex-1 tw-model-body">
                ${previewHtml}
            </div>

        </div>`;

        document.body.appendChild(modal);

        const closeButton = modal.querySelector<HTMLButtonElement>("[data-preview-close]");
        closeButton?.addEventListener("click", () => this.closeFilePreview());

        modal.addEventListener("click", (event) => {
            if (event.target === modal) {
                this.closeFilePreview();
            }
        });

        this.previewKeydownHandler = (event: KeyboardEvent) => {
            if (event.key === "Escape") {
                this.closeFilePreview();
            }
        };

        document.addEventListener(
            "keydown",
            this.previewKeydownHandler
        );

        this.previewBlobUrl = url;
    }

    private renderSpreadsheet(base64: string): string {

        try {

            const arrayBuffer = this.base64ToArrayBuffer(base64);

            const workbook = XLSX.read(arrayBuffer, {
                type: "array"
            });

            if (!workbook.SheetNames.length) {
                return `
                <div class="tw-flex tw-h-full tw-items-center tw-justify-center">
                    <span class="tw-text-slate-500">
                        No worksheets found.
                    </span>
                </div>
            `;
            }

            // const sheetsHtml = workbook.SheetNames.map((sheetName, index) => {
            //     const worksheet = workbook.Sheets[sheetName];
            //     const html = XLSX.utils.sheet_to_html(worksheet, {
            //         id: `fuc-sheet-${index}`
            //     });

            //     return `
            //         <div class="fuc-spreadsheet-sheet" data-sheet-index="${index}">
            //             <div class="fuc-spreadsheet-title">
            //                 ${this.escapeHtml(sheetName)}
            //             </div>
            //             <div class="fuc-spreadsheet-table-wrapper">
            //                 ${html}
            //             </div>
            //         </div>
            //     `;
            // }
            // ).join("");

            const sheetsHtml: string = workbook.SheetNames
                .map((sheetName: string, index: number): string => {

                    const worksheet: XLSX.WorkSheet = workbook.Sheets[sheetName];

                    const rows: unknown[][] = XLSX.utils.sheet_to_json<unknown[]>(
                        worksheet,
                        {
                            header: 1,
                            defval: ""
                        }
                    );

                    if (rows.length === 0) {
                        return `
                        <div
                            class="fuc-spreadsheet-sheet"
                            data-sheet-index="${index}"
                        >
                            <div class="fuc-spreadsheet-title">
                                ${this.escapeHtml(sheetName)}
                            </div>

                            <div class="fuc-spreadsheet-empty">
                                No data
                            </div>
                        </div>
                    `;
                    }

                    const headers: unknown[] = rows[0] ?? [];
                    const dataRows: unknown[][] = rows.slice(1);

                    // -----------------------------
                    // THEAD
                    // -----------------------------

                    const theadHtml: string = `
                    <thead>
                        <tr>
                            ${headers
                            .map((header: unknown): string => `
                                    <th>
                                        ${this.escapeHtml(this.formatCellValue(header))}
                                    </th>
                                `)
                            .join("")}
                        </tr>
                    </thead>
                `;

                    // -----------------------------
                    // TBODY
                    // -----------------------------

                    const tbodyHtml: string = `
                    <tbody>
                        ${dataRows
                            .map((row: unknown[]): string => `
                                <tr>
                                    ${headers
                                    .map(
                                        (
                                            _header: unknown,
                                            columnIndex: number
                                        ): string => {

                                            const value: unknown =
                                                row[columnIndex] ?? "";

                                            return `
                                                    <td>
                                                        ${this.escapeHtml(
                                                this.formatCellValue(value)
                                            )}
                                                    </td>
                                                `;
                                        }
                                    )
                                    .join("")}
                                </tr>
                            `)
                            .join("")}
                    </tbody>
                `;

                    // -----------------------------
                    // COMPLETE SHEET
                    // -----------------------------

                    return `
                    <div
                        class="fuc-spreadsheet-sheet"
                        data-sheet-index="${index}"
                    >

                        <div class="fuc-spreadsheet-title">
                            ${this.escapeHtml(sheetName)}
                        </div>

                        <div class="fuc-spreadsheet-table-wrapper">

                            <table
                                id="fuc-sheet-${index}"
                                class="fuc-spreadsheet-table"
                            >
                                ${theadHtml}
                                ${tbodyHtml}
                            </table>

                        </div>

                    </div>
                `;

                })
                .join("");

            return `
                <div class="fuc-spreadsheet-preview">
                    ${sheetsHtml}
                </div>`;

        } catch (error) {

            console.error(
                "Failed to render spreadsheet:",
                error
            );

            return `
            <div class="tw-flex tw-h-full tw-flex-col tw-items-center tw-justify-center tw-gap-2">
                <div class="tw-text-lg tw-font-semibold tw-text-slate-700">
                    Unable to preview spreadsheet
                </div>

                <div class="tw-text-sm tw-text-slate-500">
                    The file could not be read.
                </div>
            </div>
        `;
        }
    }

    private base64ToBlob(base64: string, mimeType: string): Blob {
        const byteCharacters = atob(base64);
        const byteNumbers = new Uint8Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        return new Blob([byteNumbers], {
            type: mimeType || "application/octet-stream"
        });
    }

    private base64ToText(base64: string): string {
        try {
            const binary = atob(base64);
            const bytes = Uint8Array.from(binary, char => char.charCodeAt(0));

            return new TextDecoder("utf-8").decode(bytes);
        } catch {
            return "Unable to read this file.";
        }
    }

    private base64ToArrayBuffer(base64: string): ArrayBuffer {
        const cleanBase64 = base64.includes(",")
            ? base64.split(",")[1]
            : base64;

        const binaryString = atob(cleanBase64);

        const bytes = new Uint8Array(binaryString.length);

        for (let i = 0; i < binaryString.length; i++) {
            bytes[i] = binaryString.charCodeAt(i);
        }

        return bytes.buffer;
    }

    public closeFilePreview(): void {
        const modal = document.getElementById("fuc-file-preview-modal");
        modal?.remove();
        if (this.previewBlobUrl) {
            URL.revokeObjectURL(this.previewBlobUrl);
            this.previewBlobUrl = undefined;
        }
        if (this.previewKeydownHandler) {
            document.removeEventListener("keydown", this.previewKeydownHandler);
            this.previewKeydownHandler = undefined;
        }
    }

    private formatCellValue(value: unknown): string {
        if (value === null || value === undefined) {
            return "";
        }

        if (value instanceof Date) {
            return value.toLocaleString();
        }

        return String(value);
    }

    private escapeHtml(value: string): string {
        return value
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }

    // =======================================================================
    // Small DOM helper + inline SVG icon set (kept dependency-free / no
    // external icon font required)
    // =======================================================================

    private el<K extends keyof HTMLElementTagNameMap>(tag: K, className: string): HTMLElementTagNameMap[K] {
        const element = document.createElement(tag);
        element.className = className;
        return element;
    }

    private static iconUpload(size = 28): string {
        return `<svg width="${size}" height="${size}" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M10 3a1 1 0 0 1 .707.293l3.5 3.5a1 1 0 0 1-1.414 1.414L11 6.414V13a1 1 0 1 1-2 0V6.414L7.207 8.207a1 1 0 0 1-1.414-1.414l3.5-3.5A1 1 0 0 1 10 3Z" fill="currentColor"/>
            <path d="M4 14a1 1 0 0 1 1 1v1h10v-1a1 1 0 1 1 2 0v1a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-1a1 1 0 0 1 1-1Z" fill="currentColor"/>
        </svg>`;
    }

    private static iconError(): string {
        return `<svg width="18" height="18" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M8 1a7 7 0 1 0 0 14A7 7 0 0 0 8 1Zm.75 3.5v4.25h-1.5V4.5h1.5ZM8 11.75a.875.875 0 1 1 0-1.75.875.875 0 0 1 0 1.75Z" fill="currentColor"/>
        </svg>`;
    }

    private static iconClose(): string {
        return `<svg width="14" height="14" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M3.28 2.22a.75.75 0 0 0-1.06 1.06L6.94 8l-4.72 4.72a.75.75 0 1 0 1.06 1.06L8 9.06l4.72 4.72a.75.75 0 0 0 1.06-1.06L9.06 8l4.72-4.72a.75.75 0 0 0-1.06-1.06L8 6.94 3.28 2.22Z" fill="currentColor"/>
        </svg>`;
    }

    private static iconTrash(): string {
        return `<svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M6.5 1a1 1 0 0 0-1 1v.5H3a.75.75 0 0 0 0 1.5h.35l.6 8.4A2 2 0 0 0 5.94 14h4.12a2 2 0 0 0 1.99-1.6l.6-8.4H13a.75.75 0 0 0 0-1.5h-2.5V2a1 1 0 0 0-1-1h-3ZM6.5 6a.5.5 0 0 1 1 0v5a.5.5 0 0 1-1 0V6Zm3 0a.5.5 0 0 1 1 0v5a.5.5 0 0 1-1 0V6Z" fill="currentColor"/>
        </svg>`;
    }

    private static iconDownload(): string {
        return `<svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M8 1.5a.75.75 0 0 1 .75.75V9.1l1.97-1.97a.75.75 0 1 1 1.06 1.06l-3.25 3.25a.75.75 0 0 1-1.06 0L4.22 8.19a.75.75 0 1 1 1.06-1.06L7.25 9.1V2.25A.75.75 0 0 1 8 1.5ZM3 12.25a.75.75 0 0 1 .75.75v.5c0 .138.112.25.25.25h8a.25.25 0 0 0 .25-.25v-.5a.75.75 0 0 1 1.5 0v.5A1.75 1.75 0 0 1 12 15H4a1.75 1.75 0 0 1-1.75-1.75v-.5A.75.75 0 0 1 3 12.25Z" fill="currentColor"/>
        </svg>`;
    }

    private static iconFileGeneric(): string {
        return `<svg width="20" height="20" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M5 2a1 1 0 0 0-1 1v14a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V7l-5-5H5Z" fill="currentColor" fill-opacity="0.12"/>
            <path d="M5 2a1 1 0 0 0-1 1v14a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V7l-5-5H5Zm6 1.5L14.5 7H11V3.5Z" fill="currentColor"/>
        </svg>`;
    }

    private static iconFilePdf(): string {
        return `<svg width="20" height="20" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M5 2a1 1 0 0 0-1 1v14a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V7l-5-5H5Z" fill="#D93025" fill-opacity="0.15"/>
            <path d="M5 2a1 1 0 0 0-1 1v14a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V7l-5-5H5Zm6 1.5L14.5 7H11V3.5Z" fill="#D93025"/>
            <text x="10" y="15" font-size="5.5" font-family="Segoe UI, sans-serif" fill="#ffffff" text-anchor="middle" font-weight="700">PDF</text>
        </svg>`;
    }

    private static iconFileImage(): string {
        return `<svg width="20" height="20" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M5 2a1 1 0 0 0-1 1v14a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V7l-5-5H5Z" fill="#0078D4" fill-opacity="0.12"/>
            <path d="M5 2a1 1 0 0 0-1 1v14a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V7l-5-5H5Zm6 1.5L14.5 7H11V3.5Z" fill="#0078D4"/>
            <circle cx="8" cy="11.5" r="1" fill="#ffffff"/>
            <path d="m7 15 2.2-2.6L11 14l1.8-2.3L15 15H7Z" fill="#ffffff"/>
        </svg>`;
    }

    private static iconFileWord(): string {
        return `<svg width="20" height="20" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M5 2a1 1 0 0 0-1 1v14a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V7l-5-5H5Z" fill="#185ABD" fill-opacity="0.15"/>
            <path d="M5 2a1 1 0 0 0-1 1v14a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V7l-5-5H5Zm6 1.5L14.5 7H11V3.5Z" fill="#185ABD"/>
            <text x="10" y="15" font-size="5" font-family="Segoe UI, sans-serif" fill="#ffffff" text-anchor="middle" font-weight="700">DOC</text>
        </svg>`;
    }

    private static iconFileExcel(): string {
        return `<svg width="20" height="20" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M5 2a1 1 0 0 0-1 1v14a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V7l-5-5H5Z" fill="#107C41" fill-opacity="0.15"/>
            <path d="M5 2a1 1 0 0 0-1 1v14a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V7l-5-5H5Zm6 1.5L14.5 7H11V3.5Z" fill="#107C41"/>
            <text x="10" y="15" font-size="5" font-family="Segoe UI, sans-serif" fill="#ffffff" text-anchor="middle" font-weight="700">XLS</text>
        </svg>`;
    }

    private static iconFileText(): string {
        return `<svg width="20" height="20" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M5 2a1 1 0 0 0-1 1v14a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V7l-5-5H5Z" fill="currentColor" fill-opacity="0.1"/>
            <path d="M5 2a1 1 0 0 0-1 1v14a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V7l-5-5H5Zm6 1.5L14.5 7H11V3.5Z" fill="currentColor"/>
            <path d="M6.5 10h7v1h-7v-1Zm0 2.5h7v1h-7v-1Zm0-5h4v1h-4v-1Z" fill="#ffffff"/>
        </svg>`;
    }

    private static iconPreview(): string {
        return `<svg width="20" height="20" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
            <!-- Document -->
            <path d="M5.25 1.75H11L15.25 6V16.75C15.25 17.5784 14.5784 18.25 13.75 18.25H5.25C4.42157 18.25 3.75 17.5784 3.75 16.75V3.25C3.75 2.42157 4.42157 1.75 5.25 1.75Z" fill="currentColor" fill-opacity="0.08" />
            <!-- Document outline -->
            <path d="M11 1.75V5.25C11 5.66421 11.3358 6 11.75 6H15.25" stroke="currentColor" stroke-width="1.25" stroke-linecap="round" stroke-linejoin="round" />
            <path d="M5.25 1.75H11L15.25 6V16.75C15.25 17.5784 14.5784 18.25 13.75 18.25H5.25C4.42157 18.25 3.75 17.5784 3.75 16.75V3.25C3.75 2.42157 4.42157 1.75 5.25 1.75Z" stroke="currentColor" stroke-width="1.25" stroke-linejoin="round" />
            <!-- Text lines -->
            <path d="M6.25 8.25H9.25" stroke="currentColor" stroke-opacity="0.55" stroke-width="1" stroke-linecap="round" />
            <path d="M6.25 10.25H9.75" stroke="currentColor" stroke-opacity="0.35" stroke-width="1" stroke-linecap="round" />
            <!-- Preview eye -->
            <path d="M6.15 13.25C7.05 11.95 8.25 11.3 9.5 11.3C10.75 11.3 11.95 11.95 12.85 13.25C11.95 14.55 10.75 15.2 9.5 15.2C8.25 15.2 7.05 14.55 6.15 13.25Z" fill="currentColor" />
            <circle cx="9.5" cy="13.25" r="1.05" fill="white" />
            <circle cx="9.5" cy="13.25" r="0.5" fill="currentColor" />
        </svg>`;
    }

    /** Picks a representative icon for a card based on the file's MIME type. */
    private static iconForMimeType(mimeType: string): string {
        if (mimeType.startsWith("image/")) return UIService.iconFileImage();
        if (mimeType === "application/pdf") return UIService.iconFilePdf();
        if (mimeType.includes("word")) return UIService.iconFileWord();
        if (mimeType.includes("excel") || mimeType.includes("spreadsheet")) return UIService.iconFileExcel();
        if (mimeType.startsWith("text/")) return UIService.iconFileText();
        return UIService.iconFileGeneric();
    }
}