import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { AnnotationService } from "./services/AnnotationService";
import { FileService } from "./services/FileService";
import { NotificationHelper } from "./Helpers/NotificationHelper";
import { UIService } from "./services/UIService";
import { ValidationService } from "./services/ValidationService";
import { XrmHelper } from "./Helpers/XrmHelper";
import { DisplayMode, FILE_FIELD_NOTE_ID, IAnnotationRecord, IBatchUploadResult, IControlConfig, IPersistedState } from "./types/Interfaces";
import { CommonHelper } from "./Helpers/CommonHelper";
import { DataverseHelper } from "./Helpers/DataverseHelper";

/**
 * FileUploadControl
 * ==================
 * A PCF control that reproduces the experience of a native Dataverse File
 * column on a CREATE form, where uploads are otherwise impossible until the
 * record exists.
 *
 * How it works, end to end:
 *  1. On a Create form, the user selects/drops files. They are validated
 *     (extension, MIME type, size, count) and held in memory by
 *     FileService - nothing is sent to Dataverse yet.
 *  2. The bound Multiline Text column is used only to persist small JSON
 *     *metadata* about the staged files (never the binary), which (a)
 *     marks the form dirty so an accidental navigation prompts the user,
 *     and (b) lets the control detect + warn if it gets re-initialized
 *     before save (browser File handles cannot survive that).
 *  3. When the user saves the form, the platform re-invokes updateView()
 *     with a populated context.mode.contextInfo.entityId. This transition
 *     (no id -> id) is the signal that the record now exists.
 *  4. On that transition, the control uploads every staged file as one
 *     Annotation (Note) per file via AnnotationService, using
 *     context.webAPI (the PCF-safe surface for Xrm.WebApi).
 *  5. The control then clears its in-memory state and loads the
 *     now-persisted Notes from Dataverse, rendering them with
 *     download/delete actions - exactly like re-opening a saved record.
 *  6. On an existing record (edit form), the control skips straight to
 *     step 5 on init.
 */
export class FileUploadControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private context!: ComponentFramework.Context<IInputs>;
    private notifyOutputChanged!: () => void;

    private fileService!: FileService;
    private annotationService!: AnnotationService;
    private uiService!: UIService;

    private config!: IControlConfig;

    /** GUID of the current record, or undefined while on an unsaved Create form */
    private entityId: string | undefined;
    private entityLogicalName = "";
    /** OData entity set name (e.g. "accounts") - resolved async via metadata, never guessed */
    private entitySetName: string// | undefined;
    private notifications!: NotificationHelper;

    /** Raw JSON persisted into the bound field */
    private boundValue = "";
    /** Guards against creating duplicate Notes if updateView fires more than once post-save */
    private notesCreatedForSession = false;
    /**
     * Number of Notes already persisted in Dataverse for this record+field,
     * refreshed every time loadExistingNotes() runs. Used by
     * persistStateToBoundField() to decide whether the bound value should
     * stay non-empty (files genuinely exist) or be cleared out entirely
     * (nothing exists - so a "Business Required" flag on this column can
     * correctly block Save).
     */
    private existingNotesCount = 0;
    /**
     * Identifies which bound field THIS control instance lives on, so that
     * when several FileUploadControl instances sit on the same form (each
     * bound to its own dedicated state field) every instance only ever
     * creates/shows the Notes it is responsible for - see resolveSourceFieldTag().
     */
    private sourceFieldTag: string | undefined;

    /**
     * init
     * PCF lifecycle entry point. Wires up services, restores any prior
     * state from the bound field, and kicks off async initialization
     * (entity metadata + loading existing notes when applicable).
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        _state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this.context = context;
        this.notifyOutputChanged = notifyOutputChanged;

        console.log("🔥🔥🔥 NEW PCF VERSION 2026-08-25 🔥🔥🔥");

        // PCF does not auto-resize the container to content by default -
        // ask the platform to track/allow resize so the file list can grow.
        this.context.mode.trackContainerResize(true);

        this.config = CommonHelper.ReadConfig(context);
        this.fileService = new FileService();
        this.annotationService = new AnnotationService(context.webAPI);
        this.notifications = new NotificationHelper(context);

        this.entityLogicalName = context.mode.contextInfo?.entityTypeName ?? "";
        this.entityId = CommonHelper.NormalizeGuid(context.mode.contextInfo?.entityId);
        this.sourceFieldTag = CommonHelper.ResolveSourceFieldTag(context);

        this.uiService = new UIService(
            container,
            {
                onFilesSelected: (files) => void this.handleFilesSelected(files),
                onRemovePendingFile: (id) => this.handleRemovePendingFile(id),
                onDownloadNote: (id, name, mime) => void this.handleDownloadNote(id, name, mime),
                onPreviewNote: (id, fileName, mimeType) => void this.handlePreviewNote(id, fileName, mimeType),
                onDeleteNote: (id) => void this.handleDeleteNote(id),
                onDeleteFromFileField: () => void this.handleDeleteFromFileField()
            },
            this.config
        );

        this.restoreStateFromBoundField(context.parameters.value.raw);
        void this.initializeAsync();
    }

    private async handlePreviewNote(annotationId: string, fileName: string, mimeType: string): Promise<void> {
        this.uiService.showLoading(`Preparing "${fileName}"...`);
        try {
            const useFileFieldPreview = !this.config.allowMultipleFiles && !!this.config.documentFileFieldLogicalName;

            if (useFileFieldPreview) {
                var data = await DataverseHelper.GetFileFieldContent(this.entitySetName, this.config.documentFileFieldLogicalName, this.entityId);
                this.uiService.showFilePreview(
                    fileName,
                    mimeType,
                    data?.base64 ?? ''
                );
            }
            else {
                const base64 = await this.annotationService.retrieveAnnotationBody(annotationId);
                this.uiService.showFilePreview(
                    fileName,
                    mimeType,
                    base64
                );
            }
        } catch (err) {
            const message = `Unable to preview "${fileName}": ${CommonHelper.DescribeError(err)}`;
            this.notifications.showError(message);
            this.uiService.showErrors([message]);
        } finally {
            this.uiService.hideLoading();
        }
    }

    /**
     * updateView
     * Called by the platform whenever bound data or inputs change. This is
     * also where we detect the Create -> Saved transition (see class doc).
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this.context = context;
        this.config = CommonHelper.ReadConfig(context);
        this.uiService.updateConfig(this.config);

        if (!this.sourceFieldTag) {
            // Field metadata (attributes.LogicalName) can be unavailable on
            // the very first render on some hosts - retry here so we don't
            // permanently fall back to "unfiltered" for the whole session.
            this.sourceFieldTag = CommonHelper.ResolveSourceFieldTag(context);
        }

        const newEntityId = CommonHelper.NormalizeGuid(context.mode.contextInfo?.entityId);

        if (!this.entityId && newEntityId) {
            // Record just received its GUID - the form was saved.
            this.entityId = newEntityId;
            void this.handleRecordSaved();
        } else {
            this.entityId = newEntityId;
        }
    }

    /** Returns the current value of the bound property to the platform. */
    public getOutputs(): IOutputs {
        return { value: this.boundValue };
    }

    /** Cleans up event listeners when the control is removed from the DOM. */
    public destroy(): void {
        this.uiService?.destroy();
    }

    // =========================================================================
    // Initialization helpers
    // =========================================================================

    private async initializeAsync(): Promise<void> {
        await this.resolveEntitySetName();

        if (this.entityId) {
            // Editing an existing record - load whatever Notes already exist.
            await this.loadExistingNotes();
            this.persistStateToBoundField()
        } else {
            // Fresh Create form - just render whatever we restored from the
            // bound field's metadata (typically nothing, on a truly new form).
            this.uiService.renderPendingFiles(this.fileService.getPendingFiles());
        }
    }

    /**
     * Resolves the correct OData entity set name for the current entity via
     * platform metadata. This is required for the @odata.bind navigation
     * property used when creating notes, and MUST NOT be guessed by naive
     * pluralization - Dataverse entity set names are not always a simple
     * "add an s" transformation of the logical name.
     */
    private async resolveEntitySetName(): Promise<void> {
        if (this.entitySetName || !this.entityLogicalName) {
            return;
        }
        try {
            const metadata = await this.context.utils.getEntityMetadata(this.entityLogicalName);
            this.entitySetName = metadata.EntitySetName;
        } catch (err: any) {
            console.error("FileUploadControl: failed to resolve entity metadata for", this.entityLogicalName, err);
            this.notifications.showError(`FileUploadControl: failed to resolve entity metadata for ${this.entityLogicalName} :: ${err.message}`);
            this.uiService.showErrors(["Unable to resolve entity metadata required to attach files. Please contact your administrator."]);
        }
    }

    // =========================================================================
    // User-driven events (selecting / removing staged files)
    // =========================================================================

    private async handleFilesSelected(files: FileList): Promise<void> {
        this.uiService.clearErrors();
        this.notifications.clear();

        const { added, errors } = await this.fileService.addFiles(this.existingNotesCount, files, this.config);

        if (errors.length > 0) {
            this.notifications.showError(errors.join(''));
            this.uiService.showErrors(errors);
        }

        if (added.length > 0) {
            this.persistStateToBoundField();
            this.uiService.renderPendingFiles(this.fileService.getPendingFiles());

            if (this.entityId) {
                this.handleRecordSaved();
            }
        }
    }

    private handleRemovePendingFile(id: string): void {
        this.fileService.removeFile(id);
        this.persistStateToBoundField();
        this.uiService.renderPendingFiles(this.fileService.getPendingFiles());
    }

    // =========================================================================
    // Record-saved handling: turns staged files into real Notes
    // =========================================================================

    private async handleRecordSaved(): Promise<void> {
        await this.resolveEntitySetName();

        // if (this.notesCreatedForSession) {
        //     // Already uploaded during a prior pass of this same save cycle -
        //     // just refresh the list to reflect the current server state.
        //     await this.loadExistingNotes();
        //     return;
        // }

        const pending = this.fileService.getPendingFiles();
        if (pending.length === 0) {
            await this.loadExistingNotes();
            return;
        }

        if (!this.entitySetName || !this.entityId) {
            this.notifications.showError("Unable to save attachments: could not resolve the parent record. Please try again.");
            this.uiService.showErrors(["Unable to save attachments: could not resolve the parent record. Please try again."]);
            return;
        }

        this.uiService.showLoading(`Uploading ${pending.length} file(s)...`);
        try {
            var result: IBatchUploadResult = { succeeded: [], failed: [] };

            // Single File
            if (pending.length == 1) {
                // Single file mode with a Document File Field configured -
                // upload straight to the native File column (source of
                // truth for preview - see loadFromDocumentField).
                if (this.config.documentFileFieldLogicalName) {
                    if (this.config.deleteNoteAfterFileFieldUpload) {
                        // Upload ONLY to the File column - no Annotation is
                        // created at all, so the file is never uploaded
                        // twice (as opposed to creating a Note and then
                        // deleting it right after).
                        result = await this.annotationService.uploadToFileFieldOnly(
                            this.entitySetName,
                            this.entityId,
                            this.config.documentFileFieldLogicalName,
                            pending[0]
                        );
                    } else {
                        // Keep the previous behavior: also create/keep an
                        // Annotation (Note) copy alongside the File column
                        // value, as a redundant audit-trail copy.
                        result = await this.annotationService.createAnnotationAndUploadToFileField(
                            this.entityLogicalName,
                            this.entitySetName,
                            this.entityId,
                            this.config.documentFileFieldLogicalName,
                            pending[0],
                            this.sourceFieldTag
                        )
                    }
                }
                // Single file, no Document File Field configured - falls
                // back to the normal single-Note path.
                else {
                    result = await this.annotationService.createAnnotationsForFiles(
                        this.entityLogicalName,
                        this.entitySetName,
                        this.entityId,
                        pending,
                        (completed, total, fileName) => this.uiService.showLoading(`Uploading "${fileName}" (${completed}/${total})...`),
                        this.sourceFieldTag
                    );
                }
            }

            // Mulitiple files
            else if (pending.length > 0) {
                result = await this.annotationService.createAnnotationsForFiles(
                    this.entityLogicalName,
                    this.entitySetName,
                    this.entityId,
                    pending,
                    (completed, total, fileName) => this.uiService.showLoading(`Uploading "${fileName}" (${completed}/${total})...`),
                    this.sourceFieldTag
                );
            }

            if (result.failed.length > 0) {
                this.notifications.showError(result.failed.map((f) => `Failed to upload "${f.fileName}": ${f.error}`).join(''));
                this.uiService.showErrors(result.failed.map((f) => `Failed to upload "${f.fileName}": ${f.error}`));
            }

            // Files that succeeded should not be re-uploaded on a future save;
            // clear staged state entirely and mark the session as processed.
            this.notesCreatedForSession = true;
            this.fileService.clear();

            // IMPORTANT: refresh existingNotesCount (via loadExistingNotes)
            // BEFORE persisting the bound field's state. Otherwise the bound
            // value would briefly be written as "no files at all" right
            // after a successful upload, which would incorrectly trip a
            // "Business Required" validation on the very next save even
            // though files were just attached.
            await this.loadExistingNotes();
            this.persistStateToBoundField();

            if (result.succeeded.length > 0) {
                this.applyPostUploadFormBehavior();
            }
        } catch (err) {
            this.notifications.showError(`Unable to save attachments: ${CommonHelper.DescribeError(err)}`);
            this.uiService.showErrors([`Unable to save attachments: ${CommonHelper.DescribeError(err)}`]);
        } finally {
            this.uiService.hideLoading();
        }
    }

    // =========================================================================
    // Existing notes: load / download / delete
    // =========================================================================

    private async loadExistingNotes(): Promise<void> {
        if (!this.entityId) { return; }

        this.uiService.showLoading("Loading attached files...");
        try {
            // Single-file mode with a Document File Field configured: the
            // file already lives natively in that File column (see
            // handleRecordSaved), so preview/download/delete must read from
            // THAT, not from a possibly-absent/stale Note. Every other case
            // (multiple files, or single-file mode with no File column
            // configured) reads from Notes as before.
            const useFileFieldPreview = !this.config.allowMultipleFiles && !!this.config.documentFileFieldLogicalName;

            if (useFileFieldPreview) {
                await this.loadFromDocumentField();
            } else {
                await this.loadFromNotes();
            }
        } catch (err) {
            this.notifications.showError(`Unable to load attached files: ${CommonHelper.DescribeError(err)}`);
            this.uiService.showErrors([`Unable to load attached files: ${CommonHelper.DescribeError(err)}`]);
        } finally {
            this.uiService.hideLoading();
        }
    }

    /** Default path: reads/renders existing Annotations (Notes) for this record+field. */
    private async loadFromNotes(): Promise<void> {
        const notes = await this.annotationService.retrieveAnnotations(this.entityId!, this.sourceFieldTag);
        this.existingNotesCount = notes.length;
        this.uiService.renderNotes(notes);
        this.uiService.renderPendingFiles(this.fileService.getPendingFiles());

        // Keep the "uploaded" form layout (hidden control / visible
        // Document field) consistent across reloads, not just the instant
        // the upload happened - and correctly reversed when zero remain.
        this.applyPostUploadFormBehavior();
    }

    /**
     * Single-file + documentFileFieldLogicalName path: reads the File
     * column's metadata (name/size/mimetype) via XrmHelper.getFileFieldInfo,
     * and renders it through the SAME renderNotes() card UI as regular
     * Notes by building one synthetic IAnnotationRecord tagged with
     * FILE_FIELD_NOTE_ID. handleDownloadNote/handleDeleteNote branch on
     * that sentinel to route to the File-column APIs instead of
     * AnnotationService.
     */
    private async loadFromDocumentField(): Promise<void> {
        if (!this.entitySetName || !this.entityId || !this.config.documentFileFieldLogicalName) {
            this.uiService.renderNotes([]);
            this.uiService.renderPendingFiles(this.fileService.getPendingFiles());
            return;
        }

        const fileInfo = await this.annotationService.getFileFieldMetadata(
            this.entityLogicalName,
            this.entityId,
            this.config.documentFileFieldLogicalName
        );

        if (fileInfo) {
            this.existingNotesCount = 1;
            this.uiService.renderNotes([{
                annotationid: FILE_FIELD_NOTE_ID,
                filename: fileInfo.fileName,
                mimetype: fileInfo.mimeType,
                filesize: fileInfo.fileSize
            }]);
            this.applyPostUploadFormBehavior();
        } else {
            this.existingNotesCount = 0;
            this.uiService.renderNotes([]);
        }

        this.uiService.renderPendingFiles(this.fileService.getPendingFiles());
    }

    private async handleDownloadNote(annotationId: string, fileName: string, mimeType: string): Promise<void> {
        this.uiService.clearErrors();
        this.notifications.clear();

        this.uiService.showLoading(`Preparing "${fileName}" for download...`);
        try {
            let base64: string;
            if (annotationId === FILE_FIELD_NOTE_ID) {
                if (!this.entitySetName || !this.entityId || !this.config.documentFileFieldLogicalName) {
                    throw new Error("Could not resolve the Document File Field to download from.");
                }
                base64 = await this.annotationService.downloadFileFieldValue(
                    this.entitySetName,
                    this.entityId,
                    this.config.documentFileFieldLogicalName
                );
            } else {
                base64 = await this.annotationService.retrieveAnnotationBody(annotationId);
            }
            this.triggerBrowserDownload(base64, fileName, mimeType);
        } catch (err) {
            this.notifications.showError(`Unable to download "${fileName}": ${CommonHelper.DescribeError(err)}`);
            this.uiService.showErrors([`Unable to download "${fileName}": ${CommonHelper.DescribeError(err)}`]);
        } finally {
            this.uiService.hideLoading();
        }
    }

    private async handleDeleteNote(annotationId: string): Promise<void> {
        if (this.config.confirmBeforeDelete) {
            // Uses the native Dataverse confirm dialog (Xrm.Navigation) via
            // XrmHelper when reachable, falling back to window.confirm.
            const confirmed = await CommonHelper.Confirm(
                "Delete file",
                "Delete this file? This action cannot be undone.",
                "Delete",
                "Cancel"
            );
            if (!confirmed) {
                return;
            }
        }

        this.uiService.clearErrors();
        this.notifications.clear();

        this.uiService.showLoading("Deleting file...");
        try {
            if (annotationId === FILE_FIELD_NOTE_ID) {
                if (!this.entityId || !this.config.documentFileFieldLogicalName) {
                    throw new Error("Could not resolve the Document File Field to delete.");
                }
                await DataverseHelper.DeleteFileField(this.entitySetName, this.entityId, this.config.documentFileFieldLogicalName)
                // await this.annotationService.deleteFileFieldValue(
                //     this.entityLogicalName,
                //     this.entityId,
                //     this.config.documentFileFieldLogicalName
                // );
            } else {
                await this.annotationService.deleteAnnotation(annotationId);
            }
            await this.loadExistingNotes();
            // Keep the bound field's "Required" behavior accurate: if that
            // was the last remaining file, this clears the value back to
            // empty; otherwise it just refreshes attachedFileCount.
            this.persistStateToBoundField();
            // Reverses hideControlAfterUpload/showDocumentFieldAfterUpload
            // if that was the last remaining file, so the control becomes
            // usable again for a re-upload instead of staying hidden.
            this.applyPostUploadFormBehavior();
        } catch (err) {
            this.notifications.showError(`Unable to delete file: ${CommonHelper.DescribeError(err)}`);
            this.uiService.showErrors([`Unable to delete file: ${CommonHelper.DescribeError(err)}`]);
        } finally {
            this.uiService.hideLoading();
        }
    }

    private handleDeleteFromFileField() {

    }

    /** Converts a base64 payload into a Blob and triggers a native browser download. */
    private triggerBrowserDownload(base64: string, fileName: string, mimeType: string): void {
        const byteCharacters = atob(base64);
        const byteNumbers = new Array<number>(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        const blob = new Blob([byteArray], { type: mimeType || "application/octet-stream" });

        const url = URL.createObjectURL(blob);
        const anchor = document.createElement("a");
        anchor.href = url;
        anchor.download = fileName;
        document.body.appendChild(anchor);
        anchor.click();
        document.body.removeChild(anchor);
        URL.revokeObjectURL(url);
    }

    /**
     * Optional post-upload form behavior, driven entirely through
     * XrmHelper (Global Xrm) since neither capability exists on the
     * documented PCF context surface:
     *  - hideControlAfterUpload: hides THIS control's own field (the bound
     *    "value" state field) once at least one file has been uploaded.
     *  - showDocumentFieldAfterUpload: reveals the Dataverse File column
     *    field so the now-attached document becomes visible natively.
     * Both are gated behind enableXrmIntegration and fail silently (via
     * XrmHelper's own try/catch + console.warn) if Xrm cannot be reached -
     * this never blocks or breaks the upload flow itself.
     */
    private applyPostUploadFormBehavior(): void {
        // if (!this.config.enableXrmIntegration) {
        //     return;
        // }

        // Driven by the CURRENT count, not just "an upload just happened" -
        // this makes the behavior reversible: deleting the last remaining
        // file (Note or File-column value) un-hides this control again and
        // re-hides the Document field, rather than leaving the form stuck
        // showing "uploaded" layout with nothing actually attached.
        const hasFiles = this.existingNotesCount > 0;

        if (this.config.hideControlAfterUpload && this.sourceFieldTag) {
            CommonHelper.SetVisible([this.sourceFieldTag], !hasFiles);
        }

        if (this.config.showDocumentFieldAfterUpload && this.config.documentFileFieldLogicalName) {
            CommonHelper.SetVisible([this.config.documentFileFieldLogicalName], hasFiles);
        }
    }

    // =========================================================================
    // Bound-field state persistence (metadata only - never the file binary)
    // =========================================================================
    private persistStateToBoundField(): void {
        const pendingFileMeta = this.fileService.getPendingFiles().map((pf) => ({
            id: pf.id,
            fileName: pf.fileName,
            fileSize: pf.fileSize,
            mimeType: pf.mimeType
        }));

        const hasAnyFiles = pendingFileMeta.length > 0 || this.existingNotesCount > 0;

        if (!hasAnyFiles) {
            // Nothing staged in memory AND nothing already attached to the
            // record - leave the bound value genuinely empty (not a JSON
            // blob with an empty pendingFileMeta array) so Dataverse's
            // native "Business Required" validation on this column behaves
            // exactly like any other required field: Save is blocked with
            // the standard required-field notification until at least one
            // file is staged or attached.
            this.boundValue = "";
            this.notifyOutputChanged();
            return;
        }
        const state: IPersistedState = {
            pendingFileMeta,
            notesCreatedForSession: this.notesCreatedForSession,
            attachedFileCount: this.existingNotesCount
        };
        this.boundValue = JSON.stringify(state);
        this.notifyOutputChanged();
    }

    /**
     * Restores whatever we can from the bound field's JSON state. Note that
     * the actual File binaries can never be recovered this way (browsers do
     * not allow re-materializing a File object from metadata alone) - if
     * metadata indicates files were staged but FileService has nothing in
     * memory, we surface a clear warning so the user knows to re-select them.
     */
    private restoreStateFromBoundField(raw: string | null): void {
        if (!raw) {
            this.boundValue = "";
            return;
        }
        this.boundValue = raw;
        try {
            const state = JSON.parse(raw) as IPersistedState;
            this.notesCreatedForSession = !!state.notesCreatedForSession;
            this.existingNotesCount = state.attachedFileCount ?? 0;

            if (state.pendingFileMeta && state.pendingFileMeta.length > 0 && this.fileService.getCount() === 0) {
                let sMsg = "Previously selected file(s) could not be restored after the form was reloaded " +
                    "(browsers do not allow this). Please re-select the file(s) below before saving: " +
                    state.pendingFileMeta.map((m) => m.fileName).join(", ");

                this.notifications.showError(sMsg);
                this.uiService.showErrors([sMsg]);
            }
        } catch {
            // Bound field did not contain valid JSON (e.g. legacy/empty value) - ignore safely.
        }
    }

    // Kept for reference/tests: exposes whether Notes have already been
    // created for the current save cycle (used indirectly via handleRecordSaved).
    private get hasCreatedNotesThisSession(): boolean {
        return this.notesCreatedForSession;
    }
}

// Re-export used by AnnotationRecord typing consumers/tests, if any.
export type { IAnnotationRecord };