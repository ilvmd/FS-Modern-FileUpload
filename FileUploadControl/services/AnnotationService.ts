import { IAnnotationRecord, IBatchUploadResult, IPendingFile } from "../types/Interfaces";
import { FileService } from "./FileService";
import { XrmHelper } from "../Helpers/XrmHelper";

/**
 * AnnotationService
 * -----------------
 * All Dataverse CRUD against the "annotation" (Note) entity.
 *
 * Uses context.webAPI rather than the global window.Xrm object. This is
 * the officially documented, sandbox-safe surface for Xrm.WebApi inside a
 * PCF control: it behaves identically to Xrm.WebApi, works whether or not
 * the control happens to be running inside an iframe with access to the
 * parent Xrm object, and is the pattern recommended by Microsoft's PCF
 * best practices.
 */
export class AnnotationService {
    /**
     * Prefix used to tag each annotation's notetext with the logical name of
     * the bound field/control instance that created it. This lets multiple
     * FileUploadControl instances on the same form (bound to different
     * fields) each show only the files THEY uploaded, instead of every
     * Note attached to the record regardless of source field.
     *
     * notetext (rather than subject) is used deliberately: it is not shown
     * prominently in the default Notes timeline UI the way subject is, so
     * tagging this way does not change what end users see when browsing
     * notes natively.
     */
    private static readonly SOURCE_TAG_PREFIX = "flc-source:";

    public constructor(private readonly webAPI: ComponentFramework.WebApi) { }

    public static buildSourceTag(fieldLogicalName: string): string {
        return `${AnnotationService.SOURCE_TAG_PREFIX}${fieldLogicalName}`;
    }

    /** Escapes single quotes for safe inclusion inside an OData string literal. */
    private static odataEscape(value: string): string {
        return value.replace(/'/g, "''");
    }

    /**
     * Creates one annotation per pending file, associated to the given
     * parent record. Uploads run sequentially (not in parallel) to keep
     * peak memory bounded when many/large base64 payloads are involved,
     * and to make partial-failure reporting straightforward.
     *
     * @param entityLogicalName logical name of the parent entity, e.g. "account"
     * @param entitySetName     OData entity set name for the parent entity,
     *                          e.g. "accounts" - resolved via getEntityMetadata,
     *                          NOT guessed, since pluralization is not
     *                          reliable for all Dataverse entities.
     * @param entityId          GUID of the parent record (without braces)
     * @param pendingFiles      files staged in memory by FileService
     * @param onProgress        optional progress callback for UI feedback
     */
    public async createAnnotationsForFiles(
        entityLogicalName: string,
        entitySetName: string,
        entityId: string,
        pendingFiles: IPendingFile[],
        onProgress?: (completed: number, total: number, fileName: string) => void,
        sourceFieldTag?: string
    ): Promise<IBatchUploadResult> {
        const result: IBatchUploadResult = { succeeded: [], failed: [] };

        for (let i = 0; i < pendingFiles.length; i++) {
            const pendingFile = pendingFiles[i];
            try {
                const annotationId = await this.createAnnotation(entityLogicalName, entitySetName, entityId, pendingFile, sourceFieldTag);
                result.succeeded.push(annotationId);
            } catch (err) {
                result.failed.push({ fileName: pendingFile.fileName, error: AnnotationService.extractErrorMessage(err) });
            }
            onProgress?.(i + 1, pendingFiles.length, pendingFile.fileName);
        }

        return result;
    }

    /** Creates a single annotation record from one staged file. */
    public async createAnnotation(
        entityLogicalName: string,
        entitySetName: string,
        entityId: string,
        pendingFile: IPendingFile,
        sourceFieldTag?: string
    ): Promise<string> {
        const documentBody = await FileService.fileToBase64(pendingFile.file);

        // The polymorphic "objectid" lookup on annotation is bound using the
        // documented "objectid_<entitylogicalname>@odata.bind" navigation
        // property. objecttypecode records which entity the lookup points to.
        const annotation: Record<string, unknown> = {
            subject: `<${sourceFieldTag}>: ${pendingFile.fileName}`,
            filename: pendingFile.fileName,
            mimetype: pendingFile.mimeType,
            documentbody: documentBody,
            objecttypecode: entityLogicalName,
            [`objectid_${entityLogicalName}@odata.bind`]: `/${entitySetName}(${entityId})`
        };

        if (sourceFieldTag) {
            annotation["notetext"] = AnnotationService.buildSourceTag(sourceFieldTag);
        }

        const result = await this.webAPI.createRecord("annotation", annotation as ComponentFramework.WebApi.Entity);
        return result.id;
    }

    /**
     * Creates an annotation record and uploads the same file
     * to a Dataverse File column.
     */
    public async createAnnotationAndUploadToFileField(
        entityLogicalName: string,
        entitySetName: string,
        entityId: string,
        fileFieldLogicalName: string,
        pendingFile: IPendingFile,
        sourceFieldTag?: string
    ): Promise<IBatchUploadResult> {
        // 1. Convert file to Base64 for the Annotation
        const documentBody = await FileService.fileToBase64(pendingFile.file);

        // 2. Create Annotation
        const annotation: Record<string, unknown> = {
            subject: `<${sourceFieldTag}>: ${pendingFile.fileName}`,
            filename: pendingFile.fileName,
            mimetype: pendingFile.mimeType,
            documentbody: documentBody,
            objecttypecode: entityLogicalName,
            [`objectid_${entityLogicalName}@odata.bind`]: `/${entitySetName}(${entityId})`
        };

        if (sourceFieldTag) {
            annotation["notetext"] = AnnotationService.buildSourceTag(sourceFieldTag);
        }

        const result = await this.webAPI.createRecord(
            "annotation",
            annotation as ComponentFramework.WebApi.Entity
        );

        const annotationId = result.id;

        // 3. Upload the same file to the Dataverse File column
        var {IsUploaded, ErrorMessage} = await this.uploadFileToDataverseField(
            entitySetName,
            entityId,
            fileFieldLogicalName,
            pendingFile.file
        );

        var oResult: IBatchUploadResult;

        if (IsUploaded) {
            oResult = { succeeded: [annotationId], failed: [] };
        }
        else {
            oResult = { succeeded: [], failed: [{ fileName: pendingFile.fileName, error: ErrorMessage }] };
        }

        return oResult;
    }

    /**
     * Uploads a File directly to a Dataverse File column.
     */
    private async uploadFileToDataverseField(
        entitySetName: string,
        entityId: string,
        fileFieldLogicalName: string,
        file: File
    ): Promise<{ IsUploaded: boolean, ErrorMessage: string }> {

        // const clientUrl = Xrm.Utility.getGlobalContext().GetClientUrl();
        // const clientUrl = window.location.origin;
        
        // Uses the Global Xrm object (Xrm.Utility.getGlobalContext().GetClientUrl())
        // via XrmHelper - the Microsoft-recommended way to resolve the
        // environment's base URL - with a safe fallback to
        // window.location.origin when Xrm cannot be reached (e.g. the
        // standalone `npm start` test harness).
        const clientUrl = XrmHelper.GetClientUrl();
        
        const url =
            `${clientUrl}/api/data/v9.2/` +
            `${entitySetName}(${entityId})/` +
            `${fileFieldLogicalName}`;

        const response = await fetch(url, {
            // Per Microsoft's documented Web API contract for a single-request
            // file-column upload, the verb is PATCH (not PUT - PUT is not a
            // valid/documented verb for this endpoint and can be silently
            // rejected or behave inconsistently across environments).
            method: "PATCH",
            headers: {
                "Content-Type": "application/octet-stream",
                "x-ms-file-name": file.name,
                "OData-Version": "4.0",
                "OData-MaxVersion": "4.0"
            },
            body: file
        });

        var IsUploaded = true;
        var ErrorMessage = "";

        if (!response.ok) {
            const errorText = await response.text();
            IsUploaded = false;
            ErrorMessage = `Failed to upload file to '${fileFieldLogicalName}'. ` + `Status: ${response.status}. ${errorText}`;
        }

        return { IsUploaded, ErrorMessage };
    }

    /**
     * Reads metadata (name/size/mimetype) for whatever file currently sits
     * in a Dataverse File column, WITHOUT transferring the binary content.
     *
     * IMPORTANT: this uses the pattern Microsoft actually documents for
     * this exact purpose - $expand={entityLogicalName}_FileAttachments
     * from the PARENT record (see "Use file column data in Microsoft
     * Dataverse"). An earlier version of this control queried the
     * `fileattachments` entity set directly with a
     * `$filter=_objectid_value eq ... and regardingfieldname eq ...` -
     * that is NOT a documented/supported query shape: the FileAttachment
     * table's `objectid` lookup only targets a fixed allow-list of system
     * tables in Microsoft's own reference docs, so filtering it directly
     * silently returned nothing for ordinary entities (account, contact,
     * custom tables, etc.) - which is why the preview, the Required
     * validation, and the max-file-count/disabled logic (all driven by
     * this same result) looked broken even though the file itself had
     * uploaded successfully. The $expand approach reads the relationship
     * FROM the specific parent record instead, which is the supported,
     * documented, always-correct way to do this - and uses the sandboxed
     * context.webAPI rather than a raw fetch.
     */
    public async getFileFieldMetadata(
        entityLogicalName: string,
        entityId: string,
        fileFieldLogicalName: string
    ): Promise<{ fileName: string; fileSize: number; mimeType: string } | null> {
        const relationshipName = `${entityLogicalName}_FileAttachments`;
        const escapedField = AnnotationService.odataEscape(fileFieldLogicalName);
        const query =
            `?$select=${entityLogicalName}id` +
            `&$expand=${relationshipName}($select=filename,filesizeinbytes,mimetype,regardingfieldname;` +
            `$filter=regardingfieldname eq '${escapedField}')`;

        let record: Record<string, unknown>;
        try {
            record = await this.webAPI.retrieveRecord(entityLogicalName, entityId, query);
        } catch {
            // Some entities' primary id attribute isn't "{logicalname}id"
            // (a handful of system tables use a different name) - retry
            // once without the top-level $select, relying on $expand alone.
            record = await this.webAPI.retrieveRecord(
                entityLogicalName,
                entityId,
                `?$expand=${relationshipName}($select=filename,filesizeinbytes,mimetype,regardingfieldname;$filter=regardingfieldname eq '${escapedField}')`
            );
        }

        const attachments = (record[relationshipName] as Record<string, unknown>[] | undefined) ?? [];
        const match = attachments[0];
        if (!match) {
            return null;
        }

        return {
            fileName: (match["filename"] as string) ?? "",
            fileSize: Number(match["filesizeinbytes"] ?? 0),
            mimeType: (match["mimetype"] as string) ?? "application/octet-stream"
        };
    }

    /**
     * Uploads a file directly to a Dataverse File column WITHOUT creating
     * any Annotation/Note at all - used when
     * IControlConfig.deleteNoteAfterFileFieldUpload is true. Deliberately
     * skips Note creation up front (rather than creating one and deleting
     * it right after) so the file is only ever base64-encoded/uploaded
     * once instead of twice.
     */
    public async uploadToFileFieldOnly(
        entitySetName: string,
        entityId: string,
        fileFieldLogicalName: string,
        pendingFile: IPendingFile
    ): Promise<IBatchUploadResult> {
        const { IsUploaded, ErrorMessage } = await this.uploadFileToDataverseField(
            entitySetName,
            entityId,
            fileFieldLogicalName,
            pendingFile.file
        );

        if (IsUploaded) {
            // No real Annotation exists in this path - the File column
            // logical name is pushed purely as a non-empty placeholder so
            // callers' `result.succeeded.length > 0` checks still work.
            return { succeeded: [fileFieldLogicalName], failed: [] };
        }
        return { succeeded: [], failed: [{ fileName: pendingFile.fileName, error: ErrorMessage }] };
    }

    /**
     * Downloads the binary content of a Dataverse File column as base64,
     * matching retrieveAnnotationBody's return shape so index.ts can reuse
     * the same triggerBrowserDownload() code path for both sources.
     * There is no context.webAPI equivalent for this (file columns are not
     * retrievable via normal record GETs - see Microsoft's "Use file
     * column data" docs), so this uses the same raw-fetch + XrmHelper
     * client-URL pattern already used for uploads.
     */
    public async downloadFileFieldValue(entitySetName: string, entityId: string, fileFieldLogicalName: string): Promise<string> {
        const clientUrl = XrmHelper.GetClientUrl();
        const url = `${clientUrl}/api/data/v9.2/${entitySetName}(${entityId})/${fileFieldLogicalName}/$value`;

        const response = await fetch(url, {
            method: "GET",
            headers: {
                "OData-Version": "4.0",
                "OData-MaxVersion": "4.0",
                "Accept": "application/json"
            }
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Failed to download file from '${fileFieldLogicalName}'. Status: ${response.status}. ${errorText}`);
        }

        const buffer = await response.arrayBuffer();
        return AnnotationService.arrayBufferToBase64(buffer);
    }

    /**
     * Clears a Dataverse File column's value. Documented behavior: "Delete
     * File data: Set column value to null" - this uses the officially
     * supported, sandboxed context.webAPI.updateRecord rather than a raw
     * DELETE fetch, since setting the value to null is the same operation
     * and stays within the documented PCF surface.
     */
    public async deleteFileFieldValue(entityLogicalName: string, entityId: string, fileFieldLogicalName: string): Promise<void> {
        await this.webAPI.updateRecord(
            entityLogicalName,
            entityId,
            { [fileFieldLogicalName]: null } as ComponentFramework.WebApi.Entity
        );
    }

    /** Converts binary data to base64 without blowing the call stack on large files (chunked String.fromCharCode). */
    private static arrayBufferToBase64(buffer: ArrayBuffer): string {
        const bytes = new Uint8Array(buffer);
        let binary = "";
        const chunkSize = 0x8000;
        for (let i = 0; i < bytes.length; i += chunkSize) {
            binary += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + chunkSize)));
        }
        return btoa(binary);
    }

    /**
     * Retrieves existing notes for a record, deliberately EXCLUDING
     * documentbody - large base64 blobs in a list query would make the
     * list slow to load. The body is fetched on demand, only when the
     * user actually clicks "Download" on a given note.
     */
    public async retrieveAnnotations(entityId: string, sourceFieldTag?: string): Promise<IAnnotationRecord[]> {
        let filter = `_objectid_value eq ${entityId}`;
        if (sourceFieldTag) {
            // Restrict to Notes created by THIS control instance (this field)
            // so multiple FileUploadControl instances on the same form each
            // only ever see the files they themselves uploaded.
            const tag = AnnotationService.odataEscape(AnnotationService.buildSourceTag(sourceFieldTag));
            filter += ` and notetext eq '${tag}'`;
        }

        const query =
            "?$select=annotationid,filename,mimetype,filesize,subject,createdon" +
            `&$filter=${filter}` +
            "&$orderby=createdon desc";

        const result = await this.webAPI.retrieveMultipleRecords("annotation", query);

        return result.entities.map((e: any) => ({
            annotationid: e["annotationid"] as string,
            filename: e["filename"] as string,
            mimetype: e["mimetype"] as string,
            filesize: e["filesize"] as number,
            subject: e["subject"] as string | undefined,
            createdon: e["createdon"] as string | undefined
        }));
    }

    /** Lazily fetches the base64 document body for one note (used only when downloading). */
    public async retrieveAnnotationBody(annotationId: string): Promise<string> {
        const result = await this.webAPI.retrieveRecord("annotation", annotationId, "?$select=documentbody");
        return result["documentbody"] as string;
    }

    /** Deletes a note. */
    public async deleteAnnotation(annotationId: string): Promise<void> {
        await this.webAPI.deleteRecord("annotation", annotationId);
    }

    /** Safely extracts a human-readable message from a Web API error of unknown shape. */
    private static extractErrorMessage(err: unknown): string {
        if (err instanceof Error) {
            return err.message;
        }
        if (typeof err === "object" && err !== null && "message" in err) {
            return String((err as { message: unknown }).message);
        }
        return "Unknown error while communicating with Dataverse.";
    }
}
