import { XrmHelper } from "./XrmHelper";

export interface IFileFieldInfo {
    fileAttachmentId: string;
    fileName: string;
    fileSize: number;
    mimeType: string;
}

export class DataverseHelper {

    public static async GetFileFieldContent(entitySetName: string, fileFieldLogicalName: string, recordId?: string): Promise<{ fileName: string; mimeType: string; size: number; base64: string; } | undefined> {
        try {
            const Xrm = XrmHelper.GetXrm();
            const clientUrl = Xrm?.Utility?.getGlobalContext?.()?.GetClientUrl?.() ?? window.location.origin;
            const formContext = XrmHelper.GetFormContext();
            const entityId = recordId?.replace(/[{}]/g, "") || formContext?.data?.entity?.getId?.()?.replace(/[{}]/g, "");

            if (!entityId) {
                console.warn("DataverseHelper.GetFileFieldContent: record ID not found.");
                return undefined;
            }

            const cleanId = entityId.replace(/[{}]/g, "");

            // ================================================================
            // 1. Get file metadata
            // ================================================================
            const metadataUrl = `${clientUrl}/api/data/v9.2/${entitySetName}(${cleanId})/${fileFieldLogicalName}`;
            const metadataResponse = await fetch(metadataUrl, {
                method: "GET",
                headers: {
                    "Accept": "application/json",
                    "OData-Version": "4.0",
                    "OData-MaxVersion": "4.0"
                },
                credentials: "same-origin"
            });

            if (!metadataResponse.ok) {
                throw new Error(`Failed to retrieve file metadata (${metadataResponse.status}).`);
            }

            const metadata = await metadataResponse.json();
            if (!metadata?.value) {
                return undefined;
            }

            // ================================================================
            // 2. Get actual file content
            // ================================================================
            const contentUrl = `${clientUrl}/api/data/v9.2/${entitySetName}(${cleanId})/${fileFieldLogicalName}?size=full`;
            const contentResponse = await fetch(contentUrl, {
                method: "GET",
                headers: {
                    // "Accept": "application/octet-stream",
                    "OData-Version": "4.0",
                    "OData-MaxVersion": "4.0"
                },
                // credentials: "same-origin"
            });

            const content = await contentResponse.json();
            if (!content?.value) {
                return undefined;
            }

            // ================================================================
            // 4. Return normalized file object
            // ================================================================
            return {
                fileName: metadata.name ?? metadata.filename ?? fileFieldLogicalName,
                mimeType: metadata.mimetype ?? contentResponse.headers.get("content-type") ?? "application/octet-stream",
                size: metadata.size ?? metadata.filesizeinbytes ?? content.value.length,
                base64: content.value
            };

        } catch (err) {
            console.error("XrmHelper.getFileFieldContent failed:", err);
            return undefined;
        }
    }

    public static async GetFileFieldInfo(entityId: string, fileFieldLogicalName: string): Promise<IFileFieldInfo | undefined> {
        try {
            const cleanId = entityId.replace(/[{}]/g, "");
            const clientUrl = XrmHelper.GetClientUrl();
            const url = `${clientUrl}/api/data/v9.2/fileattachments` +
                `?$select=fileattachmentid,filename,filesizeinbytes,mimetype,regardingfieldname` +
                `&$filter=_objectid_value eq ${cleanId} and regardingfieldname eq '${fileFieldLogicalName}'`;

            const response = await fetch(url, {
                headers: {
                    "Accept": "application/json",
                    "OData-MaxVersion": "4.0",
                    "OData-Version": "4.0"
                }
            });

            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }

            const result = await response.json();
            const file = result?.value?.[0];

            if (!file) { return undefined; }

            return {
                fileAttachmentId: file.fileattachmentid,
                fileName: file.filename ?? "",
                fileSize: Number(file.filesizeinbytes ?? 0),
                mimeType: /*file.mimetype ?? */ "application/octet-stream"
            };
        } catch (err) {
            console.warn(`XrmHelper.getFileFieldInfo("${fileFieldLogicalName}") failed:`, err);
            return undefined;
        }
    }

    /**
     * Deletes the file currently stored in a Dataverse File column.
     *
     * The file is represented by a FileAttachment record.
     * We first resolve the attachment belonging to the specified
     * File column and then delete that FileAttachment.
     */
    public static async DeleteFileField(entitySetName: string, entityId: string, fileFieldLogicalName: string): Promise<boolean> {
        try {
            const clientUrl = XrmHelper.GetClientUrl();
            const cleanId = entityId.replace(/[{}]/g, "");
            const url = `${clientUrl}/api/data/v9.2/${entitySetName}(${cleanId})/${fileFieldLogicalName}`;
            const response = await fetch(url, {
                method: "DELETE",
                headers: {
                    "Accept": "application/json",
                    "OData-MaxVersion": "4.0",
                    "OData-Version": "4.0"
                },
                credentials: "same-origin"
            });

            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Failed to delete file field. HTTP ${response.status}: ${errorText}`);
            }

            return true;
        } catch (err) {
            console.error("XrmHelper.deleteFileField failed:", err);
            return false;
        }
    }
}