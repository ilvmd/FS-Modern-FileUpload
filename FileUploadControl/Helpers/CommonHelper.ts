import { IInputs } from "../generated/ManifestTypes";
import { ValidationService } from "../services/ValidationService";
import { DisplayMode, IControlConfig } from "../types/Interfaces";
import { XrmHelper } from "./XrmHelper";

export class CommonHelper {
    /** Forces a section/tab refresh style redraw of a single control, e.g. after programmatically changing its value. */
    public static RefreshField(sFieldLogicalName: string): boolean {
        try {
            const control = CommonHelper.GetControl(sFieldLogicalName);
            if (!control || typeof control.refresh !== "function") {
                return false;
            }
            control.refresh();
            return true;
        } catch {
            return false;
        }
    }

    // =========================================================================
    // Control / field visibility (Hide/Show)
    // =========================================================================

    /** Looks up a form control by its bound field's logical name. */
    public static GetControl(sFieldLogicalName: string): any | undefined {
        try {
            const objFormContext = XrmHelper.GetFormContext();
            return objFormContext?.getControl?.(sFieldLogicalName) ?? undefined;
        } catch (err) {
            console.warn(`XrmHelper.getControl("${sFieldLogicalName}") failed:`, err);
            return undefined;
        }
    }

    // =========================================================================
    // Attribute get/set
    // =========================================================================
    /** Reads the current value of any attribute on the form (GetValue). */
    public static GetValue<T = unknown>(sFieldLogicalName: string): T | undefined {
        try {
            const objFormContext = XrmHelper.GetFormContext();
            const attribute = objFormContext?.getAttribute?.(sFieldLogicalName);
            return CommonHelper.IsValid(attribute) ? (attribute.getValue() as T) : undefined;
        } catch (err) {
            console.warn(`XrmHelper.getValue("${sFieldLogicalName}") failed:`, err);
            return undefined;
        }
    }

    /** Writes a value to any attribute on the form (SetValue). Optionally fires onChange handlers. */
    public static SetValue(sFieldLogicalName: string, value: unknown, fireOnChange = false): boolean {
        try {
            const objFormContext = XrmHelper.GetFormContext();
            const attribute = objFormContext?.getAttribute?.(sFieldLogicalName);
            if (!CommonHelper.IsValid(attribute)) {
                return false;
            }
            attribute.setValue(value);
            if (fireOnChange) {
                attribute.fireOnChange();
            }
            return true;
        } catch (err) {
            console.warn(`XrmHelper.setValue("${sFieldLogicalName}") failed:`, err);
            return false;
        }
    }

    /**
     * Function Call Format: CommonHelper.HideShowTab
     * */
    public static HideShowTab(sTabName: string, bShowTab: boolean) {
        const objFormContext = XrmHelper.GetFormContext();
        var objTab = null;
        try {
            objTab = objFormContext.ui.tabs.get(sTabName);
            if (CommonHelper.IsValid(objTab)) {
                objTab.setVisible(bShowTab);
            }
        }
        catch (ex: any) {
            CommonHelper.ShowError(ex.message);
        }
    }

    /**
     * Function Call Format: CommonHelper.HideShowSection
     * */
    public static HideShowSection(sTabName: string, sSectionName: string, bShowSection: boolean) {
        const objFormContext = XrmHelper.GetFormContext();
        var objTab = null;
        var objSection = null;
        try {
            objTab = objFormContext.ui.tabs.get(sTabName);
            if (CommonHelper.IsValid(objTab)) {
                objSection = objTab.sections.get(sSectionName);
                if (CommonHelper.IsValid(objSection)) {
                    objSection.setVisible(bShowSection);
                }
            }
        }
        catch (ex: any) {
            CommonHelper.ShowError(ex.message);
        }
    }

    /**
     * Function Call Format: CommonHelper.GetDisabled
     * */
    public static GetDisabled(sFieldName: string) {
        const objFormContext = XrmHelper.GetFormContext();
        var bIsDisabled = false;
        try {
            if (CommonHelper.IsValid(objFormContext.getControl(sFieldName))) {
                bIsDisabled = objFormContext.getControl(sFieldName).getDisabled();
            }
            return bIsDisabled;
        }
        catch (ex: any) {
            CommonHelper.ShowError(ex.message);
        }
    }

    /**
     * Function Call Format: CommonHelper.SetDisabled
     * */
    public static SetDisabled(arrFieldName: Array<string>, bDisableControl: boolean) {
        const objFormContext = XrmHelper.GetFormContext();
        try {
            if (CommonHelper.IsValid(arrFieldName)) {
                for (var i = 0; i < arrFieldName.length; i++) {
                    if (CommonHelper.IsValid(objFormContext.getAttribute(arrFieldName[i]))) {
                        objFormContext.getAttribute(arrFieldName[i]).controls.forEach(
                            function (objControl: any) {
                                objControl.setDisabled(bDisableControl);
                            }
                        );
                    }

                    let objControl = objFormContext.getControl(`header_process_${arrFieldName[i]}`);
                    if (CommonHelper.IsValid(objControl)) {
                        objControl.setDisabled(bDisableControl);
                    }
                }
            }
        }
        catch (ex: any) {
            CommonHelper.ShowError(ex.message);
        }
    }

    /**
     * Function Call Format: CommonHelper.GetVisible
     * */
    public static GetVisible(sFieldName: string) {
        const objFormContext = XrmHelper.GetFormContext();
        var bIsVisible = false;
        try {
            if (CommonHelper.IsValid(objFormContext.getControl(sFieldName))) {
                bIsVisible = objFormContext.getControl(sFieldName).getVisible();
            }
            return bIsVisible;
        }
        catch (ex: any) {
            CommonHelper.ShowError(ex.message);
        }
    }

    /**
     * Function Call Format: CommonHelper.SetVisible
     * */
    public static SetVisible(arrFieldName: Array<string>, bShowControl: boolean) {
        const objFormContext = XrmHelper.GetFormContext();
        try {
            for (var i = 0; i < arrFieldName.length; i++) {
                if (CommonHelper.IsValid(objFormContext.getAttribute(arrFieldName[i]))) {
                    objFormContext.getAttribute(arrFieldName[i]).controls.forEach(
                        function (objControl: any) {
                            objControl.setVisible(bShowControl);
                        }
                    );
                }

                let objControl = objFormContext.getControl(`header_process_${arrFieldName[i]}`);
                if (CommonHelper.IsValid(objControl)) {
                    objControl.setVisible(bShowControl);
                }
            }
        }
        catch (ex: any) {
            CommonHelper.ShowError(ex.message);
        }
    }


    /**
     * Function Call Format: CommonHelper.Clear
     * */
    public static Clear(arrFieldName: Array<string>) {
        const objFormContext = XrmHelper.GetFormContext();
        try {
            arrFieldName.forEach(sFieldName => {
                if (CommonHelper.IsValid(objFormContext.getAttribute(sFieldName))) {
                    objFormContext.getAttribute(sFieldName).setValue(null);
                }
            })
        }
        catch (ex: any) {
            CommonHelper.ShowError("Error in [Clear] :: " + ex.message);
        }
    }

    /**
     * Function Call Format: CommonHelper.SetRequiredLevel
     * */
    public static SetRequiredLevel(arrFieldName: Array<string>, sRequirementLevel: 'required' | 'none' | 'recommended') {
        const objFormContext = XrmHelper.GetFormContext();
        try {
            for (var i = 0; i < arrFieldName.length; i++) {
                if (CommonHelper.IsValid(objFormContext.getAttribute(arrFieldName[i]))) {
                    objFormContext.getAttribute(arrFieldName[i]).setRequiredLevel(sRequirementLevel);
                }
            }
        }
        catch (ex: any) {
            CommonHelper.ShowError(ex.message);
        }
    }

    // =========================================================================
    // Record / entity context
    // =========================================================================
    /**
     * Function Call Format: CommonHelper.GetEntityLogicalName
     * */
    public static GetEntityLogicalName() {
        const objFormContext = XrmHelper.GetFormContext();
        var sEntityLogicalName = "";
        try {
            sEntityLogicalName = objFormContext.data.entity.getEntityName();
        }
        catch (ex: any) {
            CommonHelper.ShowError(ex.message);
        }
        return sEntityLogicalName;
    }

    /**
     * Function Call Format: CommonHelper.GetRecordID
     * */
    public static GetRecordID() {
        const objFormContext = XrmHelper.GetFormContext();
        var sRecordID = "";
        try {
            sRecordID = CommonHelper.ValidateAndFormatGuid(objFormContext.data.entity.getId()) ?? '';
        }
        catch (ex: any) {
            CommonHelper.ShowError(ex.message);
        }
        return sRecordID;
    }

    /**
     * Function Call Format: CommonHelper.GetLoggedInUserID
     * */
    public static GetLoggedInUserID() {
        const Xrm = XrmHelper.GetXrm();
        const objFormContext = XrmHelper.GetFormContext();
        var sCurrentUserID = "";
        try {
            sCurrentUserID = CommonHelper.ValidateAndFormatGuid(Xrm.Utility.getGlobalContext().userSettings.userId) ?? '';
        }
        catch (ex: any) {
            CommonHelper.ShowError(ex.message);
        }
        return sCurrentUserID;
    }

    /**
     * Function Call Format: CommonHelper.GetLoggedInUserLanguageId
     * */
    public static GetLoggedInUserLanguageId() {
        const Xrm = XrmHelper.GetXrm();
        const objFormContext = XrmHelper.GetFormContext();
        var iUserLanguageCode = 0;
        try {
            iUserLanguageCode = Xrm.Utility.getGlobalContext().userSettings.languageId;
        }
        catch (ex: any) {
            CommonHelper.ShowError(ex.message);
        }
        return iUserLanguageCode;
    }

    /**
     * Function Call Format: CommonHelper.GetAllForms
     * */
    public static GetAllForms() {
        const objFormContext = XrmHelper.GetFormContext();
        try {
            return objFormContext?.ui?.formSelector?.items?.get() ?? [];
        }
        catch (ex: any) {
            CommonHelper.ShowError(ex.message);
        }
    }

    /**
     * Function Call Format: CommonHelper.GetFormName
     * */
    public static GetFormName() {
        const objFormContext = XrmHelper.GetFormContext();
        var sCurrentFormName = null;
        var objFormItem = null;
        try {
            objFormItem = objFormContext.ui.formSelector.getCurrentItem();
            if (CommonHelper.IsValid(objFormItem)) {
                sCurrentFormName = objFormItem.getLabel().toString().toLowerCase().trim();
            }
        }
        catch (ex: any) {
            CommonHelper.ShowError(ex.message);
        }
        return sCurrentFormName;
    }

    /**
     * Function Call Format: CommonHelper.GetFormID
     * */
    public static GetFormID() {
        const objFormContext = XrmHelper.GetFormContext();
        var sFormId = "";
        try {
            sFormId = CommonHelper.ValidateAndFormatGuid(objFormContext.ui.formSelector.getCurrentItem().getId()) ?? '';
        }
        catch (ex: any) {
            CommonHelper.ShowError(ex.message);
        }
        return sFormId;
    }

    /**
 * Function Call Format: CommonHelper.FormIsDirty
 * */
    public static FormIsDirty() {
        const objFormContext = XrmHelper.GetFormContext();
        return objFormContext.data.entity.getIsDirty();
    }

    /**
     * Function Call Format: CommonHelper.GetDirtyFields
     * */
    public static GetDirtyFields() {
        const objFormContext = XrmHelper.GetFormContext();
        let arrFields = objFormContext.data.entity.attributes;
        return arrFields.filter((objField: any) => objField.getIsDirty())
            .map((objField: any) => ({ name: objField.getName(), value: objField.getValue() }));
    }

    /**
     * Function Call Format: CommonHelper.IsCreatedForm
     * */
    public static IsCreatedForm() {
        const objFormContext = XrmHelper.GetFormContext();
        return objFormContext.ui.getFormType() === 1;
    }

    /**
     * Function Call Format: CommonHelper.IsUpdatedForm
     * */
    public static IsUpdatedForm() {
        const objFormContext = XrmHelper.GetFormContext();
        return objFormContext.ui.getFormType() === 2;
    }

    /**
     * Function Call Format: CommonHelper.IsReadOnlyForm
     * */
    public static IsReadOnlyForm() {
        const objFormContext = XrmHelper.GetFormContext();
        return objFormContext.ui.getFormType() === 3;
    }

    /**
     * Function Call Format: CommonHelper.IsValid
     * */
    public static IsValid(object: any) {
        const objFormContext = XrmHelper.GetFormContext();
        try {
            if (object != null && object !== "undefined" && object !== undefined && object !== "") {
                return true;
            }
            else {
                return false;
            }
        }
        catch (ex: any) {
            CommonHelper.ShowError("Error in function IsValid :: " + ex.message);
        }
    }

    /**
     * Function Call Format: CommonHelper.ShowError
     * */
    public static ShowError(sErrMessage: string) {
        const Xrm = XrmHelper.GetXrm();
        const objFormContext = XrmHelper.GetFormContext();
        var objAlertStrings = null;
        var objAlertOptions = null;
        objAlertStrings = { confirmButtonLabel: "OK", text: sErrMessage };
        objAlertOptions = { height: 120, width: 260 };
        Xrm.Navigation.openAlertDialog(objAlertStrings, objAlertOptions).then(
            function success(result: any) { },
            function (ex: Error) { console.error(ex.message); }
        );
    }

    // =========================================================================
    // Dialogs
    // =========================================================================
    /**
     * Native Dataverse confirm dialog (Xrm.Navigation.openConfirmDialog),
     * with a graceful fallback to window.confirm when Xrm/Navigation is not
     * reachable (e.g. the standalone test harness). Always resolves - never
     * rejects - so callers can `await` it unconditionally.
     */
    public static async Confirm(title: string, text: string, confirmLabel = "OK", cancelLabel = "Cancel"): Promise<boolean> {
        const Xrm = XrmHelper.GetXrm();
        try {
            if (Xrm?.Navigation?.openConfirmDialog) {
                const result = await Xrm.Navigation.openConfirmDialog(
                    { title, text },
                    { confirmButtonLabel: confirmLabel, cancelButtonLabel: cancelLabel }
                );
                return !!result?.confirmed;
            }
        } catch (err) {
            console.warn("XrmHelper.confirm via Xrm.Navigation failed, falling back to window.confirm:", err);
        }
        // eslint-disable-next-line no-alert
        return window.confirm(`${title}\n\n${text}`);
    }

    /**
     * Function Call Format: LESHA.XRM.WebResource.Common.ValidateAndFormatGuid
     * This function Validates the GUID passed to it and returns the same with removing Braces
     * */
    public static ValidateAndFormatGuid(sLookUpId: string) {
        var sLookUpIdToReturn = null;
        try {
            sLookUpIdToReturn = CommonHelper.IsValid(sLookUpId)
                ? sLookUpId.replace(/[{}]/g, "")
                : sLookUpId;

            return sLookUpIdToReturn;
        }
        catch (ex: any) {
            CommonHelper.ShowError(ex.message);
        }
    }

    // // =========================================================================
    // // Small shared helpers
    // // =========================================================================

    /**
     * Determines a stable identifier for "which field is THIS control
     * instance on", used to tag/filter Notes so multiple instances on the
     * same form don't show each other's files. Tries, in order:
     *
     *  1. currentFieldLogicalName (manifest input) - an explicit override
     *     the maker can type into the control's configuration in the form
     *     designer. Most predictable option; use this if you want full
     *     control over the tag (e.g. to intentionally share files between
     *     two controls by giving them the same value).
     *  2. context.parameters.value.attributes.LogicalName - the logical
     *     name of the Dataverse column actually bound to this control's
     *     "value" property, read automatically from platform metadata.
     *     Since every instance MUST be bound to its own dedicated Multiline
     *     Text state field (see class doc), this is unique per instance
     *     with zero manual configuration.
     *  3. documentFileFieldLogicalName (manifest input) - last resort for
     *     the single-file-per-field mode, in case (1) and (2) are both
     *     unavailable.
     *
     * Returns undefined only if none of the above resolved - in that case
     * the control intentionally falls back to legacy unfiltered behavior
     * (shows all Notes on the record) rather than silently hiding files.
     */
    public static ResolveSourceFieldTag(context: ComponentFramework.Context<IInputs>): string | undefined {
        // const manualOverride = context.parameters.currentFieldLogicalName.raw?.trim();
        // if (manualOverride) {
        //     return manualOverride;
        // }

        const boundFieldName = (context.parameters.value as unknown as { attributes?: { LogicalName?: string }; }).attributes?.LogicalName;
        if (boundFieldName) {
            return boundFieldName;
        }

        const fileFieldName = context.parameters.documentFileFieldLogicalName.raw?.trim();
        if (fileFieldName) {
            return fileFieldName;
        }

        return undefined;
    }

    
    /** Reads and normalizes manifest input properties into a strongly typed config object. */
    public static ReadConfig(context: ComponentFramework.Context<IInputs>): IControlConfig {
        const rawDisplayMode = context.parameters.displayMode.raw ?? "0";
        const displayMode: DisplayMode = rawDisplayMode === "1" ? "Button" : "Dropzone";

        return {
            documentFileFieldLogicalName: context.parameters.documentFileFieldLogicalName.raw ?? '',
            allowMultipleFiles: context.parameters.allowMultipleFiles.raw === true,
            allowedExtensions: ValidationService.parseCsvList(context.parameters.allowedExtensions.raw),
            allowedMimeTypes: ValidationService.parseCsvList(context.parameters.allowedMimeTypes.raw),
            maxFileSizeKB: context.parameters.maxFileSizeKB.raw ?? 10240,
            maxFileCount: context.parameters.maxFileCount.raw ?? 1,
            // currentFieldLogicalName: context.parameters.currentFieldLogicalName.raw ?? '',

            displayMode,
            buttonIdleLabel: context.parameters.buttonIdleLabel.raw?.trim() || "Upload File",

            hideControlAfterUpload: context.parameters.hideControlAfterUpload.raw === true,
            showDocumentFieldAfterUpload: context.parameters.showDocumentFieldAfterUpload.raw === true,
            deleteNoteAfterFileFieldUpload: context.parameters.deleteNoteAfterFileFieldUpload.raw === true,
            // enableXrmIntegration: context.parameters.enableXrmIntegration.raw === true,

            allowDownload: context.parameters.allowDownload.raw === true,
            allowDelete: context.parameters.allowDelete.raw === true,
            confirmBeforeDelete: context.parameters.confirmBeforeDelete.raw === true,
            showPreviewThumbnails: context.parameters.showPreviewThumbnails.raw === true,
            showFileMeta: context.parameters.showFileMeta.raw === true,
            // readOnly: context.parameters.readOnly.raw === true,
            readOnly: CommonHelper.GetDisabled(CommonHelper.ResolveSourceFieldTag(context) ?? '') ?? false,
            accentColor: context.parameters.accentColor.raw?.trim() || "#0078d4",
            labalColor: context.parameters.labalColor.raw?.trim() || "#000000",
            replacePCFWithSelectedFiles: context.parameters.replacePCFWithSelectedFiles.raw === true,
            allowPreview: context.parameters.allowPreview.raw === true
        };
    }

    /**
     * Normalizes the platform-supplied entity id: treats null/undefined/
     * empty-guid as "no record yet" so the create->saved transition can be
     * detected reliably (some hosts report an all-zero GUID rather than
     * omitting the value entirely on a brand-new Create form).
     */
    public static  NormalizeGuid(id: string | null | undefined): string | undefined {
        if (!id) {
            return undefined;
        }
        const cleaned = id.replace(/[{}]/g, "");
        if (cleaned === "00000000-0000-0000-0000-000000000000") {
            return undefined;
        }
        return cleaned;
    }

    public static  DescribeError(err: unknown): string {
        if (err instanceof Error) {
            return err.message;
        }
        if (typeof err === "object" && err !== null && "message" in err) {
            return String((err as { message: unknown }).message);
        }
        return "Unknown error";
    }
}