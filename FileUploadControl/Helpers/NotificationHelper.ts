import { IInputs } from "../generated/ManifestTypes";
import { XrmHelper } from "./XrmHelper";

type NotificationLevel = "ERROR" | "WARNING" | "INFO";
export type XrmNotificationLevel = "ERROR" | "WARNING" | "INFO";

interface UtilityWithNotifications {
    setNotification?: (message: string, uniqueId: string) => boolean;
    clearNotification?: (uniqueId: string) => void;
}

/**
 * Thin wrapper around context.utils.setNotification / clearNotification.
 * Avoids notification spam by only calling setNotification when the
 * message actually changes, and tracks a single notification id so
 * repeated errors update in place rather than stacking.
 * ---------------------------------------------------------------------
 * Thin wrapper around context.utils.setNotification / clearNotification -
 * the officially documented, sandbox-safe way to surface a message next to
 * this control. Avoids notification spam by only calling setNotification
 * when the message actually changes, and tracks a single notification id
 * so repeated errors update in place rather than stacking.
 *
 * Falls back to the Global-Xrm form notification bar (via XrmHelper) only
 * when `context.utils.setNotification` is unavailable on the current host,
 * so a message is still surfaced somewhere the user will see it.
 */
export class NotificationHelper {
    private static readonly NOTIFICATION_ID = "fileUploadControlNotification";
    private lastMessage: string | undefined;

    constructor(private context: ComponentFramework.Context<IInputs>) { }

    public showError(message: string): void {
        this.show(message, "ERROR");
    }

    public showWarning(message: string): void {
        this.show(message, "WARNING");
    }

    public showInfo(message: string): void {
        this.show(message, "INFO");
    }

    private show(message: string, level: "ERROR" | "WARNING" | "INFO"): void {
        console.log("SHOW NOTIFICATION");
        if (message === this.lastMessage) { return; }
        this.lastMessage = message;
        try {
            const utils = this.context.utils as unknown as UtilityWithNotifications;
            utils.setNotification?.(message, NotificationHelper.NOTIFICATION_ID);
        } catch (ex: any) {}

        if (level === "ERROR") {}
    }

    public clear(): void {
        console.log("CLEAR NOTIFICATION");
        this.lastMessage = undefined;
        try {
            const utils = this.context.utils as unknown as UtilityWithNotifications;
            utils.clearNotification?.(NotificationHelper.NOTIFICATION_ID);
        } catch (error) {
        }
    }

    private ShowNotification(message: string, level: NotificationLevel): void {
        if (message === this.lastMessage) {
            return;
        }
        this.lastMessage = message;

        let handled = false;
        try {
            const utils = this.context.utils as unknown as UtilityWithNotifications;
            handled = !!utils.setNotification?.(message, NotificationHelper.NOTIFICATION_ID);
        } catch (err) {
            console.warn("NotificationHelper: context.utils.setNotification failed:", err);
        }

        if (!handled) {
            this.SetFormNotification(message, level, NotificationHelper.NOTIFICATION_ID);
        }
    }

    public ClearNotification(): void {
        this.lastMessage = undefined;
        try {
            const utils = this.context.utils as unknown as UtilityWithNotifications;
            utils.clearNotification?.(NotificationHelper.NOTIFICATION_ID);
        } catch (err) {
            console.warn("NotificationHelper: context.utils.clearNotification failed:", err);
        }
        this.ClearFormNotification(NotificationHelper.NOTIFICATION_ID);
    }

    
    // =========================================================================
    // Notifications (form-level notification bar)
    // =========================================================================

    public SetFormNotification(message: string, level: XrmNotificationLevel, uniqueId: string): boolean {
        try {
            const objFormContext = XrmHelper.GetFormContext();
            if (!objFormContext?.ui?.setFormNotification) {
                return false;
            }
            objFormContext.ui.setFormNotification(message, level, uniqueId);
            return true;
        } catch (err) {
            console.warn("XrmHelper.setFormNotification failed:", err);
            return false;
        }
    }

    public ClearFormNotification(uniqueId: string): boolean {
        try {
            const objFormContext = XrmHelper.GetFormContext();
            if (!objFormContext?.ui?.clearFormNotification) {
                return false;
            }
            objFormContext.ui.clearFormNotification(uniqueId);
            return true;
        } catch {
            return false;
        }
    }

}
