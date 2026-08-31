/**
 * XrmHelper
 * =========
 * Centralized, defensive wrapper around the **Global Xrm object** (the
 * classic `Xrm` API available on `window`/`window.parent`/`window.top` in
 * every model-driven app form) plus the client-side form context it
 * exposes (`Xrm.Page`, still the most broadly compatible way to reach the
 * *current* form context from inside a child frame such as a PCF control).
 *
 * WHY THIS EXISTS
 * ----------------
 * `context.webAPI` and `context.utils` (the officially documented PCF
 * surface) cover data access and a couple of utility calls, but they do
 * NOT expose form-level capabilities such as showing/hiding a sibling
 * field, reading/writing another attribute's value, or driving the native
 * "Save blocked" notification bar. Those capabilities only exist on the
 * Global Xrm object / form context. Every place in this control that needs
 * that kind of capability goes through THIS file, so:
 *   - there is a single, well-documented place that knows how to reach Xrm,
 *   - every call is wrapped in try/catch and fails soft (never throws back
 *     into the control's core upload/save flow), and
 *   - the whole feature can be switched off centrally via
 *     `IControlConfig.enableXrmIntegration` for environments where reaching
 *     outside the PCF sandbox is undesirable/unavailable.
 *
 * IMPORTANT CAVEATS
 * -----------------
 *  - Reaching `window.parent`/`window.top` only works when the control is
 *    NOT cross-origin from the model-driven app frame. This is true for
 *    every standard on-premises/online model-driven app hosting scenario,
 *    but will safely no-op (never throw) inside e.g. a fully sandboxed
 *    test harness (`npm start`) where no such frame exists.
 *  - `Xrm.Page` is a deprecated alias for the *currently active* form's
 *    context. Microsoft has not removed it and it remains the most
 *    reliable zero-configuration way for a field-level control to reach
 *    "its own form" without the maker having to pass anything extra in.
 *    If Microsoft removes it in a future release, only this file needs to
 *    change.
 */

import { IInputs } from "../generated/ManifestTypes";

// import { IInputs } from "../generated/ManifestTypes";

export class XrmHelper {
    /**
     * Walks window -> window.parent -> window.top looking for a live Xrm
     * object. Every hop is guarded individually because reading
     * `window.top` (or `.Xrm` off of it) can itself throw a
     * SecurityError/DOMException in a genuinely cross-origin embed - a
     * single try/catch around the whole chain would let an early throw
     * hide the fact that a later hop might have worked.
     */
    public static GetXrm(): any | undefined {
        const candidates: Array<() => any> = [
            () => (window as any).Xrm,
            () => (window.parent as any)?.Xrm,
            () => (window.top as any)?.Xrm
        ];

        for (const read of candidates) {
            try {
                const Xrm = read();
                if (Xrm) {
                    return Xrm;
                }
            } catch {
                // Cross-origin or otherwise unreachable frame - try the next one.
            }
        }
        return undefined;
    }

    /** True as soon as any reachable Xrm object is found (does not guarantee a form context exists). */
    public static IsAvailable(): boolean {
        return !!XrmHelper.GetXrm();
    }

    /**
     * Resolves the *current* form context. Prefers the modern
     * `formContext` pattern (in case a host application stashes one on the
     * Xrm object for us), falling back to the still-fully-functional
     * `Xrm.Page` alias used across the vast majority of real deployments.
     */
    public static GetFormContext(): any | undefined {
        const Xrm = XrmHelper.GetXrm();
        if (!Xrm) {
            return undefined;
        }
        try {
            if (Xrm.Page && Xrm.Page.data && Xrm.Page.ui) {
                return Xrm.Page;
            }
        } catch {
            /* fall through */
        }
        return undefined;
    }

    /**
     * Documented, Microsoft-recommended way of resolving the environment's
     * base URL (used to build the raw fetch() URL for uploading directly
     * into a Dataverse File column, which has no context.webAPI equivalent
     * today). Falls back to window.location.origin when Xrm is unreachable
     * (e.g. inside the standalone `npm start` test harness).
     */
    public static GetClientUrl(): string {
        try {
            const Xrm = XrmHelper.GetXrm();
            const url = Xrm?.Utility?.getGlobalContext?.()?.GetClientUrl?.();
            if (url) {
                return url as string;
            }
        } catch {
            /* fall through to the safe default below */
        }
        return window.location.origin;
    }
}