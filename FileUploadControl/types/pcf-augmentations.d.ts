/**
 * pcf-augmentations.d.ts
 * -----------------------
 * The published `@types/powerapps-component-framework` package does not
 * currently declare `contextInfo` on `ComponentFramework.Mode`, even
 * though it is a real, Microsoft-documented runtime property available to
 * every PCF control hosted on a model-driven app entity form (see
 * "Component Framework API context.mode.contextInfo" on Microsoft Learn).
 *
 * Rather than scattering `as any` casts through index.ts, we restore the
 * missing member here via TypeScript declaration merging. If/when the
 * upstream type package adds this member, this file becomes a harmless
 * duplicate and can be deleted.
 */
declare namespace ComponentFramework {
    interface Mode {
        /**
         * Present when the control is hosted on a model-driven app entity
         * form. `entityId` is empty/undefined until the record has been
         * saved for the first time; the control relies on this transition
         * to detect that a Create form was just saved.
         */
        contextInfo?: {
            entityTypeName: string;
            entityId: string;
            entityRecordName: string;
        };
    }
}
