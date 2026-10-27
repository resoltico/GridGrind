package dev.erst.gridgrind.core.plan;

/**
 * Source-materialized actions; protection secrets remain references until their operation executes.
 */
public sealed interface PreparedMutationAction
    permits ResolvedMutationAction,
        WorkbookMutationAction.SetSheetProtection,
        WorkbookMutationAction.SetWorkbookProtection {}
