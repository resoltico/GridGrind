package dev.erst.gridgrind.contract.step;

/** One statically quantified worksheet-work estimate or a standalone shape violation. */
sealed interface WorkbookStaticWorkEstimate
    permits ValidWorkbookStaticWorkEstimate, InvalidWorkbookStaticWorkEstimate {}

/** One valid worksheet-work count and the JSON path that owns it. */
record ValidWorkbookStaticWorkEstimate(long workItems, String jsonPath)
    implements WorkbookStaticWorkEstimate {}

/** One independently provable worksheet-work request violation. */
record InvalidWorkbookStaticWorkEstimate(WorkbookStaticViolation violation)
    implements WorkbookStaticWorkEstimate {}
