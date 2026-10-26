package dev.erst.gridgrind.core.model.value;

import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.formula.ResolvedRawFormula;
import dev.erst.gridgrind.core.model.style.ResolvedRichText;

/** Materialized cell-write operands; source-free scalar and temporal records are shared. */
public sealed interface ResolvedCellWrite
    permits CellValue,
        CellValueSpec.Date,
        CellValueSpec.DateTime,
        ResolvedRichText,
        ResolvedFormula,
        ResolvedRawFormula {}
