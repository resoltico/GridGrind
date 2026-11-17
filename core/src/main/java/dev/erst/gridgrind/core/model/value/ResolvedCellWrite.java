package dev.erst.gridgrind.core.model.value;

import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.formula.ResolvedRawFormula;
import dev.erst.gridgrind.core.model.style.ResolvedRichText;

/**
 * Materialized cell-write operands; validated scalar and temporal leaves share authored contracts.
 */
public sealed interface ResolvedCellWrite
    permits CellValue.Blank,
        CellValue.BooleanValue,
        CellValueSpec.NumberValue,
        CellValueSpec.ErrorValue,
        ResolvedCellText,
        CellValueSpec.Date,
        CellValueSpec.DateTime,
        ResolvedRichText,
        ResolvedFormula,
        ResolvedRawFormula {}
