package dev.erst.gridgrind.core.plan;

/** Executable mutations: source-free actions are shared with authored plans. */
public sealed interface ResolvedMutationAction
    permits CellMutationAction.ApplyStyle,
        CellMutationAction.ClearArrayFormula,
        CellMutationAction.ClearComment,
        CellMutationAction.ClearHyperlink,
        CellMutationAction.ClearRange,
        CellMutationAction.SetHyperlink,
        DrawingMutationAction.DeleteDrawingObject,
        DrawingMutationAction.SetDrawingObjectAnchor,
        StructuredMutationAction.ClearAutofilter,
        StructuredMutationAction.ClearConditionalFormatting,
        StructuredMutationAction.ClearDataValidations,
        StructuredMutationAction.DeleteNamedRange,
        StructuredMutationAction.DeletePivotTable,
        StructuredMutationAction.DeleteTable,
        StructuredMutationAction.SetAutofilter,
        StructuredMutationAction.SetConditionalFormatting,
        StructuredMutationAction.SetNamedRange,
        StructuredMutationAction.SetPivotTable,
        WorkbookMutationAction.AutoSizeColumns,
        WorkbookMutationAction.ClearPrintLayout,
        WorkbookMutationAction.ClearSheetProtection,
        WorkbookMutationAction.ClearWorkbookProtection,
        WorkbookMutationAction.CopySheet,
        WorkbookMutationAction.DeleteColumns,
        WorkbookMutationAction.DeleteRows,
        WorkbookMutationAction.DeleteSheet,
        WorkbookMutationAction.EnsureSheet,
        WorkbookMutationAction.GroupColumns,
        WorkbookMutationAction.GroupRows,
        WorkbookMutationAction.InsertColumns,
        WorkbookMutationAction.InsertRows,
        WorkbookMutationAction.MergeCells,
        WorkbookMutationAction.MoveSheet,
        WorkbookMutationAction.RenameSheet,
        WorkbookMutationAction.SetActiveSheet,
        WorkbookMutationAction.SetColumnVisibility,
        WorkbookMutationAction.SetColumnWidth,
        WorkbookMutationAction.SetRowHeight,
        WorkbookMutationAction.SetRowVisibility,
        WorkbookMutationAction.SetSelectedSheets,
        WorkbookMutationAction.SetSheetPane,
        WorkbookMutationAction.SetSheetPresentation,
        WorkbookMutationAction.SetSheetVisibility,
        WorkbookMutationAction.SetSheetZoom,
        WorkbookMutationAction.ShiftColumns,
        WorkbookMutationAction.ShiftRows,
        WorkbookMutationAction.UngroupColumns,
        WorkbookMutationAction.UngroupRows,
        WorkbookMutationAction.UnmergeCells,
        ResolvedMutationAction.SetCell,
        ResolvedMutationAction.SetRange,
        ResolvedMutationAction.AppendRow,
        ResolvedMutationAction.SetArrayFormula,
        ResolvedMutationAction.SetComment,
        ResolvedMutationAction.SetChart,
        ResolvedMutationAction.SetPicture,
        ResolvedMutationAction.SetEmbeddedObject,
        ResolvedMutationAction.SetShape,
        ResolvedMutationAction.SetSignatureLine,
        ResolvedMutationAction.ImportCustomXmlMapping,
        ResolvedMutationAction.SetDataValidation,
        ResolvedMutationAction.SetTable,
        ResolvedMutationAction.SetPrintLayout,
        ResolvedMutationAction.SetSheetProtection,
        ResolvedMutationAction.SetWorkbookProtection {

  /** Returns the authored metadata owner for diagnostics and backend routing. */
  default Class<? extends MutationAction> authoredType() {
    return switch (this) {
      case MutationAction shared -> shared.getClass().asSubclass(MutationAction.class);
      case AppendRow _ -> CellMutationAction.AppendRow.class;
      case SetArrayFormula _ -> CellMutationAction.SetArrayFormula.class;
      case SetCell _ -> CellMutationAction.SetCell.class;
      case SetComment _ -> CellMutationAction.SetComment.class;
      case SetRange _ -> CellMutationAction.SetRange.class;
      case SetChart _ -> DrawingMutationAction.SetChart.class;
      case SetEmbeddedObject _ -> DrawingMutationAction.SetEmbeddedObject.class;
      case SetPicture _ -> DrawingMutationAction.SetPicture.class;
      case SetShape _ -> DrawingMutationAction.SetShape.class;
      case SetSignatureLine _ -> DrawingMutationAction.SetSignatureLine.class;
      case ImportCustomXmlMapping _ -> StructuredMutationAction.ImportCustomXmlMapping.class;
      case SetDataValidation _ -> StructuredMutationAction.SetDataValidation.class;
      case SetTable _ -> StructuredMutationAction.SetTable.class;
      case SetPrintLayout _ -> WorkbookMutationAction.SetPrintLayout.class;
      case SetSheetProtection _ -> WorkbookMutationAction.SetSheetProtection.class;
      case SetWorkbookProtection _ -> WorkbookMutationAction.SetWorkbookProtection.class;
    };
  }

  /** Materialized operands for SetCell. */
  record SetCell(dev.erst.gridgrind.core.model.value.ResolvedCellWrite value)
      implements ResolvedMutationAction {
    public SetCell {
      java.util.Objects.requireNonNull(value, "value must not be null");
    }
  }

  /** Materialized operands for SetRange. */
  record SetRange(
      java.util.List<java.util.List<dev.erst.gridgrind.core.model.value.ResolvedCellWrite>> rows)
      implements ResolvedMutationAction {
    public SetRange {
      java.util.Objects.requireNonNull(rows, "rows must not be null");
      rows = rows.stream().map(java.util.List::copyOf).toList();
      int width = rows.isEmpty() ? 0 : rows.getFirst().size();
      if (width == 0 || rows.stream().anyMatch(row -> row.size() != width))
        throw new IllegalArgumentException("rows must describe a nonempty rectangular matrix");
    }
  }

  /** Materialized operands for AppendRow. */
  record AppendRow(java.util.List<dev.erst.gridgrind.core.model.value.ResolvedCellWrite> values)
      implements ResolvedMutationAction {
    public AppendRow {
      java.util.Objects.requireNonNull(values, "values must not be null");
      values = java.util.List.copyOf(values);
      if (values.isEmpty()) throw new IllegalArgumentException("values must not be empty");
    }
  }

  /** Materialized operands for SetArrayFormula. */
  record SetArrayFormula(dev.erst.gridgrind.core.model.formula.ResolvedArrayFormula formula)
      implements ResolvedMutationAction {
    public SetArrayFormula {
      java.util.Objects.requireNonNull(formula, "formula must not be null");
    }
  }

  /** Materialized operands for SetComment. */
  record SetComment(dev.erst.gridgrind.core.model.drawing.ResolvedComment comment)
      implements ResolvedMutationAction {
    public SetComment {
      java.util.Objects.requireNonNull(comment, "comment must not be null");
    }
  }

  /** Materialized operands for SetChart. */
  record SetChart(dev.erst.gridgrind.core.model.chart.ResolvedChart chart)
      implements ResolvedMutationAction {
    public SetChart {
      java.util.Objects.requireNonNull(chart, "chart must not be null");
    }
  }

  /** Materialized operands for SetPicture. */
  record SetPicture(dev.erst.gridgrind.core.model.drawing.ResolvedPicture picture)
      implements ResolvedMutationAction {
    public SetPicture {
      java.util.Objects.requireNonNull(picture, "picture must not be null");
    }
  }

  /** Materialized operands for SetEmbeddedObject. */
  record SetEmbeddedObject(
      dev.erst.gridgrind.core.model.drawing.ResolvedEmbeddedObject embeddedObject)
      implements ResolvedMutationAction {
    public SetEmbeddedObject {
      java.util.Objects.requireNonNull(embeddedObject, "embeddedObject must not be null");
    }
  }

  /** Materialized operands for SetShape. */
  record SetShape(dev.erst.gridgrind.core.model.drawing.ResolvedShape shape)
      implements ResolvedMutationAction {
    public SetShape {
      java.util.Objects.requireNonNull(shape, "shape must not be null");
    }
  }

  /** Materialized operands for SetSignatureLine. */
  record SetSignatureLine(dev.erst.gridgrind.core.model.drawing.ResolvedSignatureLine signatureLine)
      implements ResolvedMutationAction {
    public SetSignatureLine {
      java.util.Objects.requireNonNull(signatureLine, "signatureLine must not be null");
    }
  }

  /** Materialized operands for ImportCustomXmlMapping. */
  record ImportCustomXmlMapping(
      dev.erst.gridgrind.core.model.customxml.ResolvedCustomXmlImport mapping)
      implements ResolvedMutationAction {
    public ImportCustomXmlMapping {
      java.util.Objects.requireNonNull(mapping, "mapping must not be null");
    }
  }

  /** Materialized operands for SetDataValidation. */
  record SetDataValidation(
      dev.erst.gridgrind.core.model.validation.ResolvedDataValidation validation)
      implements ResolvedMutationAction {
    public SetDataValidation {
      java.util.Objects.requireNonNull(validation, "validation must not be null");
    }
  }

  /** Materialized operands for SetTable. */
  record SetTable(dev.erst.gridgrind.core.model.table.ResolvedTable table)
      implements ResolvedMutationAction {
    public SetTable {
      java.util.Objects.requireNonNull(table, "table must not be null");
    }
  }

  /** Materialized operands for SetPrintLayout. */
  record SetPrintLayout(dev.erst.gridgrind.core.model.print.ResolvedPrintLayout printLayout)
      implements ResolvedMutationAction {
    public SetPrintLayout {
      java.util.Objects.requireNonNull(printLayout, "printLayout must not be null");
    }
  }

  /** Materialized operands for SetSheetProtection. */
  record SetSheetProtection(
      dev.erst.gridgrind.core.model.protection.SheetProtectionSettings protection,
      java.util.Optional<String> password)
      implements ResolvedMutationAction {
    public SetSheetProtection {
      java.util.Objects.requireNonNull(protection, "protection must not be null");
      java.util.Objects.requireNonNull(password, "password must not be null");
    }

    @Override
    public String toString() {
      return "SetSheetProtection[protection=" + protection + ", password=<redacted>]";
    }
  }

  /** Materialized operands for SetWorkbookProtection. */
  record SetWorkbookProtection(
      dev.erst.gridgrind.core.model.protection.ResolvedWorkbookProtection protection)
      implements ResolvedMutationAction {
    public SetWorkbookProtection {
      java.util.Objects.requireNonNull(protection, "protection must not be null");
    }
  }
}
