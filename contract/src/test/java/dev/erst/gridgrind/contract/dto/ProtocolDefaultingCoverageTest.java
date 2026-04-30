package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import dev.erst.gridgrind.excel.foundation.ExcelAutofilterSortMethod;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisCrosses;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarShape;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelChartGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartMarkerStyle;
import dev.erst.gridgrind.excel.foundation.ExcelChartRadarStyle;
import dev.erst.gridgrind.excel.foundation.ExcelChartScatterStyle;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelIgnoredErrorType;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureDigestAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.json.JsonMapper;

/** Covers explicit request constructors, convenience factories, and strict creator branches. */
class ProtocolDefaultingCoverageTest {
  @Test
  void explicitRequestConstructorsPreserveAuthoredValues() {
    JsonMapper mapper = JsonMapper.builder().build();
    CalculationPolicyInput calculation =
        new CalculationPolicyInput(new CalculationStrategyInput.DoNotCalculate(), false);
    ExecutionModeInput mode =
        new ExecutionModeInput(
            ExecutionModeInput.ReadMode.FULL_XSSF, ExecutionModeInput.WriteMode.FULL_XSSF);
    ExecutionJournalInput journalInput = new ExecutionJournalInput(ExecutionJournalLevel.NORMAL);
    FormulaEnvironmentInput environment = FormulaEnvironmentInput.empty();
    SheetOutlineSummaryInput outlineSummary = SheetOutlineSummaryInput.defaults();
    SheetDisplayInput display = SheetDisplayInput.defaults();
    SheetDefaultsInput sheetDefaults = SheetDefaultsInput.defaults();
    HeaderFooterTextInput headerFooter = HeaderFooterTextInput.blank();
    SheetPresentationInput presentation = SheetPresentationInput.defaults();
    WorkbookProtectionInput protection =
        new WorkbookProtectionInput(false, false, false, null, null);
    AutofilterSortConditionInput sortCondition =
        new AutofilterSortConditionInput.Value("A1:A2", true);
    AutofilterFilterColumnInput filterColumn =
        new AutofilterFilterColumnInput(
            1L, true, new AutofilterFilterCriterionInput.Values(List.of("Ready"), false));
    AutofilterSortStateInput sortState =
        AutofilterSortStateInput.withoutSortMethod("A1:B2", false, false, List.of(sortCondition));
    FormulaUdfFunctionInput udfFunction = new FormulaUdfFunctionInput("DOUBLE", 1, "ARG1*2");
    FormulaUdfFunctionInput explicitUdfFunction =
        new FormulaUdfFunctionInput("TRIPLE", 1, Integer.valueOf(3), "ARG1*3");
    DataValidationPromptInput validationPrompt =
        readJson(
            mapper,
            """
            {
              "title": { "type": "INLINE", "text": "Status" },
              "text": { "type": "INLINE", "text": "Choose one." },
              "showPromptBox": true
            }
            """,
            DataValidationPromptInput.class);
    DataValidationErrorAlertInput validationAlert =
        readJson(
            mapper,
            """
            {
              "style": "STOP",
              "title": { "type": "INLINE", "text": "Invalid status" },
              "text": { "type": "INLINE", "text": "Use one of the allowed values." },
              "showErrorBox": true
            }
            """,
            DataValidationErrorAlertInput.class);
    OoxmlSignatureInput signature =
        OoxmlSignatureInput.sameKeyPassword(
            "keys/signing.p12",
            "store-pass",
            Optional.empty(),
            ExcelOoxmlSignatureDigestAlgorithm.SHA256,
            Optional.empty());
    OoxmlSignatureInput explicitSignature =
        OoxmlSignatureInput.create(
            "keys/signing.p12",
            "store-pass",
            "key-pass",
            Optional.of("gridgrind-signing"),
            ExcelOoxmlSignatureDigestAlgorithm.SHA512,
            Optional.of("Signed workbook"));
    RequestDoctorReport report =
        new RequestDoctorReport(
            GridGrindProtocolVersion.current(),
            AnalysisSeverity.INFO,
            true,
            Optional.of(
                new RequestDoctorReport.Summary(
                    "NEW",
                    "NONE",
                    "FULL_XSSF",
                    "FULL_XSSF",
                    "DO_NOT_CALCULATE",
                    false,
                    false,
                    0,
                    0,
                    0,
                    0)),
            List.of(),
            Optional.empty());

    assertEquals(new CalculationStrategyInput.DoNotCalculate(), calculation.strategy());
    assertFalse(calculation.markRecalculateOnOpen());
    assertTrue(mode.isDefault());
    assertEquals(ExecutionJournalLevel.NORMAL, journalInput.level());
    assertTrue(environment.isEmpty());
    assertEquals(SheetOutlineSummaryInput.defaults(), outlineSummary);
    assertEquals(SheetDisplayInput.defaults(), display);
    assertEquals(SheetDefaultsInput.defaults(), sheetDefaults);
    assertEquals(HeaderFooterTextInput.blank(), headerFooter);
    assertEquals(SheetPresentationInput.defaults(), presentation);
    assertFalse(protection.structureLocked());
    assertFalse(protection.windowsLocked());
    assertFalse(protection.revisionsLocked());
    assertInstanceOf(AutofilterSortConditionInput.Value.class, sortCondition);
    assertTrue(filterColumn.showButton());
    assertFalse(sortState.caseSensitive());
    assertFalse(sortState.columnSort());
    assertEquals(Optional.empty(), sortState.sortMethod());
    assertEquals(1, udfFunction.maximumArgumentCount());
    assertEquals(3, explicitUdfFunction.maximumArgumentCount());
    assertTrue(validationPrompt.showPromptBox());
    assertTrue(validationAlert.showErrorBox());
    assertEquals("store-pass", signature.keyPassword());
    assertEquals(Optional.empty(), signature.alias());
    assertEquals(ExcelOoxmlSignatureDigestAlgorithm.SHA256, signature.digestAlgorithm());
    assertEquals(Optional.empty(), signature.description());
    assertEquals("key-pass", explicitSignature.keyPassword());
    assertEquals(Optional.of("gridgrind-signing"), explicitSignature.alias());
    assertEquals(ExcelOoxmlSignatureDigestAlgorithm.SHA512, explicitSignature.digestAlgorithm());
    assertEquals(Optional.of("Signed workbook"), explicitSignature.description());
    assertEquals(GridGrindProtocolVersion.current(), report.protocolVersion());
    assertEquals(List.of(), report.warnings());
    assertEquals(Optional.empty(), ProtocolRgbColorSupport.normalizeRgbHex(null, null));
  }

  @Test
  void requestConstructorsRejectOmittedWireValues() {
    assertThrows(NullPointerException.class, () -> new CalculationPolicyInput(null, null));
    assertThrows(NullPointerException.class, () -> new ExecutionModeInput(null, null));
    assertThrows(NullPointerException.class, () -> new ExecutionJournalInput(null));
    assertThrows(NullPointerException.class, () -> new FormulaEnvironmentInput(null, null, null));
    assertThrows(
        NullPointerException.class, () -> new SheetDisplayInput(null, null, null, null, null));
    assertThrows(NullPointerException.class, () -> new SheetOutlineSummaryInput(null, null));
    assertThrows(NullPointerException.class, () -> new SheetDefaultsInput(null, null));
    assertThrows(
        NullPointerException.class,
        () -> new DataValidationPromptInput(text("Status"), text("Choose one."), (Boolean) null));
    assertThrows(
        NullPointerException.class,
        () ->
            new DataValidationErrorAlertInput(
                ExcelDataValidationErrorStyle.STOP,
                text("Invalid status"),
                text("Use one of the allowed values."),
                (Boolean) null));
  }

  @Test
  void autofilterAndExecutionInputsCoverExplicitAndStrictBranches() {
    AutofilterFilterCriterionInput.Values criterion =
        new AutofilterFilterCriterionInput.Values(List.of("Ready"), false);
    AutofilterSortConditionInput valueCondition =
        new AutofilterSortConditionInput.Value("A1:A4", true);
    AutofilterFilterColumnInput visibleButtonColumn =
        new AutofilterFilterColumnInput(2L, Boolean.TRUE, criterion);
    AutofilterFilterColumnInput hiddenButtonColumn =
        new AutofilterFilterColumnInput(2L, Boolean.FALSE, criterion);
    AutofilterSortStateInput defaultedSortState =
        AutofilterSortStateInput.withoutSortMethod("A1:B4", false, false, List.of(valueCondition));
    AutofilterSortStateInput explicitSortState =
        new AutofilterSortStateInput(
            "A1:B4",
            Boolean.TRUE,
            Boolean.TRUE,
            Optional.of(ExcelAutofilterSortMethod.PINYIN),
            List.of(valueCondition));
    ExecutionModeInput customMode =
        new ExecutionModeInput(
            ExecutionModeInput.ReadMode.EVENT_READ, ExecutionModeInput.WriteMode.STREAMING_WRITE);
    ExecutionModeInput writeOnlyCustomMode =
        ExecutionModeInput.writeMode(ExecutionModeInput.WriteMode.STREAMING_WRITE);
    ExecutionPolicyInput customPolicy = ExecutionPolicyInput.mode(customMode);
    ExecutionPolicyInput customJournalPolicy =
        ExecutionPolicyInput.journal(new ExecutionJournalInput(ExecutionJournalLevel.VERBOSE));
    SheetDisplayInput explicitDisplay =
        new SheetDisplayInput(
            Boolean.FALSE, Boolean.TRUE, Boolean.FALSE, Boolean.TRUE, Boolean.FALSE);
    SheetOutlineSummaryInput explicitOutline =
        new SheetOutlineSummaryInput(Boolean.FALSE, Boolean.TRUE);
    SheetDefaultsInput explicitSheetDefaults =
        new SheetDefaultsInput(Integer.valueOf(12), Double.valueOf(18.5d));
    assertTrue(visibleButtonColumn.showButton());
    assertFalse(hiddenButtonColumn.showButton());
    assertInstanceOf(AutofilterSortConditionInput.Value.class, valueCondition);
    assertFalse(defaultedSortState.caseSensitive());
    assertFalse(defaultedSortState.columnSort());
    assertEquals(Optional.empty(), defaultedSortState.sortMethod());
    assertTrue(explicitSortState.caseSensitive());
    assertTrue(explicitSortState.columnSort());
    assertEquals(Optional.of(ExcelAutofilterSortMethod.PINYIN), explicitSortState.sortMethod());
    assertFalse(customMode.isDefault());
    assertFalse(writeOnlyCustomMode.isDefault());
    assertFalse(customPolicy.isDefault());
    assertFalse(customJournalPolicy.isDefault());
    assertFalse(explicitDisplay.displayGridlines());
    assertTrue(explicitDisplay.displayZeros());
    assertFalse(explicitDisplay.displayRowColHeadings());
    assertTrue(explicitDisplay.displayFormulas());
    assertFalse(explicitDisplay.rightToLeft());
    assertFalse(explicitOutline.rowSumsBelow());
    assertTrue(explicitOutline.rowSumsRight());
    assertEquals(12, explicitSheetDefaults.defaultColumnWidth());
    assertEquals(18.5d, explicitSheetDefaults.defaultRowHeightPoints());
    assertThrows(
        NullPointerException.class,
        () -> new AutofilterFilterColumnInput(2L, (Boolean) null, criterion));
    assertThrows(
        NullPointerException.class,
        () ->
            new AutofilterSortStateInput(
                "A1:B4", (Boolean) null, (Boolean) null, null, List.of(valueCondition)));
  }

  @Test
  void tableWireConstructorsCoverExplicitAndDefaultedBranches() {
    TableInput defaultedTable =
        TableInput.withDefaultMetadata(
            "Budget", "Budget", "A1:B3", false, new TableStyleInput.None());
    TableInput explicitTable =
        new TableInput(
            "StyledBudget",
            "Budget",
            "A1:C3",
            true,
            false,
            new TableStyleInput.Named("TableStyleMedium2", true, false, true, false),
            text("Quarterly budget"),
            true,
            true,
            true,
            "HeaderStyle",
            "DataStyle",
            "TotalsStyle",
            List.of(new TableColumnInput(0, "owner", "total", "sum", "=[@Amount]")));
    TableInput explicitAutofilterTable =
        new TableInput(
            "VisibleAutofilterBudget",
            "Budget",
            "A1:C3",
            Boolean.FALSE,
            Boolean.TRUE,
            new TableStyleInput.None(),
            text(""),
            Boolean.FALSE,
            Boolean.FALSE,
            Boolean.FALSE,
            "",
            "",
            "",
            List.of());
    TableInput explicitTotalsAndNoAutofilterTable =
        new TableInput(
            "TotalsWithoutAutofilterBudget",
            "Budget",
            "A1:C3",
            Boolean.TRUE,
            Boolean.FALSE,
            new TableStyleInput.None(),
            text(""),
            Boolean.FALSE,
            Boolean.FALSE,
            Boolean.FALSE,
            "",
            "",
            "",
            List.of());
    TableEntryReport defaultedTableReport =
        new TableEntryReport(
            "BudgetTable",
            "Budget",
            "A1:B3",
            1,
            0,
            List.of("Owner"),
            List.of(new TableColumnReport(0L, "Owner")),
            new TableStyleReport.None(),
            true,
            Optional.empty(),
            false,
            false,
            false,
            Optional.empty(),
            Optional.empty(),
            Optional.empty());

    assertEquals(text(""), defaultedTable.comment());
    assertFalse(defaultedTable.published());
    assertFalse(defaultedTable.insertRow());
    assertFalse(defaultedTable.insertRowShift());
    assertEquals("", defaultedTable.headerRowCellStyle());
    assertEquals("", defaultedTable.dataCellStyle());
    assertEquals("", defaultedTable.totalsRowCellStyle());
    assertEquals(List.of(), defaultedTable.columns());
    assertTrue(explicitTable.published());
    assertTrue(explicitTable.insertRow());
    assertTrue(explicitTable.insertRowShift());
    assertEquals("HeaderStyle", explicitTable.headerRowCellStyle());
    assertEquals("DataStyle", explicitTable.dataCellStyle());
    assertEquals("TotalsStyle", explicitTable.totalsRowCellStyle());
    assertFalse(explicitAutofilterTable.showTotalsRow());
    assertTrue(explicitAutofilterTable.hasAutofilter());
    assertTrue(explicitTotalsAndNoAutofilterTable.showTotalsRow());
    assertFalse(explicitTotalsAndNoAutofilterTable.hasAutofilter());
    assertEquals(Optional.empty(), defaultedTableReport.comment());
    assertEquals(Optional.empty(), defaultedTableReport.headerRowCellStyle());
    assertEquals(Optional.empty(), defaultedTableReport.dataCellStyle());
    assertEquals(Optional.empty(), defaultedTableReport.totalsRowCellStyle());
    assertThrows(
        NullPointerException.class,
        () ->
            new TableInput(
                "Budget",
                "Budget",
                "A1:B3",
                null,
                null,
                new TableStyleInput.None(),
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null));
  }

  @Test
  void validationChartPrintAndSignatureWireConstructorsCoverExplicitAndDefaultedBranches() {
    JsonMapper mapper = JsonMapper.builder().build();
    DataValidationErrorAlertInput explicitAlert =
        new DataValidationErrorAlertInput(
            ExcelDataValidationErrorStyle.STOP,
            text("Invalid"),
            text("Choose an allowed value."),
            Boolean.TRUE);
    DataValidationErrorAlertInput hiddenAlert =
        new DataValidationErrorAlertInput(
            ExcelDataValidationErrorStyle.WARNING,
            text("Caution"),
            text("Review the warning."),
            Boolean.FALSE);
    ChartSeriesInput series =
        new ChartSeriesInput(
            new ChartTitleInput.None(),
            new ChartDataSourceInput.StringLiteral(List.of("Jan")),
            new ChartDataSourceInput.NumericLiteral(List.of(10.0d)),
            true,
            ExcelChartMarkerStyle.CIRCLE,
            (short) 6,
            null);
    ChartInput defaultedChart =
        ChartInput.withStandardDisplay(
            "BudgetChart",
            twoCellAnchor(),
            List.of(new ChartPlotInput.Line(false, ExcelChartGrouping.STANDARD, List.of(series))));
    ChartInput explicitChart =
        new ChartInput(
            "BudgetChartExplicit",
            twoCellAnchor(),
            new ChartTitleInput.Text(text("Budget")),
            new ChartLegendInput.Hidden(),
            ExcelChartDisplayBlanksAs.SPAN,
            Boolean.FALSE,
            List.of(new ChartPlotInput.Line(false, ExcelChartGrouping.STANDARD, List.of(series))));
    ChartAxisInput defaultedAxis =
        ChartAxisInput.visible(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO);
    ChartAxisInput hiddenAxis =
        new ChartAxisInput(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            Boolean.FALSE);
    PrintSetupInput defaultedPrintSetup = PrintSetupInput.defaults();
    PrintSetupInput explicitPrintSetup =
        new PrintSetupInput(
            new PrintMarginsInput(0.6d, 0.6d, 0.7d, 0.7d, 0.2d, 0.2d),
            Boolean.TRUE,
            Boolean.TRUE,
            Boolean.TRUE,
            Integer.valueOf(9),
            Boolean.TRUE,
            Boolean.TRUE,
            Integer.valueOf(2),
            Boolean.TRUE,
            Integer.valueOf(7),
            List.of(3),
            List.of(4));
    SignatureLineInput defaultedSignatureLine =
        readJson(
            mapper,
            """
            {
              "name": "OpsSignature",
              "anchor": {
                "type": "TWO_CELL",
                "from": { "columnIndex": 1, "rowIndex": 0, "dx": 1, "dy": 0 },
                "to": { "columnIndex": 5, "rowIndex": 10, "dx": 1, "dy": 0 },
                "behavior": "MOVE_AND_RESIZE"
              },
              "allowComments": true,
              "suggestedSigner": "Ada Lovelace",
              "caption": "Budget approval"
            }
            """,
            SignatureLineInput.class);

    assertTrue(explicitAlert.showErrorBox());
    assertFalse(hiddenAlert.showErrorBox());
    assertTrue(defaultedChart.title() instanceof ChartTitleInput.None);
    assertEquals(
        new ChartLegendInput.Visible(ExcelChartLegendPosition.RIGHT), defaultedChart.legend());
    assertEquals(ExcelChartDisplayBlanksAs.GAP, defaultedChart.displayBlanksAs());
    assertTrue(defaultedChart.plotOnlyVisibleCells());
    assertTrue(explicitChart.legend() instanceof ChartLegendInput.Hidden);
    assertEquals(ExcelChartDisplayBlanksAs.SPAN, explicitChart.displayBlanksAs());
    assertFalse(explicitChart.plotOnlyVisibleCells());
    assertTrue(defaultedAxis.visible());
    assertFalse(hiddenAxis.visible());
    assertEquals(PrintSetupInput.defaults(), defaultedPrintSetup);
    assertTrue(explicitPrintSetup.printGridlines());
    assertTrue(explicitPrintSetup.horizontallyCentered());
    assertTrue(explicitPrintSetup.verticallyCentered());
    assertEquals(9, explicitPrintSetup.paperSize());
    assertTrue(explicitPrintSetup.draft());
    assertTrue(explicitPrintSetup.blackAndWhite());
    assertEquals(2, explicitPrintSetup.copies());
    assertTrue(explicitPrintSetup.useFirstPageNumber());
    assertEquals(7, explicitPrintSetup.firstPageNumber());
    assertEquals(List.of(3), explicitPrintSetup.rowBreaks());
    assertEquals(List.of(4), explicitPrintSetup.columnBreaks());
    assertTrue(defaultedSignatureLine.allowComments());
    assertThrows(
        NullPointerException.class,
        () ->
            new ChartInput(
                "BudgetChart",
                twoCellAnchor(),
                null,
                null,
                null,
                null,
                List.of(
                    new ChartPlotInput.Line(false, ExcelChartGrouping.STANDARD, List.of(series)))));
    assertThrows(
        NullPointerException.class,
        () ->
            new ChartAxisInput(
                ExcelChartAxisKind.CATEGORY,
                ExcelChartAxisPosition.BOTTOM,
                ExcelChartAxisCrosses.AUTO_ZERO,
                null));
    assertThrows(
        NullPointerException.class,
        () ->
            new PrintSetupInput(
                null, null, null, null, null, null, null, null, null, null, null, null));
  }

  @Test
  void convenienceConstructorsCoverDefaultPlanAndReportBranches() {
    WorkbookPlan defaultPlan =
        WorkbookPlan.identified(
            GridGrindProtocolVersion.current(),
            "budget-defaults",
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());
    WorkbookPlan explicitExecutionPlan =
        WorkbookPlan.identified(
            GridGrindProtocolVersion.current(),
            "budget-execution",
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.journal(new ExecutionJournalInput(ExecutionJournalLevel.VERBOSE)),
            FormulaEnvironmentInput.empty(),
            List.of());
    TableColumnReport tableColumn = new TableColumnReport(4L, "Owner");
    AutofilterSortConditionReport conditionReport =
        new AutofilterSortConditionReport.CellColor("B2:B9", false, CellColorReport.rgb("#aabbcc"));
    AutofilterSortConditionReport createdConditionReport =
        new AutofilterSortConditionReport.Icon("C2:C9", true, 2);
    AutofilterSortStateReport stateReport =
        AutofilterSortStateReport.withoutSortMethod("A1:F9", false, true, List.of(conditionReport));
    AutofilterSortStateReport createdStateReport =
        AutofilterSortStateReport.withoutSortMethod(
            "A1:F9", true, false, List.of(createdConditionReport));
    DifferentialBorderSideInput borderSide =
        new DifferentialBorderSideInput(ExcelBorderStyle.THIN, Optional.empty());
    DifferentialBorderInput borderInput =
        new DifferentialBorderInput(
            Optional.of(borderSide),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    DifferentialBorderReport borderReport =
        new DifferentialBorderReport(
            Optional.empty(),
            Optional.of(new DifferentialBorderSideReport(ExcelBorderStyle.THIN, Optional.empty())),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    AutofilterSortConditionInput explicitSortCondition =
        new AutofilterSortConditionInput.Icon("D2:D9", false, 1);
    AutofilterFilterColumnInput explicitFilterColumn =
        new AutofilterFilterColumnInput(
            2L, false, new AutofilterFilterCriterionInput.Values(List.of("Queued"), true));
    AutofilterSortStateInput explicitSortState =
        new AutofilterSortStateInput(
            "A1:D9",
            true,
            true,
            Optional.of(ExcelAutofilterSortMethod.PINYIN),
            List.of(explicitSortCondition));
    AutofilterSortConditionReport explicitConditionReport =
        new AutofilterSortConditionReport.Icon("D2:D9", false, 1);
    AutofilterSortStateReport explicitStateReport =
        new AutofilterSortStateReport(
            "A1:D9",
            true,
            true,
            Optional.of(ExcelAutofilterSortMethod.PINYIN),
            List.of(explicitConditionReport));
    SheetDisplayInput explicitDisplay = new SheetDisplayInput(false, true, false, true, true);
    SheetOutlineSummaryInput explicitOutline = new SheetOutlineSummaryInput(false, true);
    SheetDefaultsInput explicitSheetDefaults = new SheetDefaultsInput(18, 24.5d);
    SheetPresentationInput explicitPresentation =
        SheetPresentationInput.create(
            explicitDisplay,
            Optional.of(ColorInput.rgb("#aabbcc")),
            explicitOutline,
            explicitSheetDefaults,
            List.of(
                new IgnoredErrorInput(
                    "A1:A3", List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT))));
    FormulaEnvironmentInput explicitEnvironment =
        new FormulaEnvironmentInput(
            List.of(new FormulaExternalWorkbookInput("Rates.xlsx", "tmp/rates.xlsx")),
            FormulaMissingWorkbookPolicy.USE_CACHED_VALUE,
            List.of());

    assertTrue(defaultPlan.execution().isDefault());
    assertTrue(defaultPlan.formulaEnvironment().isEmpty());
    assertEquals(
        ExecutionJournalLevel.VERBOSE, explicitExecutionPlan.execution().journal().level());
    assertTrue(explicitExecutionPlan.formulaEnvironment().isEmpty());
    assertEquals(Optional.empty(), tableColumn.uniqueName());
    assertEquals(Optional.empty(), tableColumn.totalsRowLabel());
    assertInstanceOf(AutofilterSortConditionReport.CellColor.class, conditionReport);
    assertInstanceOf(AutofilterSortConditionReport.Icon.class, createdConditionReport);
    assertEquals(Optional.empty(), stateReport.sortMethod());
    assertEquals(Optional.empty(), createdStateReport.sortMethod());
    assertEquals(borderSide, borderInput.all().orElseThrow());
    assertEquals(ExcelBorderStyle.THIN, borderReport.top().orElseThrow().style());
    assertInstanceOf(AutofilterSortConditionInput.Icon.class, explicitSortCondition);
    assertFalse(explicitFilterColumn.showButton());
    assertEquals(Optional.of(ExcelAutofilterSortMethod.PINYIN), explicitSortState.sortMethod());
    assertInstanceOf(AutofilterSortConditionReport.Icon.class, explicitConditionReport);
    assertEquals(Optional.of(ExcelAutofilterSortMethod.PINYIN), explicitStateReport.sortMethod());
    assertEquals(explicitDisplay, explicitPresentation.display());
    assertEquals(explicitOutline, explicitPresentation.outlineSummary());
    assertEquals(explicitSheetDefaults, explicitPresentation.sheetDefaults());
    assertEquals(1, explicitPresentation.ignoredErrors().size());
    assertEquals(
        FormulaMissingWorkbookPolicy.USE_CACHED_VALUE, explicitEnvironment.missingWorkbookPolicy());
  }

  @Test
  void chartConstructorsAndConvenienceConstructorsSupplyExplicitDefaults() {
    ChartSeriesInput explicitSeries =
        new ChartSeriesInput(
            new ChartTitleInput.None(),
            new ChartDataSourceInput.StringLiteral(List.of("Jan", "Feb")),
            new ChartDataSourceInput.NumericLiteral(List.of(10.0d, 18.0d)),
            true,
            ExcelChartMarkerStyle.DIAMOND,
            (short) 8,
            12L);
    ChartSeriesInput convenienceSeries =
        ChartSeriesInput.untitled(
            new ChartDataSourceInput.Reference("Categories"),
            new ChartDataSourceInput.Reference("Values"),
            false,
            ExcelChartMarkerStyle.CIRCLE,
            (short) 6,
            null);
    ChartAxisInput defaultAxis =
        ChartAxisInput.visible(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO);
    List<ChartAxisInput> explicitAxes =
        List.of(
            defaultAxis,
            new ChartAxisInput(
                ExcelChartAxisKind.VALUE,
                ExcelChartAxisPosition.LEFT,
                ExcelChartAxisCrosses.AUTO_ZERO,
                false));
    ChartPlotInput.Area explicitArea =
        new ChartPlotInput.Area(
            true, ExcelChartGrouping.PERCENT_STACKED, explicitAxes, List.of(explicitSeries));
    ChartPlotInput.Area convenienceArea =
        new ChartPlotInput.Area(false, ExcelChartGrouping.STANDARD, List.of(explicitSeries));
    ChartPlotInput.Area3D explicitArea3D =
        new ChartPlotInput.Area3D(
            false, ExcelChartGrouping.STACKED, 24, explicitAxes, List.of(explicitSeries));
    ChartPlotInput.Area3D convenienceArea3D =
        new ChartPlotInput.Area3D(false, ExcelChartGrouping.STANDARD, 16, List.of(explicitSeries));
    ChartPlotInput.Bar3D explicitBar3D =
        new ChartPlotInput.Bar3D(
            false,
            ExcelChartBarDirection.BAR,
            ExcelChartBarGrouping.PERCENT_STACKED,
            12,
            64,
            ExcelChartBarShape.CONE,
            explicitAxes,
            List.of(explicitSeries));
    ChartPlotInput.Bar3D convenienceBar3D =
        new ChartPlotInput.Bar3D(
            false,
            ExcelChartBarDirection.COLUMN,
            ExcelChartBarGrouping.CLUSTERED,
            18,
            90,
            ExcelChartBarShape.BOX,
            List.of(explicitSeries));
    ChartPlotInput.Line explicitLine =
        new ChartPlotInput.Line(
            true, ExcelChartGrouping.STACKED, explicitAxes, List.of(explicitSeries));
    ChartPlotInput.Line convenienceLine =
        new ChartPlotInput.Line(false, ExcelChartGrouping.STANDARD, List.of(explicitSeries));
    ChartPlotInput.Line3D explicitLine3D =
        new ChartPlotInput.Line3D(
            true, ExcelChartGrouping.PERCENT_STACKED, 32, explicitAxes, List.of(explicitSeries));
    ChartPlotInput.Line3D convenienceLine3D =
        new ChartPlotInput.Line3D(false, ExcelChartGrouping.STANDARD, 8, List.of(explicitSeries));
    ChartPlotInput.Radar explicitRadar =
        new ChartPlotInput.Radar(
            true, ExcelChartRadarStyle.MARKER, explicitAxes, List.of(explicitSeries));
    ChartPlotInput.Radar convenienceRadar =
        new ChartPlotInput.Radar(false, ExcelChartRadarStyle.STANDARD, List.of(explicitSeries));
    ChartPlotInput.Scatter explicitScatter =
        new ChartPlotInput.Scatter(
            true, ExcelChartScatterStyle.SMOOTH, explicitAxes, List.of(explicitSeries));
    ChartPlotInput.Scatter convenienceScatter =
        new ChartPlotInput.Scatter(false, ExcelChartScatterStyle.LINE, List.of(explicitSeries));
    ChartPlotInput.Surface explicitSurface =
        new ChartPlotInput.Surface(true, true, explicitAxes, List.of(explicitSeries));
    ChartPlotInput.Surface convenienceSurface =
        new ChartPlotInput.Surface(false, false, List.of(explicitSeries));
    ChartPlotInput.Surface3D explicitSurface3D =
        new ChartPlotInput.Surface3D(true, true, explicitAxes, List.of(explicitSeries));
    ChartPlotInput.Surface3D convenienceSurface3D =
        new ChartPlotInput.Surface3D(false, false, List.of(explicitSeries));
    ChartInput createdChart =
        new ChartInput(
            "CreatedChart",
            twoCellAnchor(),
            new ChartTitleInput.None(),
            new ChartLegendInput.Visible(ExcelChartLegendPosition.RIGHT),
            ExcelChartDisplayBlanksAs.GAP,
            true,
            List.of(explicitArea));
    ChartInput convenienceChart =
        ChartInput.withStandardDisplay("ConvenienceChart", twoCellAnchor(), List.of(explicitLine));

    assertTrue(defaultAxis.visible());
    assertTrue(convenienceSeries.title() instanceof ChartTitleInput.None);
    assertEquals(2, explicitArea.axes().size());
    assertEquals(2, convenienceArea.axes().size());
    assertEquals(2, explicitArea3D.axes().size());
    assertEquals(2, convenienceArea3D.axes().size());
    assertEquals(2, explicitBar3D.axes().size());
    assertEquals(2, convenienceBar3D.axes().size());
    assertEquals(2, explicitLine.axes().size());
    assertEquals(2, convenienceLine.axes().size());
    assertEquals(2, explicitLine3D.axes().size());
    assertEquals(2, convenienceLine3D.axes().size());
    assertEquals(2, explicitRadar.axes().size());
    assertEquals(2, convenienceRadar.axes().size());
    assertEquals(2, explicitScatter.axes().size());
    assertEquals(2, convenienceScatter.axes().size());
    assertEquals(2, explicitSurface.axes().size());
    assertEquals(3, convenienceSurface.axes().size());
    assertEquals(2, explicitSurface3D.axes().size());
    assertEquals(3, convenienceSurface3D.axes().size());
    assertTrue(createdChart.title() instanceof ChartTitleInput.None);
    assertEquals(
        new ChartLegendInput.Visible(ExcelChartLegendPosition.RIGHT), createdChart.legend());
    assertEquals(ExcelChartDisplayBlanksAs.GAP, createdChart.displayBlanksAs());
    assertTrue(createdChart.plotOnlyVisibleCells());
    assertTrue(convenienceChart.title() instanceof ChartTitleInput.None);
  }

  @Test
  void creatorsRequireExplicitWireValues() {
    JsonMapper mapper = JsonMapper.builder().build();
    PrintSetupInput printSetup =
        mapper.readValue(
            """
            {
              "margins": {
                "left": 0.6,
                "right": 0.6,
                "top": 0.7,
                "bottom": 0.7,
                "header": 0.3,
                "footer": 0.3
              },
              "printGridlines": false,
              "horizontallyCentered": false,
              "verticallyCentered": false,
              "paperSize": 9,
              "draft": false,
              "blackAndWhite": false,
              "copies": 1,
              "useFirstPageNumber": false,
              "firstPageNumber": 0,
              "rowBreaks": [],
              "columnBreaks": []
            }
            """,
            PrintSetupInput.class);
    PrintLayoutInput printLayout =
        mapper.readValue(
            """
            {
              "printArea": { "type": "NONE" },
              "orientation": "PORTRAIT",
              "scaling": { "type": "AUTOMATIC" },
              "repeatingRows": { "type": "NONE" },
              "repeatingColumns": { "type": "NONE" },
              "header": {
                "left": { "type": "INLINE", "text": "" },
                "center": { "type": "INLINE", "text": "" },
                "right": { "type": "INLINE", "text": "" }
              },
              "footer": {
                "left": { "type": "INLINE", "text": "" },
                "center": { "type": "INLINE", "text": "" },
                "right": { "type": "INLINE", "text": "" }
              },
              "setup": {
                "margins": {
                  "left": 0.6,
                  "right": 0.6,
                  "top": 0.7,
                  "bottom": 0.7,
                  "header": 0.3,
                  "footer": 0.3
                },
                "printGridlines": true,
                "horizontallyCentered": false,
                "verticallyCentered": false,
                "paperSize": 9,
                "draft": false,
                "blackAndWhite": false,
                "copies": 2,
                "useFirstPageNumber": false,
                "firstPageNumber": 0,
                "rowBreaks": [],
                "columnBreaks": []
              }
            }
            """,
            PrintLayoutInput.class);
    PrintLayoutInput conveniencePrintLayout =
        PrintLayoutInput.withDefaultSetup(
            new PrintAreaInput.None(),
            ExcelPrintOrientation.LANDSCAPE,
            new PrintScalingInput.Automatic(),
            new PrintTitleRowsInput.None(),
            new PrintTitleColumnsInput.None(),
            HeaderFooterTextInput.blank(),
            HeaderFooterTextInput.blank());
    PrintLayoutInput createdPrintLayout =
        PrintLayoutInput.withDefaultSetup(
            new PrintAreaInput.None(),
            ExcelPrintOrientation.LANDSCAPE,
            new PrintScalingInput.Automatic(),
            new PrintTitleRowsInput.None(),
            new PrintTitleColumnsInput.None(),
            HeaderFooterTextInput.blank(),
            HeaderFooterTextInput.blank());
    SignatureLineInput signatureLine =
        mapper.readValue(
            """
            {
              "name": "OpsSignature",
              "anchor": {
                "type": "TWO_CELL",
                "from": { "columnIndex": 1, "rowIndex": 0, "dx": 1, "dy": 0 },
                "to": { "columnIndex": 5, "rowIndex": 10, "dx": 1, "dy": 0 },
                "behavior": "MOVE_AND_RESIZE"
              },
              "allowComments": true,
              "signingInstructions": "Review before signing.",
              "suggestedSigner": "Ada Lovelace",
              "suggestedSigner2": "Finance",
              "suggestedSignerEmail": "ada@example.com",
              "caption": "Budget approval",
              "invalidStamp": "invalid"
            }
            """,
            SignatureLineInput.class);
    TableEntryReport tableReport =
        mapper.readValue(
            """
            {
              "name": "BudgetTable",
              "sheetName": "Budget",
              "range": "A1:B4",
              "headerRowCount": 1,
              "totalsRowCount": 0,
              "columnNames": ["Owner"],
              "columns": [{ "id": 0, "name": "Owner" }],
              "style": { "type": "NONE" },
              "hasAutofilter": true,
              "comment": "",
              "published": false,
              "insertRow": false,
              "insertRowShift": false,
              "headerRowCellStyle": "",
              "dataCellStyle": "",
              "totalsRowCellStyle": ""
            }
            """,
            TableEntryReport.class);
    ExecutionJournal journal =
        mapper.readValue(
            """
            {
              "planId": "journal-plan",
              "level": "NORMAL",
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "validation": { "status": "NOT_STARTED", "durationMillis": 0 },
              "inputResolution": { "status": "NOT_STARTED", "durationMillis": 0 },
              "open": { "status": "NOT_STARTED", "durationMillis": 0 },
              "calculation": {
                "preflight": { "status": "NOT_STARTED", "durationMillis": 0 },
                "execution": { "status": "NOT_STARTED", "durationMillis": 0 }
              },
              "persistencePhase": { "status": "NOT_STARTED", "durationMillis": 0 },
              "close": { "status": "NOT_STARTED", "durationMillis": 0 },
              "steps": [],
              "warnings": [],
              "outcome": {
                "status": "SUCCEEDED",
                "plannedStepCount": 0,
                "completedStepCount": 0,
                "durationMillis": 0
              },
              "events": []
            }
            """,
            ExecutionJournal.class);

    assertEquals(new PrintMarginsInput(0.6d, 0.6d, 0.7d, 0.7d, 0.3d, 0.3d), printSetup.margins());
    assertEquals(ExcelPrintOrientation.PORTRAIT, printLayout.orientation());
    assertTrue(printLayout.setup().printGridlines());
    assertEquals(PrintSetupInput.defaults(), conveniencePrintLayout.setup());
    assertEquals(PrintSetupInput.defaults(), createdPrintLayout.setup());
    assertTrue(signatureLine.allowComments());
    assertEquals(Optional.empty(), tableReport.comment());
    assertEquals(Optional.empty(), tableReport.headerRowCellStyle());
    assertEquals(Optional.empty(), tableReport.dataCellStyle());
    assertEquals(Optional.empty(), tableReport.totalsRowCellStyle());
    assertEquals(ExecutionJournalLevel.NORMAL, journal.level());
    assertEquals(List.of(), journal.warnings());
    assertEquals(List.of(), journal.events());
  }

  private static DrawingAnchorInput.TwoCell twoCellAnchor() {
    return new DrawingAnchorInput.TwoCell(
        new DrawingMarkerInput(1, 0, 1, 0),
        new DrawingMarkerInput(5, 0, 10, 0),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  private static TextSourceInput text(String value) {
    return TextSourceInput.inline(value);
  }

  private static <T> T readJson(JsonMapper mapper, String json, Class<T> targetType) {
    return mapper.readValue(json, targetType);
  }
}
