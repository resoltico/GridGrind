package dev.erst.gridgrind.core.model.security;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.chart.ChartAxisSpec;
import dev.erst.gridgrind.core.model.chart.ChartDataSourceSpec;
import dev.erst.gridgrind.core.model.chart.ChartLegendSpec;
import dev.erst.gridgrind.core.model.chart.ChartPlotSpec;
import dev.erst.gridgrind.core.model.chart.ChartSeriesSpec;
import dev.erst.gridgrind.core.model.chart.ChartSpec;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.drawing.CommentSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.SignatureLineSpec;
import dev.erst.gridgrind.core.model.formula.CalculationPolicySpec;
import dev.erst.gridgrind.core.model.formula.CalculationStrategySpec;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.formula.FormulaExternalWorkbookSpec;
import dev.erst.gridgrind.core.model.formula.FormulaMissingWorkbookPolicy;
import dev.erst.gridgrind.core.model.formula.FormulaUdfFunctionSpec;
import dev.erst.gridgrind.core.model.layout.SheetDefaultsSpec;
import dev.erst.gridgrind.core.model.layout.SheetDisplaySpec;
import dev.erst.gridgrind.core.model.layout.SheetOutlineSummarySpec;
import dev.erst.gridgrind.core.model.layout.SheetPresentationSpec;
import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextSpec;
import dev.erst.gridgrind.core.model.print.PrintAreaSpec;
import dev.erst.gridgrind.core.model.print.PrintLayoutSpec;
import dev.erst.gridgrind.core.model.print.PrintMarginsSpec;
import dev.erst.gridgrind.core.model.print.PrintScalingSpec;
import dev.erst.gridgrind.core.model.print.PrintSetupSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsSpec;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionSpec;
import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.DifferentialBorderPatch;
import dev.erst.gridgrind.core.model.style.DifferentialBorderReport;
import dev.erst.gridgrind.core.model.style.DifferentialBorderSideReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnSpec;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateSpec;
import dev.erst.gridgrind.core.model.table.TableColumnReport;
import dev.erst.gridgrind.core.model.table.TableColumnSpec;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.core.model.table.TableStyleReport;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationSpec;
import dev.erst.gridgrind.core.model.validation.IgnoredErrorSpec;
import dev.erst.gridgrind.core.plan.ExecutionJournalSpec;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.GridGrindProtocolVersion;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.RequestDoctorReport;
import dev.erst.gridgrind.core.result.ExecutionJournal;
import dev.erst.gridgrind.core.result.ExecutionJournalLevel;
import dev.erst.gridgrind.core.source.SecretReference;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import dev.erst.gridgrind.core.vocabulary.ExcelAutofilterSortMethod;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisCrosses;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisKind;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarShape;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelChartMarkerStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelChartRadarStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelChartScatterStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelIgnoredErrorType;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureDigestAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteCipher;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteHash;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.PlanId;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableLabel;
import dev.erst.gridgrind.core.vocabulary.TableName;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.json.JsonMapper;

/** Covers explicit request constructors, convenience factories, and record omission defaults. */
class ProtocolDefaultingCoverageTest {
  @Test
  void explicitRequestConstructorsPreserveAuthoredValues() {
    JsonMapper mapper = dev.erst.gridgrind.json.codec.GridGrindJsonMapperSupport.JSON_MAPPER;
    CalculationPolicySpec calculation =
        new CalculationPolicySpec(new CalculationStrategySpec.DoNotCalculate(), false);
    ExecutionModeSpec mode = ExecutionModeSpec.fullXssf();
    ExecutionJournalSpec journalInput = new ExecutionJournalSpec(ExecutionJournalLevel.NORMAL);
    FormulaEnvironmentSpec environment = FormulaEnvironmentSpec.empty();
    SheetOutlineSummarySpec outlineSummary = SheetOutlineSummarySpec.defaults();
    SheetDisplaySpec display = SheetDisplaySpec.defaults();
    SheetDefaultsSpec sheetDefaults = SheetDefaultsSpec.defaults();
    HeaderFooterTextSpec headerFooter = HeaderFooterTextSpec.blank();
    SheetPresentationSpec presentation = SheetPresentationSpec.defaults();
    WorkbookProtectionSpec protection =
        new WorkbookProtectionSpec(
            false, false, false, java.util.Optional.empty(), java.util.Optional.empty());
    AutofilterSortConditionSpec sortCondition =
        new AutofilterSortConditionSpec.Value(new RangeRef("A1:A2"), true);
    AutofilterFilterColumnSpec filterColumn =
        new AutofilterFilterColumnSpec(
            1L, true, new AutofilterFilterCriterionSpec.Values(List.of("Ready"), false));
    AutofilterSortStateSpec sortState =
        AutofilterSortStateSpec.withoutSortMethod(
            new RangeRef("A1:B2"), false, false, List.of(sortCondition));
    FormulaUdfFunctionSpec udfFunction = new FormulaUdfFunctionSpec("DOUBLE", 1, "ARG1*2");
    FormulaUdfFunctionSpec explicitUdfFunction =
        new FormulaUdfFunctionSpec("TRIPLE", 1, Integer.valueOf(3), "ARG1*3");
    DataValidationPromptSpec validationPrompt =
        readJson(
            mapper,
            """
            {
              "title": { "type": "INLINE", "text": "Status" },
              "text": { "type": "INLINE", "text": "Choose one." },
              "showPromptBox": true
            }
            """,
            DataValidationPromptSpec.class);
    DataValidationErrorAlertSpec validationAlert =
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
            DataValidationErrorAlertSpec.class);
    OoxmlSignatureSpec signature =
        OoxmlSignatureSpec.sameKeyPassword(
            "keys/signing.p12",
            new SecretReference("store-pass"),
            Optional.empty(),
            ExcelOoxmlSignatureDigestAlgorithm.SHA256,
            Optional.empty());
    OoxmlSignatureSpec explicitSignature =
        readOmitted(
            OoxmlSignatureSpec.class,
            "keys/signing.p12",
            new SecretReference("store-pass"),
            Optional.of(new SecretReference("key-pass")),
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
                    "NEW", "NONE", "FULL_XSSF", "DO_NOT_CALCULATE", false, false, 0, 0, 0, 0)),
            List.of(),
            List.of());

    assertEquals(new CalculationStrategySpec.DoNotCalculate(), calculation.strategy());
    assertFalse(calculation.markRecalculateOnOpen());
    assertTrue(mode.isDefault());
    assertEquals(ExecutionJournalLevel.NORMAL, journalInput.level());
    assertTrue(environment.isEmpty());
    assertEquals(SheetOutlineSummarySpec.defaults(), outlineSummary);
    assertEquals(SheetDisplaySpec.defaults(), display);
    assertEquals(SheetDefaultsSpec.defaults(), sheetDefaults);
    assertEquals(HeaderFooterTextSpec.blank(), headerFooter);
    assertEquals(SheetPresentationSpec.defaults(), presentation);
    assertFalse(protection.structureLocked());
    assertFalse(protection.windowsLocked());
    assertFalse(protection.revisionsLocked());
    assertInstanceOf(AutofilterSortConditionSpec.Value.class, sortCondition);
    assertTrue(filterColumn.showButton());
    assertFalse(sortState.caseSensitive());
    assertFalse(sortState.columnSort());
    assertEquals(Optional.empty(), sortState.sortMethod());
    assertEquals(1, udfFunction.maximumArgumentCount());
    assertEquals(3, explicitUdfFunction.maximumArgumentCount());
    assertTrue(validationPrompt.showPromptBox());
    assertTrue(validationAlert.showErrorBox());
    assertEquals(Optional.empty(), signature.keyPasswordRef());
    assertEquals(Optional.empty(), signature.alias());
    assertEquals(ExcelOoxmlSignatureDigestAlgorithm.SHA256, signature.digestAlgorithm());
    assertEquals(Optional.empty(), signature.description());
    assertEquals(Optional.of(new SecretReference("key-pass")), explicitSignature.keyPasswordRef());
    assertEquals(Optional.of("gridgrind-signing"), explicitSignature.alias());
    assertEquals(ExcelOoxmlSignatureDigestAlgorithm.SHA512, explicitSignature.digestAlgorithm());
    assertEquals(Optional.of("Signed workbook"), explicitSignature.description());
    assertEquals(GridGrindProtocolVersion.current(), report.protocolVersion());
    assertEquals(List.of(), report.warnings());
  }

  @Test
  void requestConstructorsDefaultWireBooleansWhileRejectingMissingRequiredState() {
    assertThrows(NullPointerException.class, () -> new CalculationPolicySpec(null, false));
    assertThrows(NullPointerException.class, () -> ExecutionPolicySpec.mode(null));
    assertThrows(NullPointerException.class, () -> new ExecutionJournalSpec(null));
    assertThrows(NullPointerException.class, () -> new FormulaEnvironmentSpec(null, null, null));
    assertThrows(
        tools.jackson.core.JacksonException.class,
        () -> readOmitted(SheetDisplaySpec.class, null, null, null, null, null));
    assertThrows(
        tools.jackson.core.JacksonException.class,
        () -> readOmitted(SheetOutlineSummarySpec.class, null, null));
    assertThrows(
        tools.jackson.core.JacksonException.class,
        () -> readOmitted(SheetDefaultsSpec.class, null, null));
    assertTrue(
        readOmitted(DataValidationPromptSpec.class, text("Status"), text("Choose one."), null)
            .showPromptBox());
    assertTrue(
        readOmitted(
                DataValidationErrorAlertSpec.class,
                ExcelDataValidationErrorStyle.STOP,
                text("Invalid status"),
                text("Use one of the allowed values."),
                null)
            .showErrorBox());
  }

  @Test
  void recordMetadataAppliesCatalogOwnedOmissionDefaults() {
    JsonMapper mapper = dev.erst.gridgrind.json.codec.GridGrindJsonMapperSupport.JSON_MAPPER;

    ExecutionPolicySpec execution = readJson(mapper, "{}", ExecutionPolicySpec.class);
    ExecutionJournalSpec journal = readJson(mapper, "{}", ExecutionJournalSpec.class);
    CalculationPolicySpec calculation = readJson(mapper, "{}", CalculationPolicySpec.class);
    FormulaEnvironmentSpec formulaEnvironment =
        readJson(mapper, "{}", FormulaEnvironmentSpec.class);
    OoxmlEncryptionSpec encryption =
        readJson(
            mapper,
            """
            {
              "passwordRef": { "id": "persist-pass" }
            }
            """,
            OoxmlEncryptionSpec.class);
    OoxmlSignatureSpec signature =
        readJson(
            mapper,
            """
            {
              "pkcs12Path": "keys/signing.p12",
              "keystorePasswordRef": { "id": "store-pass" }
            }
            """,
            OoxmlSignatureSpec.class);
    OoxmlEncryptionSpec explicitEncryption =
        readOmitted(
            OoxmlEncryptionSpec.class,
            new SecretReference("persist-pass"),
            ExcelOoxmlWriteCipher.AES_192,
            ExcelOoxmlWriteHash.SHA_384);
    OoxmlSignatureSpec createdSignatureWithNullOptionals =
        readOmitted(
            OoxmlSignatureSpec.class,
            "keys/signing.p12",
            new SecretReference("store-pass"),
            Optional.of(new SecretReference("key-pass")),
            null,
            ExcelOoxmlSignatureDigestAlgorithm.SHA512,
            null);

    assertTrue(execution.isDefault());
    assertTrue(journal.isDefault());
    assertTrue(calculation.isDefault());
    assertTrue(formulaEnvironment.isEmpty());
    assertEquals(ExcelOoxmlWriteCipher.AES_256, encryption.cipher());
    assertEquals(ExcelOoxmlWriteHash.SHA_512, encryption.hash());
    assertEquals(ExcelOoxmlWriteCipher.AES_192, explicitEncryption.cipher());
    assertEquals(ExcelOoxmlWriteHash.SHA_384, explicitEncryption.hash());
    assertEquals(Optional.empty(), signature.keyPasswordRef());
    assertEquals(ExcelOoxmlSignatureDigestAlgorithm.SHA256, signature.digestAlgorithm());
    assertEquals(Optional.empty(), signature.alias());
    assertEquals(Optional.empty(), signature.description());
    assertEquals(
        Optional.of(new SecretReference("key-pass")),
        createdSignatureWithNullOptionals.keyPasswordRef());
    assertEquals(
        ExcelOoxmlSignatureDigestAlgorithm.SHA512,
        createdSignatureWithNullOptionals.digestAlgorithm());
    assertEquals(Optional.empty(), createdSignatureWithNullOptionals.alias());
    assertEquals(Optional.empty(), createdSignatureWithNullOptionals.description());
  }

  @Test
  void autofilterAndExecutionInputsCoverExplicitAndStrictBranches() {
    AutofilterFilterCriterionSpec.Values criterion =
        new AutofilterFilterCriterionSpec.Values(List.of("Ready"), false);
    AutofilterSortConditionSpec valueCondition =
        new AutofilterSortConditionSpec.Value(new RangeRef("A1:A4"), true);
    AutofilterFilterColumnSpec visibleButtonColumn =
        new AutofilterFilterColumnSpec(2L, true, criterion);
    AutofilterFilterColumnSpec hiddenButtonColumn =
        new AutofilterFilterColumnSpec(2L, false, criterion);
    AutofilterSortStateSpec defaultedSortState =
        AutofilterSortStateSpec.withoutSortMethod(
            new RangeRef("A1:B4"), false, false, List.of(valueCondition));
    AutofilterSortStateSpec explicitSortState =
        new AutofilterSortStateSpec(
            new RangeRef("A1:B4"),
            true,
            true,
            Optional.of(ExcelAutofilterSortMethod.PINYIN),
            List.of(valueCondition));
    ExecutionModeSpec customMode = ExecutionModeSpec.eventRead();
    ExecutionModeSpec writeOnlyCustomMode = ExecutionModeSpec.streamingWrite();
    ExecutionPolicySpec customPolicy = ExecutionPolicySpec.mode(customMode);
    ExecutionPolicySpec customJournalPolicy =
        ExecutionPolicySpec.journal(new ExecutionJournalSpec(ExecutionJournalLevel.VERBOSE));
    SheetDisplaySpec explicitDisplay = new SheetDisplaySpec(false, true, false, true, false);
    SheetOutlineSummarySpec explicitOutline = new SheetOutlineSummarySpec(false, true);
    SheetDefaultsSpec explicitSheetDefaults =
        new SheetDefaultsSpec(Integer.valueOf(12), Double.valueOf(18.5d));
    assertTrue(visibleButtonColumn.showButton());
    assertFalse(hiddenButtonColumn.showButton());
    assertInstanceOf(AutofilterSortConditionSpec.Value.class, valueCondition);
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
    assertTrue(readOmitted(AutofilterFilterColumnSpec.class, 2L, null, criterion).showButton());
    AutofilterSortStateSpec defaultedWireSortState =
        readOmitted(
            AutofilterSortStateSpec.class,
            new RangeRef("A1:B4"),
            null,
            null,
            Optional.empty(),
            List.of(valueCondition));
    assertFalse(defaultedWireSortState.caseSensitive());
    assertFalse(defaultedWireSortState.columnSort());
    assertEquals(Optional.empty(), defaultedWireSortState.sortMethod());
  }

  @Test
  void residualWireDefaultsHandleOmittedFields() {
    CommentSpec comment =
        readOmitted(CommentSpec.class, text("Owner note"), "Ada", null, null, null);
    AutofilterSortConditionSpec valueCondition =
        new AutofilterSortConditionSpec.Value(new RangeRef("A1:A4"), true);
    AutofilterSortStateSpec defaultedSortState =
        readOmitted(
            AutofilterSortStateSpec.class,
            new RangeRef("A1:B4"),
            false,
            true,
            null,
            List.of(valueCondition));
    WorkbookProtectionSpec protection =
        readOmitted(WorkbookProtectionSpec.class, null, null, null, null, null);
    PivotTableSpec.DataField dataField =
        readOmitted(
            PivotTableSpec.DataField.class,
            "Amount",
            ExcelPivotDataConsolidateFunction.SUM,
            null,
            null);
    TableColumnSpec tableColumn = readOmitted(TableColumnSpec.class, 0, null, null, null, null);
    TableColumnSpec explicitFactoryColumn =
        readOmitted(TableColumnSpec.class, 1, "owner", "Total", " SUM ", "=[@Amount]");
    DataValidationSpec validation =
        readOmitted(
            DataValidationSpec.class,
            new DataValidationRuleSpec.ExplicitList(List.of("Ready")),
            null,
            null,
            null,
            null);

    assertFalse(comment.visible());
    assertEquals(Optional.empty(), comment.runs());
    assertEquals(Optional.empty(), comment.anchor());
    assertEquals(Optional.empty(), defaultedSortState.sortMethod());
    assertFalse(protection.structureLocked());
    assertFalse(protection.windowsLocked());
    assertFalse(protection.revisionsLocked());
    assertEquals(Optional.empty(), protection.workbookPasswordRef());
    assertEquals(Optional.empty(), protection.revisionsPasswordRef());
    assertEquals("Amount", dataField.displayName());
    assertEquals(Optional.empty(), dataField.valueFormat());
    assertEquals("", tableColumn.uniqueName());
    assertEquals("", tableColumn.totalsRowLabel());
    assertEquals("", tableColumn.totalsRowFunction());
    assertEquals("", tableColumn.calculatedColumnFormula());
    assertEquals("owner", explicitFactoryColumn.uniqueName());
    assertEquals("sum", explicitFactoryColumn.totalsRowFunction());
    assertFalse(validation.allowBlank());
    assertFalse(validation.suppressDropDownArrow());
    assertEquals(Optional.empty(), validation.prompt());
    assertEquals(Optional.empty(), validation.errorAlert());
  }

  @Test
  void tableWireConstructorsCoverExplicitAndDefaultedBranches() {
    TableSpec defaultedTable =
        TableSpec.withDefaultMetadata(
            new TableName("Budget"),
            new SheetName("Budget"),
            new RangeRef("A1:B3"),
            false,
            new TableStyleSpec.None());
    TableSpec explicitTable =
        new TableSpec(
            new TableName("StyledBudget"),
            new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
            new RangeRef("A1:C3"),
            true,
            false,
            new TableStyleSpec.Named("TableStyleMedium2", true, false, true, false),
            text("Quarterly budget"),
            true,
            true,
            true,
            "HeaderStyle",
            "DataStyle",
            "TotalsStyle",
            List.of(new TableColumnSpec(0, "owner", "total", "sum", "=[@Amount]")));
    TableSpec explicitAutofilterTable =
        new TableSpec(
            new TableName("VisibleAutofilterBudget"),
            new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
            new RangeRef("A1:C3"),
            false,
            true,
            new TableStyleSpec.None(),
            text(""),
            false,
            false,
            false,
            "",
            "",
            "",
            List.of());
    TableSpec explicitTotalsAndNoAutofilterTable =
        new TableSpec(
            new TableName("TotalsWithoutAutofilterBudget"),
            new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
            new RangeRef("A1:C3"),
            true,
            false,
            new TableStyleSpec.None(),
            text(""),
            false,
            false,
            false,
            "",
            "",
            "",
            List.of());
    TableEntryReport defaultedTableReport =
        new TableEntryReport(
            new TableLabel("BudgetTable"),
            new SheetLabel("Budget"),
            new RangeLabel("A1:B3"),
            new TableEntryReport.Structure(
                1, 0, List.of("Owner"), List.of(new TableColumnReport(0L, "Owner"))),
            new TableStyleReport.None(),
            new TableEntryReport.Behavior(true, false, false, false),
            new TableEntryReport.Presentation(
                Optional.empty(), Optional.empty(), Optional.empty(), Optional.empty()));

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
    assertEquals(Optional.empty(), defaultedTableReport.presentation().comment());
    assertEquals(Optional.empty(), defaultedTableReport.presentation().headerRowCellStyle());
    assertEquals(Optional.empty(), defaultedTableReport.presentation().dataCellStyle());
    assertEquals(Optional.empty(), defaultedTableReport.presentation().totalsRowCellStyle());
    TableSpec defaultedWireTable =
        readOmitted(
            TableSpec.class,
            new TableName("Budget"),
            new SheetName("Budget"),
            new RangeRef("A1:B3"),
            null,
            null,
            new TableStyleSpec.None(),
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null);
    assertFalse(defaultedWireTable.showTotalsRow());
    assertTrue(defaultedWireTable.hasAutofilter());
    assertEquals(text(""), defaultedWireTable.comment());
    assertFalse(defaultedWireTable.published());
    assertFalse(defaultedWireTable.insertRow());
    assertFalse(defaultedWireTable.insertRowShift());
    assertEquals("", defaultedWireTable.headerRowCellStyle());
    assertEquals("", defaultedWireTable.dataCellStyle());
    assertEquals("", defaultedWireTable.totalsRowCellStyle());
    assertEquals(List.of(), defaultedWireTable.columns());
  }

  @Test
  void validationChartPrintAndSignatureWireConstructorsCoverExplicitAndDefaultedBranches() {
    JsonMapper mapper = dev.erst.gridgrind.json.codec.GridGrindJsonMapperSupport.JSON_MAPPER;
    DataValidationErrorAlertSpec explicitAlert =
        new DataValidationErrorAlertSpec(
            ExcelDataValidationErrorStyle.STOP,
            text("Invalid"),
            text("Choose an allowed value."),
            true);
    DataValidationErrorAlertSpec hiddenAlert =
        new DataValidationErrorAlertSpec(
            ExcelDataValidationErrorStyle.WARNING,
            text("Caution"),
            text("Review the warning."),
            false);
    ChartSeriesSpec series =
        new ChartSeriesSpec(
            new ChartTitleSpec.None(),
            new ChartDataSourceSpec.StringLiteral(List.of("Jan")),
            new ChartDataSourceSpec.NumericLiteral(List.of(10.0d)),
            Optional.of(true),
            Optional.of(ExcelChartMarkerStyle.CIRCLE),
            Optional.of((short) 6),
            Optional.empty());
    ChartSpec defaultedChart =
        ChartSpec.withStandardDisplay(
            "BudgetChart",
            twoCellAnchor(),
            List.of(new ChartPlotSpec.Line(false, ExcelChartGrouping.STANDARD, List.of(series))));
    ChartSpec explicitChart =
        new ChartSpec(
            "BudgetChartExplicit",
            twoCellAnchor(),
            new ChartTitleSpec.Text(text("Budget")),
            new ChartLegendSpec.Hidden(),
            ExcelChartDisplayBlanksAs.SPAN,
            false,
            List.of(new ChartPlotSpec.Line(false, ExcelChartGrouping.STANDARD, List.of(series))));
    ChartAxisSpec defaultedAxis =
        ChartAxisSpec.visible(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO);
    ChartAxisSpec hiddenAxis =
        new ChartAxisSpec(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            false);
    PrintSetupSpec defaultedPrintSetup = PrintSetupSpec.defaults();
    PrintSetupSpec explicitPrintSetup =
        new PrintSetupSpec(
            new PrintMarginsSpec(0.6d, 0.6d, 0.7d, 0.7d, 0.2d, 0.2d),
            true,
            true,
            true,
            Integer.valueOf(9),
            true,
            true,
            Integer.valueOf(2),
            true,
            Integer.valueOf(7),
            List.of(3),
            List.of(4));
    SignatureLineSpec defaultedSignatureLine =
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
            SignatureLineSpec.class);

    assertTrue(explicitAlert.showErrorBox());
    assertFalse(hiddenAlert.showErrorBox());
    assertTrue(defaultedChart.title() instanceof ChartTitleSpec.None);
    assertEquals(
        new ChartLegendSpec.Visible(ExcelChartLegendPosition.RIGHT), defaultedChart.legend());
    assertEquals(ExcelChartDisplayBlanksAs.GAP, defaultedChart.displayBlanksAs());
    assertTrue(defaultedChart.plotOnlyVisibleCells());
    assertTrue(explicitChart.legend() instanceof ChartLegendSpec.Hidden);
    assertEquals(ExcelChartDisplayBlanksAs.SPAN, explicitChart.displayBlanksAs());
    assertFalse(explicitChart.plotOnlyVisibleCells());
    assertTrue(defaultedAxis.visible());
    assertFalse(hiddenAxis.visible());
    assertEquals(PrintSetupSpec.defaults(), defaultedPrintSetup);
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
        tools.jackson.core.JacksonException.class,
        () ->
            readOmitted(
                ChartSpec.class,
                "BudgetChart",
                twoCellAnchor(),
                null,
                null,
                null,
                null,
                List.of(
                    new ChartPlotSpec.Line(false, ExcelChartGrouping.STANDARD, List.of(series)))));
    assertThrows(
        tools.jackson.core.JacksonException.class,
        () ->
            readOmitted(
                ChartAxisSpec.class,
                ExcelChartAxisKind.CATEGORY,
                ExcelChartAxisPosition.BOTTOM,
                ExcelChartAxisCrosses.AUTO_ZERO,
                null));
    assertThrows(
        tools.jackson.core.JacksonException.class,
        () ->
            readOmitted(
                PrintSetupSpec.class,
                null,
                null,
                null,
                null,
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
  void convenienceConstructorsCoverDefaultPlanAndReportBranches() {
    WorkbookPlan defaultPlan =
        WorkbookPlan.identified(
            GridGrindProtocolVersion.current(),
            new PlanId("budget-defaults"),
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of());
    WorkbookPlan explicitExecutionPlan =
        WorkbookPlan.identified(
            GridGrindProtocolVersion.current(),
            new PlanId("budget-execution"),
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.journal(new ExecutionJournalSpec(ExecutionJournalLevel.VERBOSE)),
            FormulaEnvironmentSpec.empty(),
            List.of());
    WorkbookPlan wireDefaultedPlan =
        readOmitted(
            WorkbookPlan.class,
            GridGrindProtocolVersion.current(),
            Optional.empty(),
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            null,
            null,
            List.of());
    TableColumnReport tableColumn = new TableColumnReport(4L, "Owner");
    AutofilterSortConditionReport conditionReport =
        new AutofilterSortConditionReport.CellColor(
            new RangeLabel("B2:B9"), false, CellColorReport.rgb(new RgbHex("#aabbcc")));
    AutofilterSortConditionReport createdConditionReport =
        new AutofilterSortConditionReport.Icon(new RangeLabel("C2:C9"), true, 2);
    AutofilterSortStateReport stateReport =
        AutofilterSortStateReport.withoutSortMethod(
            new RangeLabel("A1:F9"), false, true, List.of(conditionReport));
    AutofilterSortStateReport createdStateReport =
        AutofilterSortStateReport.withoutSortMethod(
            new RangeLabel("A1:F9"), true, false, List.of(createdConditionReport));
    BorderSidePatch borderSide =
        new BorderSidePatch(Optional.of(ExcelBorderStyle.THIN), Optional.empty());
    DifferentialBorderPatch borderInput =
        new DifferentialBorderPatch(
            Optional.of(borderSide),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    DifferentialBorderReport borderReport =
        new DifferentialBorderReport(
            Optional.empty(),
            Optional.of(
                new DifferentialBorderSideReport(
                    Optional.of(ExcelBorderStyle.THIN), Optional.empty())),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    AutofilterSortConditionSpec explicitSortCondition =
        new AutofilterSortConditionSpec.Icon(new RangeRef("D2:D9"), false, 1);
    AutofilterFilterColumnSpec explicitFilterColumn =
        new AutofilterFilterColumnSpec(
            2L, false, new AutofilterFilterCriterionSpec.Values(List.of("Queued"), true));
    AutofilterSortStateSpec explicitSortState =
        new AutofilterSortStateSpec(
            new RangeRef("A1:D9"),
            true,
            true,
            Optional.of(ExcelAutofilterSortMethod.PINYIN),
            List.of(explicitSortCondition));
    AutofilterSortConditionReport explicitConditionReport =
        new AutofilterSortConditionReport.Icon(new RangeLabel("D2:D9"), false, 1);
    AutofilterSortStateReport explicitStateReport =
        new AutofilterSortStateReport(
            new RangeLabel("A1:D9"),
            true,
            true,
            Optional.of(ExcelAutofilterSortMethod.PINYIN),
            List.of(explicitConditionReport));
    SheetDisplaySpec explicitDisplay = new SheetDisplaySpec(false, true, false, true, true);
    SheetOutlineSummarySpec explicitOutline = new SheetOutlineSummarySpec(false, true);
    SheetDefaultsSpec explicitSheetDefaults = new SheetDefaultsSpec(18, 24.5d);
    SheetPresentationSpec explicitPresentation =
        new SheetPresentationSpec(
            explicitDisplay,
            Optional.of(ColorSpec.rgb(new RgbHex("#aabbcc"))),
            explicitOutline,
            explicitSheetDefaults,
            List.of(
                new IgnoredErrorSpec(
                    new RangeRef("A1:A3"), List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT))));
    FormulaEnvironmentSpec explicitEnvironment =
        new FormulaEnvironmentSpec(
            List.of(new FormulaExternalWorkbookSpec("Rates.xlsx", "tmp/rates.xlsx")),
            FormulaMissingWorkbookPolicy.USE_CACHED_VALUE,
            List.of());

    assertTrue(defaultPlan.execution().isDefault());
    assertTrue(defaultPlan.formulaEnvironment().isEmpty());
    assertTrue(wireDefaultedPlan.execution().isDefault());
    assertTrue(wireDefaultedPlan.formulaEnvironment().isEmpty());
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
    assertEquals(ExcelBorderStyle.THIN, borderReport.top().orElseThrow().style().orElseThrow());
    assertThrows(
        IllegalArgumentException.class,
        () -> new DifferentialBorderSideReport(Optional.empty(), Optional.empty()));
    assertEquals(
        Optional.of(CellColorReport.rgb(new RgbHex("#aabbcc"))),
        new DifferentialBorderSideReport(
                Optional.empty(), Optional.of(CellColorReport.rgb(new RgbHex("#aabbcc"))))
            .color());
    assertEquals(
        Optional.of(ExcelBorderStyle.NONE),
        new DifferentialBorderSideReport(Optional.of(ExcelBorderStyle.NONE), Optional.empty())
            .style());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DifferentialBorderSideReport(
                Optional.of(ExcelBorderStyle.NONE),
                Optional.of(CellColorReport.rgb(new RgbHex("#aabbcc")))));
    assertInstanceOf(AutofilterSortConditionSpec.Icon.class, explicitSortCondition);
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
    ChartSeriesSpec explicitSeries =
        new ChartSeriesSpec(
            new ChartTitleSpec.None(),
            new ChartDataSourceSpec.StringLiteral(List.of("Jan", "Feb")),
            new ChartDataSourceSpec.NumericLiteral(List.of(10.0d, 18.0d)),
            Optional.of(true),
            Optional.of(ExcelChartMarkerStyle.DIAMOND),
            Optional.of((short) 8),
            Optional.of(12L));
    ChartSeriesSpec convenienceSeries =
        ChartSeriesSpec.untitled(
            new ChartDataSourceSpec.Reference("Categories"),
            new ChartDataSourceSpec.Reference("Values"),
            Optional.of(false),
            Optional.of(ExcelChartMarkerStyle.CIRCLE),
            Optional.of((short) 6),
            Optional.empty());
    ChartAxisSpec defaultAxis =
        ChartAxisSpec.visible(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO);
    List<ChartAxisSpec> explicitAxes =
        List.of(
            defaultAxis,
            new ChartAxisSpec(
                ExcelChartAxisKind.VALUE,
                ExcelChartAxisPosition.LEFT,
                ExcelChartAxisCrosses.AUTO_ZERO,
                false));
    ChartPlotSpec.Area explicitArea =
        new ChartPlotSpec.Area(
            true, ExcelChartGrouping.PERCENT_STACKED, explicitAxes, List.of(explicitSeries));
    ChartPlotSpec.Area convenienceArea =
        new ChartPlotSpec.Area(false, ExcelChartGrouping.STANDARD, List.of(explicitSeries));
    ChartPlotSpec.Area3D explicitArea3D =
        new ChartPlotSpec.Area3D(
            false,
            ExcelChartGrouping.STACKED,
            Optional.of(24),
            explicitAxes,
            List.of(explicitSeries));
    ChartPlotSpec.Area3D convenienceArea3D =
        new ChartPlotSpec.Area3D(
            false, ExcelChartGrouping.STANDARD, Optional.of(16), List.of(explicitSeries));
    ChartPlotSpec.Bar3D explicitBar3D =
        new ChartPlotSpec.Bar3D(
            false,
            ExcelChartBarDirection.BAR,
            ExcelChartBarGrouping.PERCENT_STACKED,
            Optional.of(12),
            Optional.of(64),
            Optional.of(ExcelChartBarShape.CONE),
            explicitAxes,
            List.of(explicitSeries));
    ChartPlotSpec.Bar3D convenienceBar3D =
        new ChartPlotSpec.Bar3D(
            false,
            ExcelChartBarDirection.COLUMN,
            ExcelChartBarGrouping.CLUSTERED,
            Optional.of(18),
            Optional.of(90),
            Optional.of(ExcelChartBarShape.BOX),
            List.of(explicitSeries));
    ChartPlotSpec.Line explicitLine =
        new ChartPlotSpec.Line(
            true, ExcelChartGrouping.STACKED, explicitAxes, List.of(explicitSeries));
    ChartPlotSpec.Line convenienceLine =
        new ChartPlotSpec.Line(false, ExcelChartGrouping.STANDARD, List.of(explicitSeries));
    ChartPlotSpec.Line3D explicitLine3D =
        new ChartPlotSpec.Line3D(
            true,
            ExcelChartGrouping.PERCENT_STACKED,
            Optional.of(32),
            explicitAxes,
            List.of(explicitSeries));
    ChartPlotSpec.Line3D convenienceLine3D =
        new ChartPlotSpec.Line3D(
            false, ExcelChartGrouping.STANDARD, Optional.of(8), List.of(explicitSeries));
    ChartPlotSpec.Radar explicitRadar =
        new ChartPlotSpec.Radar(
            true, ExcelChartRadarStyle.MARKER, explicitAxes, List.of(explicitSeries));
    ChartPlotSpec.Radar convenienceRadar =
        new ChartPlotSpec.Radar(false, ExcelChartRadarStyle.STANDARD, List.of(explicitSeries));
    ChartPlotSpec.Scatter explicitScatter =
        new ChartPlotSpec.Scatter(
            true, ExcelChartScatterStyle.SMOOTH, explicitAxes, List.of(explicitSeries));
    ChartPlotSpec.Scatter convenienceScatter =
        new ChartPlotSpec.Scatter(false, ExcelChartScatterStyle.LINE, List.of(explicitSeries));
    ChartPlotSpec.Surface explicitSurface =
        new ChartPlotSpec.Surface(true, true, explicitAxes, List.of(explicitSeries));
    ChartPlotSpec.Surface convenienceSurface =
        new ChartPlotSpec.Surface(false, false, List.of(explicitSeries));
    ChartPlotSpec.Surface3D explicitSurface3D =
        new ChartPlotSpec.Surface3D(true, true, explicitAxes, List.of(explicitSeries));
    ChartPlotSpec.Surface3D convenienceSurface3D =
        new ChartPlotSpec.Surface3D(false, false, List.of(explicitSeries));
    ChartSpec createdChart =
        new ChartSpec(
            "CreatedChart",
            twoCellAnchor(),
            new ChartTitleSpec.None(),
            new ChartLegendSpec.Visible(ExcelChartLegendPosition.RIGHT),
            ExcelChartDisplayBlanksAs.GAP,
            true,
            List.of(explicitArea));
    ChartSpec convenienceChart =
        ChartSpec.withStandardDisplay("ConvenienceChart", twoCellAnchor(), List.of(explicitLine));

    assertTrue(defaultAxis.visible());
    assertTrue(convenienceSeries.title() instanceof ChartTitleSpec.None);
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
    assertTrue(createdChart.title() instanceof ChartTitleSpec.None);
    assertEquals(
        new ChartLegendSpec.Visible(ExcelChartLegendPosition.RIGHT), createdChart.legend());
    assertEquals(ExcelChartDisplayBlanksAs.GAP, createdChart.displayBlanksAs());
    assertTrue(createdChart.plotOnlyVisibleCells());
    assertTrue(convenienceChart.title() instanceof ChartTitleSpec.None);
  }

  @Test
  void chartPlotWireDefaultsHandleOmittedBooleansAndOptionals() {
    ChartSeriesSpec series = sampleSeries();
    List<ChartAxisSpec> axes = sampleAxes();
    ChartPlotSpec.Area3D area3D =
        readOmitted(
            ChartPlotSpec.Area3D.class,
            null,
            ExcelChartGrouping.STANDARD,
            null,
            axes,
            List.of(series));
    ChartPlotSpec.Bar3D bar3D =
        readOmitted(
            ChartPlotSpec.Bar3D.class,
            null,
            ExcelChartBarDirection.COLUMN,
            ExcelChartBarGrouping.CLUSTERED,
            null,
            null,
            null,
            axes,
            List.of(series));
    ChartPlotSpec.Doughnut doughnut =
        readOmitted(ChartPlotSpec.Doughnut.class, null, null, null, List.of(series));
    ChartPlotSpec.Line line =
        readOmitted(
            ChartPlotSpec.Line.class, null, ExcelChartGrouping.STANDARD, axes, List.of(series));
    ChartPlotSpec.Line3D line3D =
        readOmitted(
            ChartPlotSpec.Line3D.class,
            null,
            ExcelChartGrouping.STANDARD,
            null,
            axes,
            List.of(series));
    ChartPlotSpec.Pie pie = readOmitted(ChartPlotSpec.Pie.class, null, null, List.of(series));
    ChartPlotSpec.Pie3D pie3D = readOmitted(ChartPlotSpec.Pie3D.class, null, List.of(series));
    ChartPlotSpec.Radar radar =
        readOmitted(
            ChartPlotSpec.Radar.class, null, ExcelChartRadarStyle.STANDARD, axes, List.of(series));
    ChartPlotSpec.Scatter scatter =
        readOmitted(
            ChartPlotSpec.Scatter.class, null, ExcelChartScatterStyle.LINE, axes, List.of(series));
    ChartPlotSpec.Surface surface =
        readOmitted(ChartPlotSpec.Surface.class, null, null, axes, List.of(series));
    ChartPlotSpec.Surface3D surface3D =
        readOmitted(ChartPlotSpec.Surface3D.class, null, null, axes, List.of(series));

    assertFalse(area3D.varyColors());
    assertEquals(Optional.empty(), area3D.gapDepth());
    assertFalse(bar3D.varyColors());
    assertEquals(Optional.empty(), bar3D.gapDepth());
    assertEquals(Optional.empty(), bar3D.gapWidth());
    assertEquals(Optional.empty(), bar3D.shape());
    assertFalse(doughnut.varyColors());
    assertEquals(Optional.empty(), doughnut.firstSliceAngle());
    assertEquals(Optional.empty(), doughnut.holeSize());
    assertFalse(line.varyColors());
    assertFalse(line3D.varyColors());
    assertEquals(Optional.empty(), line3D.gapDepth());
    assertFalse(pie.varyColors());
    assertEquals(Optional.empty(), pie.firstSliceAngle());
    assertFalse(pie3D.varyColors());
    assertFalse(radar.varyColors());
    assertFalse(scatter.varyColors());
    assertFalse(surface.varyColors());
    assertFalse(surface.wireframe());
    assertFalse(surface3D.varyColors());
    assertFalse(surface3D.wireframe());
  }

  @Test
  void creatorsRequireExplicitWireValues() {
    JsonMapper mapper = dev.erst.gridgrind.json.codec.GridGrindJsonMapperSupport.JSON_MAPPER;
    PrintSetupSpec printSetup =
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
            PrintSetupSpec.class);
    PrintLayoutSpec printLayout =
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
            PrintLayoutSpec.class);
    PrintLayoutSpec conveniencePrintLayout =
        PrintLayoutSpec.withDefaultSetup(
            new PrintAreaSpec.None(),
            ExcelPrintOrientation.LANDSCAPE,
            new PrintScalingSpec.Automatic(),
            new PrintTitleRowsSpec.None(),
            new PrintTitleColumnsSpec.None(),
            HeaderFooterTextSpec.blank(),
            HeaderFooterTextSpec.blank());
    PrintLayoutSpec createdPrintLayout =
        PrintLayoutSpec.withDefaultSetup(
            new PrintAreaSpec.None(),
            ExcelPrintOrientation.LANDSCAPE,
            new PrintScalingSpec.Automatic(),
            new PrintTitleRowsSpec.None(),
            new PrintTitleColumnsSpec.None(),
            HeaderFooterTextSpec.blank(),
            HeaderFooterTextSpec.blank());
    SignatureLineSpec signatureLine =
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
            SignatureLineSpec.class);
    TableEntryReport tableReport =
        mapper.readValue(
            """
            {
              "name": "BudgetTable",
              "sheetName": "Budget",
              "range": "A1:B4",
              "structure": {
                "headerRowCount": 1,
                "totalsRowCount": 0,
                "columnNames": ["Owner"],
                "columns": [{ "id": 0, "name": "Owner" }]
              },
              "style": { "type": "NONE" },
              "behavior": {
                "hasAutofilter": true,
                "published": false,
                "insertRow": false,
                "insertRowShift": false
              },
              "presentation": {
                "comment": "",
                "headerRowCellStyle": "",
                "dataCellStyle": "",
                "totalsRowCellStyle": ""
              }
            }
            """,
            TableEntryReport.class);
    TableEntryReport defaultedBehaviorTableReport =
        mapper.readValue(
            """
            {
              "name": "BudgetTable",
              "sheetName": "Budget",
              "range": "A1:B4",
              "structure": {
                "headerRowCount": 1,
                "totalsRowCount": 0,
                "columnNames": ["Owner"],
                "columns": [{ "id": 0, "name": "Owner" }]
              },
              "style": { "type": "NONE" }
            }
            """,
            TableEntryReport.class);
    ExecutionJournal journal =
        mapper.readValue(
            """
            {
              "level": "NORMAL",
              "source": { "type": "NEW" },
              "validation": { "status": "NOT_STARTED" },
              "inputResolution": { "status": "NOT_STARTED" },
              "open": { "status": "NOT_STARTED" },
              "calculation": {
                "preflight": { "status": "NOT_STARTED" },
                "execution": { "status": "NOT_STARTED" }
              },
              "persistencePhase": { "status": "NOT_STARTED" },
              "close": { "status": "NOT_STARTED" },
              "steps": [],
              "outcome": {
                "status": "SUCCEEDED",
                "plannedStepCount": 0,
                "completedStepCount": 0,
                "durationMillis": 0
              }
            }
            """,
            ExecutionJournal.class);

    assertEquals(new PrintMarginsSpec(0.6d, 0.6d, 0.7d, 0.7d, 0.3d, 0.3d), printSetup.margins());
    assertEquals(ExcelPrintOrientation.PORTRAIT, printLayout.orientation());
    assertTrue(printLayout.setup().printGridlines());
    assertEquals(PrintSetupSpec.defaults(), conveniencePrintLayout.setup());
    assertEquals(PrintSetupSpec.defaults(), createdPrintLayout.setup());
    assertTrue(signatureLine.allowComments());
    assertEquals(Optional.empty(), tableReport.presentation().comment());
    assertEquals(Optional.empty(), tableReport.presentation().headerRowCellStyle());
    assertEquals(Optional.empty(), tableReport.presentation().dataCellStyle());
    assertEquals(Optional.empty(), tableReport.presentation().totalsRowCellStyle());
    assertEquals(
        new TableEntryReport.Behavior(false, false, false, false),
        defaultedBehaviorTableReport.behavior());
    assertEquals(ExecutionJournalLevel.NORMAL, journal.level());
  }

  private static DrawingAnchorSpec.TwoCell twoCellAnchor() {
    return new DrawingAnchorSpec.TwoCell(
        new DrawingMarkerSpec(1, 0, 1, 0),
        new DrawingMarkerSpec(5, 0, 10, 0),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  private static TextSourceSpec text(String value) {
    return TextSourceSpec.inline(value);
  }

  private static ChartSeriesSpec sampleSeries() {
    return ChartSeriesSpec.untitled(
        new ChartDataSourceSpec.StringLiteral(List.of("Q1")),
        new ChartDataSourceSpec.NumericLiteral(List.of(1.0d)),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
  }

  private static List<ChartAxisSpec> sampleAxes() {
    return List.of(
        new ChartAxisSpec(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisSpec(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static <T> T readJson(JsonMapper mapper, String json, Class<T> targetType) {
    return mapper.readValue(json, targetType);
  }

  private static <T> T readOmitted(Class<T> type, Object... values) {
    var mapper = dev.erst.gridgrind.json.codec.GridGrindJsonMapperSupport.JSON_MAPPER;
    var fields = mapper.createObjectNode();
    var variant = type.getAnnotation(dev.erst.gridgrind.core.metadata.Variant.class);
    if (variant != null) {
      fields.put("type", variant.id());
    }
    var components = type.getRecordComponents();
    assertEquals(
        components.length, values.length, "Fixture supplies a position for every component");
    for (int index = 0; index < values.length; index++) {
      if (values[index] != null) {
        fields.set(
            components[index].getName(),
            mapper.readTree(
                mapper
                    .writerFor(mapper.constructType(components[index].getGenericType()))
                    .writeValueAsBytes(values[index])));
      }
    }
    return mapper.readValue(mapper.writeValueAsBytes(fields), type);
  }
}
