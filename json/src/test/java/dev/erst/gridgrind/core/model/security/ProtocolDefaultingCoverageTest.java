package dev.erst.gridgrind.core.model.security;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.chart.ChartAxisInput;
import dev.erst.gridgrind.core.model.chart.ChartDataSourceInput;
import dev.erst.gridgrind.core.model.chart.ChartInput;
import dev.erst.gridgrind.core.model.chart.ChartLegendInput;
import dev.erst.gridgrind.core.model.chart.ChartPlotInput;
import dev.erst.gridgrind.core.model.chart.ChartSeriesInput;
import dev.erst.gridgrind.core.model.chart.ChartTitleInput;
import dev.erst.gridgrind.core.model.drawing.CommentInput;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorInput;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerInput;
import dev.erst.gridgrind.core.model.drawing.SignatureLineInput;
import dev.erst.gridgrind.core.model.formula.CalculationPolicyInput;
import dev.erst.gridgrind.core.model.formula.CalculationStrategyInput;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentInput;
import dev.erst.gridgrind.core.model.formula.FormulaExternalWorkbookInput;
import dev.erst.gridgrind.core.model.formula.FormulaMissingWorkbookPolicy;
import dev.erst.gridgrind.core.model.formula.FormulaUdfFunctionInput;
import dev.erst.gridgrind.core.model.layout.SheetDefaultsInput;
import dev.erst.gridgrind.core.model.layout.SheetDisplayInput;
import dev.erst.gridgrind.core.model.layout.SheetOutlineSummaryInput;
import dev.erst.gridgrind.core.model.layout.SheetPresentationInput;
import dev.erst.gridgrind.core.model.pivot.PivotTableInput;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextInput;
import dev.erst.gridgrind.core.model.print.PrintAreaInput;
import dev.erst.gridgrind.core.model.print.PrintLayoutInput;
import dev.erst.gridgrind.core.model.print.PrintMarginsInput;
import dev.erst.gridgrind.core.model.print.PrintScalingInput;
import dev.erst.gridgrind.core.model.print.PrintSetupInput;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsInput;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsInput;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionInput;
import dev.erst.gridgrind.core.model.style.BorderSideInput;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.ColorInput;
import dev.erst.gridgrind.core.model.style.DifferentialBorderInput;
import dev.erst.gridgrind.core.model.style.DifferentialBorderReport;
import dev.erst.gridgrind.core.model.style.DifferentialBorderSideReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnInput;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionInput;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionInput;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateInput;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateReport;
import dev.erst.gridgrind.core.model.table.TableColumnInput;
import dev.erst.gridgrind.core.model.table.TableColumnReport;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.table.TableInput;
import dev.erst.gridgrind.core.model.table.TableStyleInput;
import dev.erst.gridgrind.core.model.table.TableStyleReport;
import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertInput;
import dev.erst.gridgrind.core.model.validation.DataValidationInput;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptInput;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleInput;
import dev.erst.gridgrind.core.model.validation.IgnoredErrorInput;
import dev.erst.gridgrind.core.plan.ExecutionJournalInput;
import dev.erst.gridgrind.core.plan.ExecutionModeInput;
import dev.erst.gridgrind.core.plan.ExecutionPolicyInput;
import dev.erst.gridgrind.core.plan.GridGrindProtocolVersion;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.RequestDoctorReport;
import dev.erst.gridgrind.core.result.ExecutionJournal;
import dev.erst.gridgrind.core.result.ExecutionJournalLevel;
import dev.erst.gridgrind.core.source.SecretReference;
import dev.erst.gridgrind.core.source.TextSourceInput;
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
    CalculationPolicyInput calculation =
        new CalculationPolicyInput(new CalculationStrategyInput.DoNotCalculate(), false);
    ExecutionModeInput mode = ExecutionModeInput.fullXssf();
    ExecutionJournalInput journalInput = new ExecutionJournalInput(ExecutionJournalLevel.NORMAL);
    FormulaEnvironmentInput environment = FormulaEnvironmentInput.empty();
    SheetOutlineSummaryInput outlineSummary = SheetOutlineSummaryInput.defaults();
    SheetDisplayInput display = SheetDisplayInput.defaults();
    SheetDefaultsInput sheetDefaults = SheetDefaultsInput.defaults();
    HeaderFooterTextInput headerFooter = HeaderFooterTextInput.blank();
    SheetPresentationInput presentation = SheetPresentationInput.defaults();
    WorkbookProtectionInput protection =
        new WorkbookProtectionInput(
            false, false, false, java.util.Optional.empty(), java.util.Optional.empty());
    AutofilterSortConditionInput sortCondition =
        new AutofilterSortConditionInput.Value(new RangeRef("A1:A2"), true);
    AutofilterFilterColumnInput filterColumn =
        new AutofilterFilterColumnInput(
            1L, true, new AutofilterFilterCriterionInput.Values(List.of("Ready"), false));
    AutofilterSortStateInput sortState =
        AutofilterSortStateInput.withoutSortMethod(
            new RangeRef("A1:B2"), false, false, List.of(sortCondition));
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
            new SecretReference("store-pass"),
            Optional.empty(),
            ExcelOoxmlSignatureDigestAlgorithm.SHA256,
            Optional.empty());
    OoxmlSignatureInput explicitSignature =
        readOmitted(
            OoxmlSignatureInput.class,
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
    assertThrows(NullPointerException.class, () -> new CalculationPolicyInput(null, false));
    assertThrows(NullPointerException.class, () -> ExecutionPolicyInput.mode(null));
    assertThrows(NullPointerException.class, () -> new ExecutionJournalInput(null));
    assertThrows(NullPointerException.class, () -> new FormulaEnvironmentInput(null, null, null));
    assertThrows(
        tools.jackson.core.JacksonException.class,
        () -> readOmitted(SheetDisplayInput.class, null, null, null, null, null));
    assertThrows(
        tools.jackson.core.JacksonException.class,
        () -> readOmitted(SheetOutlineSummaryInput.class, null, null));
    assertThrows(
        tools.jackson.core.JacksonException.class,
        () -> readOmitted(SheetDefaultsInput.class, null, null));
    assertTrue(
        readOmitted(DataValidationPromptInput.class, text("Status"), text("Choose one."), null)
            .showPromptBox());
    assertTrue(
        readOmitted(
                DataValidationErrorAlertInput.class,
                ExcelDataValidationErrorStyle.STOP,
                text("Invalid status"),
                text("Use one of the allowed values."),
                null)
            .showErrorBox());
  }

  @Test
  void recordMetadataAppliesCatalogOwnedOmissionDefaults() {
    JsonMapper mapper = dev.erst.gridgrind.json.codec.GridGrindJsonMapperSupport.JSON_MAPPER;

    ExecutionPolicyInput execution = readJson(mapper, "{}", ExecutionPolicyInput.class);
    ExecutionJournalInput journal = readJson(mapper, "{}", ExecutionJournalInput.class);
    CalculationPolicyInput calculation = readJson(mapper, "{}", CalculationPolicyInput.class);
    FormulaEnvironmentInput formulaEnvironment =
        readJson(mapper, "{}", FormulaEnvironmentInput.class);
    OoxmlEncryptionInput encryption =
        readJson(
            mapper,
            """
            {
              "passwordRef": { "id": "persist-pass" }
            }
            """,
            OoxmlEncryptionInput.class);
    OoxmlSignatureInput signature =
        readJson(
            mapper,
            """
            {
              "pkcs12Path": "keys/signing.p12",
              "keystorePasswordRef": { "id": "store-pass" }
            }
            """,
            OoxmlSignatureInput.class);
    OoxmlEncryptionInput explicitEncryption =
        readOmitted(
            OoxmlEncryptionInput.class,
            new SecretReference("persist-pass"),
            ExcelOoxmlWriteCipher.AES_192,
            ExcelOoxmlWriteHash.SHA_384);
    OoxmlSignatureInput createdSignatureWithNullOptionals =
        readOmitted(
            OoxmlSignatureInput.class,
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
    AutofilterFilterCriterionInput.Values criterion =
        new AutofilterFilterCriterionInput.Values(List.of("Ready"), false);
    AutofilterSortConditionInput valueCondition =
        new AutofilterSortConditionInput.Value(new RangeRef("A1:A4"), true);
    AutofilterFilterColumnInput visibleButtonColumn =
        new AutofilterFilterColumnInput(2L, true, criterion);
    AutofilterFilterColumnInput hiddenButtonColumn =
        new AutofilterFilterColumnInput(2L, false, criterion);
    AutofilterSortStateInput defaultedSortState =
        AutofilterSortStateInput.withoutSortMethod(
            new RangeRef("A1:B4"), false, false, List.of(valueCondition));
    AutofilterSortStateInput explicitSortState =
        new AutofilterSortStateInput(
            new RangeRef("A1:B4"),
            true,
            true,
            Optional.of(ExcelAutofilterSortMethod.PINYIN),
            List.of(valueCondition));
    ExecutionModeInput customMode = ExecutionModeInput.eventRead();
    ExecutionModeInput writeOnlyCustomMode = ExecutionModeInput.streamingWrite();
    ExecutionPolicyInput customPolicy = ExecutionPolicyInput.mode(customMode);
    ExecutionPolicyInput customJournalPolicy =
        ExecutionPolicyInput.journal(new ExecutionJournalInput(ExecutionJournalLevel.VERBOSE));
    SheetDisplayInput explicitDisplay = new SheetDisplayInput(false, true, false, true, false);
    SheetOutlineSummaryInput explicitOutline = new SheetOutlineSummaryInput(false, true);
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
    assertTrue(readOmitted(AutofilterFilterColumnInput.class, 2L, null, criterion).showButton());
    AutofilterSortStateInput defaultedWireSortState =
        readOmitted(
            AutofilterSortStateInput.class,
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
    CommentInput comment =
        readOmitted(CommentInput.class, text("Owner note"), "Ada", null, null, null);
    AutofilterSortConditionInput valueCondition =
        new AutofilterSortConditionInput.Value(new RangeRef("A1:A4"), true);
    AutofilterSortStateInput defaultedSortState =
        readOmitted(
            AutofilterSortStateInput.class,
            new RangeRef("A1:B4"),
            false,
            true,
            null,
            List.of(valueCondition));
    WorkbookProtectionInput protection =
        readOmitted(WorkbookProtectionInput.class, null, null, null, null, null);
    PivotTableInput.DataField dataField =
        readOmitted(
            PivotTableInput.DataField.class,
            "Amount",
            ExcelPivotDataConsolidateFunction.SUM,
            null,
            null);
    TableColumnInput tableColumn = readOmitted(TableColumnInput.class, 0, null, null, null, null);
    TableColumnInput explicitFactoryColumn =
        readOmitted(TableColumnInput.class, 1, "owner", "Total", " SUM ", "=[@Amount]");
    DataValidationInput validation =
        readOmitted(
            DataValidationInput.class,
            new DataValidationRuleInput.ExplicitList(List.of("Ready")),
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
    TableInput defaultedTable =
        TableInput.withDefaultMetadata(
            new TableName("Budget"),
            new SheetName("Budget"),
            new RangeRef("A1:B3"),
            false,
            new TableStyleInput.None());
    TableInput explicitTable =
        new TableInput(
            new TableName("StyledBudget"),
            new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
            new RangeRef("A1:C3"),
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
            new TableName("VisibleAutofilterBudget"),
            new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
            new RangeRef("A1:C3"),
            false,
            true,
            new TableStyleInput.None(),
            text(""),
            false,
            false,
            false,
            "",
            "",
            "",
            List.of());
    TableInput explicitTotalsAndNoAutofilterTable =
        new TableInput(
            new TableName("TotalsWithoutAutofilterBudget"),
            new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
            new RangeRef("A1:C3"),
            true,
            false,
            new TableStyleInput.None(),
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
    TableInput defaultedWireTable =
        readOmitted(
            TableInput.class,
            new TableName("Budget"),
            new SheetName("Budget"),
            new RangeRef("A1:B3"),
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
    DataValidationErrorAlertInput explicitAlert =
        new DataValidationErrorAlertInput(
            ExcelDataValidationErrorStyle.STOP,
            text("Invalid"),
            text("Choose an allowed value."),
            true);
    DataValidationErrorAlertInput hiddenAlert =
        new DataValidationErrorAlertInput(
            ExcelDataValidationErrorStyle.WARNING,
            text("Caution"),
            text("Review the warning."),
            false);
    ChartSeriesInput series =
        new ChartSeriesInput(
            new ChartTitleInput.None(),
            new ChartDataSourceInput.StringLiteral(List.of("Jan")),
            new ChartDataSourceInput.NumericLiteral(List.of(10.0d)),
            Optional.of(true),
            Optional.of(ExcelChartMarkerStyle.CIRCLE),
            Optional.of((short) 6),
            Optional.empty());
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
            false,
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
            false);
    PrintSetupInput defaultedPrintSetup = PrintSetupInput.defaults();
    PrintSetupInput explicitPrintSetup =
        new PrintSetupInput(
            new PrintMarginsInput(0.6d, 0.6d, 0.7d, 0.7d, 0.2d, 0.2d),
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
        tools.jackson.core.JacksonException.class,
        () ->
            readOmitted(
                ChartInput.class,
                "BudgetChart",
                twoCellAnchor(),
                null,
                null,
                null,
                null,
                List.of(
                    new ChartPlotInput.Line(false, ExcelChartGrouping.STANDARD, List.of(series)))));
    assertThrows(
        tools.jackson.core.JacksonException.class,
        () ->
            readOmitted(
                ChartAxisInput.class,
                ExcelChartAxisKind.CATEGORY,
                ExcelChartAxisPosition.BOTTOM,
                ExcelChartAxisCrosses.AUTO_ZERO,
                null));
    assertThrows(
        tools.jackson.core.JacksonException.class,
        () ->
            readOmitted(
                PrintSetupInput.class,
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
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());
    WorkbookPlan explicitExecutionPlan =
        WorkbookPlan.identified(
            GridGrindProtocolVersion.current(),
            new PlanId("budget-execution"),
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.journal(new ExecutionJournalInput(ExecutionJournalLevel.VERBOSE)),
            FormulaEnvironmentInput.empty(),
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
    BorderSideInput borderSide =
        new BorderSideInput(Optional.of(ExcelBorderStyle.THIN), Optional.empty());
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
            Optional.of(
                new DifferentialBorderSideReport(
                    Optional.of(ExcelBorderStyle.THIN), Optional.empty())),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    AutofilterSortConditionInput explicitSortCondition =
        new AutofilterSortConditionInput.Icon(new RangeRef("D2:D9"), false, 1);
    AutofilterFilterColumnInput explicitFilterColumn =
        new AutofilterFilterColumnInput(
            2L, false, new AutofilterFilterCriterionInput.Values(List.of("Queued"), true));
    AutofilterSortStateInput explicitSortState =
        new AutofilterSortStateInput(
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
    SheetDisplayInput explicitDisplay = new SheetDisplayInput(false, true, false, true, true);
    SheetOutlineSummaryInput explicitOutline = new SheetOutlineSummaryInput(false, true);
    SheetDefaultsInput explicitSheetDefaults = new SheetDefaultsInput(18, 24.5d);
    SheetPresentationInput explicitPresentation =
        new SheetPresentationInput(
            explicitDisplay,
            Optional.of(ColorInput.rgb(new RgbHex("#aabbcc"))),
            explicitOutline,
            explicitSheetDefaults,
            List.of(
                new IgnoredErrorInput(
                    new RangeRef("A1:A3"), List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT))));
    FormulaEnvironmentInput explicitEnvironment =
        new FormulaEnvironmentInput(
            List.of(new FormulaExternalWorkbookInput("Rates.xlsx", "tmp/rates.xlsx")),
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
            Optional.of(true),
            Optional.of(ExcelChartMarkerStyle.DIAMOND),
            Optional.of((short) 8),
            Optional.of(12L));
    ChartSeriesInput convenienceSeries =
        ChartSeriesInput.untitled(
            new ChartDataSourceInput.Reference("Categories"),
            new ChartDataSourceInput.Reference("Values"),
            Optional.of(false),
            Optional.of(ExcelChartMarkerStyle.CIRCLE),
            Optional.of((short) 6),
            Optional.empty());
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
            false,
            ExcelChartGrouping.STACKED,
            Optional.of(24),
            explicitAxes,
            List.of(explicitSeries));
    ChartPlotInput.Area3D convenienceArea3D =
        new ChartPlotInput.Area3D(
            false, ExcelChartGrouping.STANDARD, Optional.of(16), List.of(explicitSeries));
    ChartPlotInput.Bar3D explicitBar3D =
        new ChartPlotInput.Bar3D(
            false,
            ExcelChartBarDirection.BAR,
            ExcelChartBarGrouping.PERCENT_STACKED,
            Optional.of(12),
            Optional.of(64),
            Optional.of(ExcelChartBarShape.CONE),
            explicitAxes,
            List.of(explicitSeries));
    ChartPlotInput.Bar3D convenienceBar3D =
        new ChartPlotInput.Bar3D(
            false,
            ExcelChartBarDirection.COLUMN,
            ExcelChartBarGrouping.CLUSTERED,
            Optional.of(18),
            Optional.of(90),
            Optional.of(ExcelChartBarShape.BOX),
            List.of(explicitSeries));
    ChartPlotInput.Line explicitLine =
        new ChartPlotInput.Line(
            true, ExcelChartGrouping.STACKED, explicitAxes, List.of(explicitSeries));
    ChartPlotInput.Line convenienceLine =
        new ChartPlotInput.Line(false, ExcelChartGrouping.STANDARD, List.of(explicitSeries));
    ChartPlotInput.Line3D explicitLine3D =
        new ChartPlotInput.Line3D(
            true,
            ExcelChartGrouping.PERCENT_STACKED,
            Optional.of(32),
            explicitAxes,
            List.of(explicitSeries));
    ChartPlotInput.Line3D convenienceLine3D =
        new ChartPlotInput.Line3D(
            false, ExcelChartGrouping.STANDARD, Optional.of(8), List.of(explicitSeries));
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
  void chartPlotWireDefaultsHandleOmittedBooleansAndOptionals() {
    ChartSeriesInput series = sampleSeries();
    List<ChartAxisInput> axes = sampleAxes();
    ChartPlotInput.Area3D area3D =
        readOmitted(
            ChartPlotInput.Area3D.class,
            null,
            ExcelChartGrouping.STANDARD,
            null,
            axes,
            List.of(series));
    ChartPlotInput.Bar3D bar3D =
        readOmitted(
            ChartPlotInput.Bar3D.class,
            null,
            ExcelChartBarDirection.COLUMN,
            ExcelChartBarGrouping.CLUSTERED,
            null,
            null,
            null,
            axes,
            List.of(series));
    ChartPlotInput.Doughnut doughnut =
        readOmitted(ChartPlotInput.Doughnut.class, null, null, null, List.of(series));
    ChartPlotInput.Line line =
        readOmitted(
            ChartPlotInput.Line.class, null, ExcelChartGrouping.STANDARD, axes, List.of(series));
    ChartPlotInput.Line3D line3D =
        readOmitted(
            ChartPlotInput.Line3D.class,
            null,
            ExcelChartGrouping.STANDARD,
            null,
            axes,
            List.of(series));
    ChartPlotInput.Pie pie = readOmitted(ChartPlotInput.Pie.class, null, null, List.of(series));
    ChartPlotInput.Pie3D pie3D = readOmitted(ChartPlotInput.Pie3D.class, null, List.of(series));
    ChartPlotInput.Radar radar =
        readOmitted(
            ChartPlotInput.Radar.class, null, ExcelChartRadarStyle.STANDARD, axes, List.of(series));
    ChartPlotInput.Scatter scatter =
        readOmitted(
            ChartPlotInput.Scatter.class, null, ExcelChartScatterStyle.LINE, axes, List.of(series));
    ChartPlotInput.Surface surface =
        readOmitted(ChartPlotInput.Surface.class, null, null, axes, List.of(series));
    ChartPlotInput.Surface3D surface3D =
        readOmitted(ChartPlotInput.Surface3D.class, null, null, axes, List.of(series));

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

    assertEquals(new PrintMarginsInput(0.6d, 0.6d, 0.7d, 0.7d, 0.3d, 0.3d), printSetup.margins());
    assertEquals(ExcelPrintOrientation.PORTRAIT, printLayout.orientation());
    assertTrue(printLayout.setup().printGridlines());
    assertEquals(PrintSetupInput.defaults(), conveniencePrintLayout.setup());
    assertEquals(PrintSetupInput.defaults(), createdPrintLayout.setup());
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

  private static DrawingAnchorInput.TwoCell twoCellAnchor() {
    return new DrawingAnchorInput.TwoCell(
        new DrawingMarkerInput(1, 0, 1, 0),
        new DrawingMarkerInput(5, 0, 10, 0),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  private static TextSourceInput text(String value) {
    return TextSourceInput.inline(value);
  }

  private static ChartSeriesInput sampleSeries() {
    return ChartSeriesInput.untitled(
        new ChartDataSourceInput.StringLiteral(List.of("Q1")),
        new ChartDataSourceInput.NumericLiteral(List.of(1.0d)),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
  }

  private static List<ChartAxisInput> sampleAxes() {
    return List.of(
        new ChartAxisInput(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisInput(
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
