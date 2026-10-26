package dev.erst.gridgrind.engine.runtime.parity;

import static dev.erst.gridgrind.engine.runtime.parity.ParityPlanSupport.mutate;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.drawing.CommentAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.CommentSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.formula.FormulaExternalWorkbookSpec;
import dev.erst.gridgrind.core.model.formula.FormulaMissingWorkbookPolicy;
import dev.erst.gridgrind.core.model.formula.FormulaUdfFunctionSpec;
import dev.erst.gridgrind.core.model.formula.FormulaUdfToolpackSpec;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextReport;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextSpec;
import dev.erst.gridgrind.core.model.print.PrintAreaReport;
import dev.erst.gridgrind.core.model.print.PrintAreaSpec;
import dev.erst.gridgrind.core.model.print.PrintLayoutReport;
import dev.erst.gridgrind.core.model.print.PrintLayoutSpec;
import dev.erst.gridgrind.core.model.print.PrintMarginsSpec;
import dev.erst.gridgrind.core.model.print.PrintScalingReport;
import dev.erst.gridgrind.core.model.print.PrintScalingSpec;
import dev.erst.gridgrind.core.model.print.PrintSetupReport;
import dev.erst.gridgrind.core.model.print.PrintSetupSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsReport;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsReport;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsSpec;
import dev.erst.gridgrind.core.model.protection.SheetProtectionReport;
import dev.erst.gridgrind.core.model.protection.SheetProtectionSettings;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionReport;
import dev.erst.gridgrind.core.model.security.OoxmlEncryptionReport;
import dev.erst.gridgrind.core.model.security.OoxmlOpenSecuritySpec;
import dev.erst.gridgrind.core.model.security.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.core.model.security.OoxmlSignatureSpec;
import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.CellBorderPatch;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellFillSpec;
import dev.erst.gridgrind.core.model.style.CellFontPatch;
import dev.erst.gridgrind.core.model.style.CellGradientFillSpec;
import dev.erst.gridgrind.core.model.style.CellGradientStopSpec;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionReport;
import dev.erst.gridgrind.core.model.table.TableColumnReport;
import dev.erst.gridgrind.core.model.table.TableColumnSpec;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingThresholdSpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.CellReadFacet;
import dev.erst.gridgrind.core.plan.CellReadProjection;
import dev.erst.gridgrind.core.plan.SheetInspectionResult;
import dev.erst.gridgrind.core.plan.SheetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingIconSet;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlChainingMode;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlCipherAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlHashAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureState;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.engine.runtime.parity.ParityPlanSupport.PendingMutation;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.ConcurrentHashMap;
import java.util.function.Function;
import java.util.stream.Collectors;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Registry of executable parity probes keyed by stable ledger probe identifiers. */
final class XlsxParityProbeRegistry {
  private XlsxParityProbeRegistry() {}

  static ProbeResult runProbe(String probeId, ProbeContext context) {
    return switch (probeId) {
      case "probe-core-read" -> XlsxParityCoreProbeGroup.probeCoreRead(context);
      case "probe-named-range-read" -> XlsxParityCoreProbeGroup.probeNamedRangeRead(context);
      case "probe-workbook-protection-read-gap" ->
          XlsxParityCoreProbeGroup.probeWorkbookProtectionReadGap(context);
      case "probe-advanced-print-read-gap" ->
          XlsxParityCoreProbeGroup.probeAdvancedPrintReadGap(context);
      case "probe-advanced-style-read-gap" ->
          XlsxParityCoreProbeGroup.probeAdvancedStyleReadGap(context);
      case "probe-rich-comment-read-gap" ->
          XlsxParityCoreProbeGroup.probeRichCommentReadGap(context);
      case "probe-data-validation-analysis-gap" ->
          XlsxParityCoreProbeGroup.probeDataValidationAnalysisGap(context);
      case "probe-autofilter-criteria-gap" ->
          XlsxParityCoreProbeGroup.probeAutofilterCriteriaReadGap(context);
      case "probe-table-advanced-gap" ->
          XlsxParityCoreProbeGroup.probeTableAdvancedReadGap(context);
      case "probe-conditional-formatting-modeled-read" ->
          XlsxParityCoreProbeGroup.probeConditionalFormattingModeledRead(context);
      case "probe-conditional-formatting-unmodeled-read-gap" ->
          XlsxParityCoreProbeGroup.probeConditionalFormattingUnmodeledReadGap(context);
      case "probe-core-mutation" -> XlsxParityCoreProbeGroup.probeCoreMutation(context);
      case "probe-workbook-protection-mutation-gap" ->
          XlsxParityCoreProbeGroup.probeWorkbookProtectionMutationGap(context);
      case "probe-sheet-password-mutation-gap" ->
          XlsxParityCoreProbeGroup.probeSheetPasswordMutationGap(context);
      case "probe-sheet-copy-gap" -> XlsxParityCoreProbeGroup.probeSheetCopyGap(context);
      case "probe-advanced-print-mutation-gap" ->
          XlsxParityCoreProbeGroup.probeAdvancedPrintMutationGap(context);
      case "probe-advanced-style-mutation-gap" ->
          XlsxParityCoreProbeGroup.probeAdvancedStyleMutationGap(context);
      case "probe-rich-comment-mutation-gap" ->
          XlsxParityCoreProbeGroup.probeRichCommentMutationGap(context);
      case "probe-named-range-formula-mutation-gap" ->
          XlsxParityCoreProbeGroup.probeNamedRangeFormulaMutationGap(context);
      case "probe-autofilter-criteria-mutation-gap" ->
          XlsxParityCoreProbeGroup.probeAutofilterCriteriaMutationGap(context);
      case "probe-table-advanced-mutation-gap" ->
          XlsxParityCoreProbeGroup.probeTableAdvancedMutationGap(context);
      case "probe-conditional-formatting-advanced-mutation-gap" ->
          XlsxParityCoreProbeGroup.probeConditionalFormattingAdvancedMutationGap(context);
      case "probe-formula-core" -> XlsxParityFormulaProbeGroup.probeFormulaCore(context);
      case "probe-formula-external" -> XlsxParityFormulaProbeGroup.probeFormulaExternal(context);
      case "probe-formula-missing-workbook-policy" ->
          XlsxParityFormulaProbeGroup.probeFormulaMissingWorkbookPolicy(context);
      case "probe-formula-udf" -> XlsxParityFormulaProbeGroup.probeFormulaUdf(context);
      case "probe-formula-lifecycle" -> XlsxParityFormulaProbeGroup.probeFormulaLifecycle(context);
      case "probe-drawing-readback" ->
          XlsxParityDrawingChartPivotProbeGroup.probeDrawingReadback(context);
      case "probe-drawing-authoring" ->
          XlsxParityDrawingChartPivotProbeGroup.probeDrawingAuthoring(context);
      case "probe-drawing-comment-coexistence" ->
          XlsxParityDrawingChartPivotProbeGroup.probeDrawingCommentCoexistence(context);
      case "probe-drawing-merged-image-preservation" ->
          XlsxParityDrawingChartPivotProbeGroup.probeDrawingMergedImagePreservation(context);
      case "probe-embedded-object-platform" ->
          XlsxParityDrawingChartPivotProbeGroup.probeEmbeddedObjectPlatform(context);
      case "probe-chart-readback" ->
          XlsxParityDrawingChartPivotProbeGroup.probeChartReadback(context);
      case "probe-chart-authoring" ->
          XlsxParityDrawingChartPivotProbeGroup.probeChartAuthoring(context);
      case "probe-chart-mutation" ->
          XlsxParityDrawingChartPivotProbeGroup.probeChartMutation(context);
      case "probe-chart-preservation" ->
          XlsxParityDrawingChartPivotProbeGroup.probeChartPreservation(context);
      case "probe-chart-unsupported" ->
          XlsxParityDrawingChartPivotProbeGroup.probeChartUnsupported(context);
      case "probe-pivot-readback" ->
          XlsxParityDrawingChartPivotProbeGroup.probePivotReadback(context);
      case "probe-pivot-authoring" ->
          XlsxParityDrawingChartPivotProbeGroup.probePivotAuthoring(context);
      case "probe-pivot-mutation" ->
          XlsxParityDrawingChartPivotProbeGroup.probePivotMutation(context);
      case "probe-pivot-preservation" ->
          XlsxParityDrawingChartPivotProbeGroup.probePivotPreservation(context);
      case "probe-event-model-gap" -> XlsxParitySecurityProbeGroup.probeEventModelGap(context);
      case "probe-sxssf-gap" -> XlsxParitySecurityProbeGroup.probeSxssfGap(context);
      case "probe-encryption-gap" -> XlsxParitySecurityProbeGroup.probeEncryptionGap(context);
      case "probe-encryption-password-errors" ->
          XlsxParitySecurityProbeGroup.probeEncryptionPasswordErrors(context);
      case "probe-signing-gap" -> XlsxParitySecurityProbeGroup.probeSigningGap(context);
      case "probe-signing-invalid-signature" ->
          XlsxParitySecurityProbeGroup.probeSigningInvalidSignature(context);
      default -> throw new IllegalArgumentException("Unknown parity probe id: " + probeId);
    };
  }

  static FormulaEnvironmentSpec externalFormulaEnvironment(Path referencedWorkbookPath) {
    return new FormulaEnvironmentSpec(
        List.of(
            new FormulaExternalWorkbookSpec(
                "referenced.xlsx", referencedWorkbookPath.toAbsolutePath().toString())),
        FormulaMissingWorkbookPolicy.ERROR,
        List.of());
  }

  static FormulaEnvironmentSpec missingWorkbookCachedValueEnvironment() {
    return new FormulaEnvironmentSpec(
        List.of(), FormulaMissingWorkbookPolicy.USE_CACHED_VALUE, List.of());
  }

  static FormulaEnvironmentSpec udfFormulaEnvironment() {
    return new FormulaEnvironmentSpec(
        List.of(),
        FormulaMissingWorkbookPolicy.ERROR,
        List.of(
            new FormulaUdfToolpackSpec(
                "math", List.of(new FormulaUdfFunctionSpec("DOUBLE", 1, "ARG1*2")))));
  }

  static OoxmlOpenSecuritySpec encryptedOpenSecurity() {
    return new OoxmlOpenSecuritySpec(
        java.util.Optional.of(
            new dev.erst.gridgrind.core.source.SecretReference("source-password")));
  }

  static OoxmlSignatureSpec signingInput(Path pkcs12Path, String description) {
    return new OoxmlSignatureSpec(
        pkcs12Path.toAbsolutePath().toString(),
        new dev.erst.gridgrind.core.source.SecretReference("keystore-password"),
        java.util.Optional.of(new dev.erst.gridgrind.core.source.SecretReference("key-password")),
        java.util.Optional.of(XlsxParityScenarios.SIGNING_KEY_ALIAS),
        dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureDigestAlgorithm.SHA256,
        java.util.Optional.ofNullable(description));
  }

  static <T> T cast(Class<T> type, Object value) {
    return type.cast(value);
  }

  static boolean hasEncryptedAgilePackage(OoxmlPackageSecurityReport security) {
    return security.encryption() instanceof OoxmlEncryptionReport.Encrypted encrypted
        && encrypted.mode() == ExcelOoxmlEncryptionMode.AGILE
        && encrypted.cipherAlgorithm() == ExcelOoxmlCipherAlgorithm.AES_256
        && encrypted.hashAlgorithm() == ExcelOoxmlHashAlgorithm.SHA_512
        && encrypted.chainingMode() == ExcelOoxmlChainingMode.CBC
        && security.signatures().isEmpty();
  }

  static boolean hasSingleSignatureState(
      OoxmlPackageSecurityReport security, ExcelOoxmlSignatureState state) {
    return security.encryption() instanceof OoxmlEncryptionReport.None
        && security.signatures().size() == 1
        && security.signatures().getFirst().state() == state;
  }

  static String textCell(SheetInspectionResult.CellsResult cells, String address) {
    return cast(
            dev.erst.gridgrind.core.model.value.CellReport.TextReport.class,
            byAddress(cells.cells()).get(address))
        .textValue()
        .orElseThrow();
  }

  static SheetIntrospectionQuery.GetCells allFacetCellsQuery() {
    return new SheetIntrospectionQuery.GetCells(
        Optional.of(
            CellReadProjection.of(
                CellReadFacet.VALUE,
                CellReadFacet.STYLE,
                CellReadFacet.FORMAT,
                CellReadFacet.HYPERLINK,
                CellReadFacet.COMMENT,
                CellReadFacet.FORMULA,
                CellReadFacet.RICH_TEXT_RUNS,
                CellReadFacet.TEMPORAL)));
  }

  static Map<String, dev.erst.gridgrind.core.model.value.CellReport> byAddress(
      List<dev.erst.gridgrind.core.model.value.CellReport> cells) {
    return cells.stream()
        .collect(Collectors.toUnmodifiableMap(cell -> cell.address().value(), Function.identity()));
  }

  static boolean matchesWorkbookProtectionReport(
      WorkbookProtectionReport observed, XlsxParityOracle.WorkbookProtectionSnapshot direct) {
    return observed.structureLocked() == direct.structureLocked()
        && observed.windowsLocked() == direct.windowsLocked()
        && observed.revisionsLocked() == direct.revisionsLocked()
        && observed.workbookPasswordHashPresent() == direct.workbookPasswordHashPresent()
        && observed.revisionsPasswordHashPresent() == direct.revisionsPasswordHashPresent();
  }

  static SheetProtectionSettings sheetProtectionSettings(
      SheetInspectionResult.SheetSummaryResult summary) {
    return switch (summary.sheet().protection()) {
      case SheetProtectionReport.Unprotected _ -> null;
      case SheetProtectionReport.Protected protectedReport -> protectedReport.settings();
    };
  }

  static boolean sheetPasswordMatches(Path workbookPath, String sheetName, String password) {
    return XlsxParitySupport.call(
        "validate protected-sheet password " + sheetName + "@" + workbookPath.getFileName(),
        () -> {
          try (XSSFWorkbook workbook =
              (XSSFWorkbook) WorkbookFactory.create(workbookPath.toFile())) {
            return workbook.getSheet(sheetName).validateSheetPassword(password);
          }
        });
  }

  static PrintLayoutReport defaultPrintLayoutReport(String sheetName) {
    return new PrintLayoutReport(
        new SheetLabel(sheetName),
        new PrintAreaReport.None(),
        ExcelPrintOrientation.PORTRAIT,
        new PrintScalingReport.Automatic(),
        new PrintTitleRowsReport.None(),
        new PrintTitleColumnsReport.None(),
        new HeaderFooterTextReport("", "", ""),
        new HeaderFooterTextReport("", "", ""),
        PrintSetupReport.defaults());
  }

  static PrintLayoutSpec advancedPrintLayoutInput() {
    return new PrintLayoutSpec(
        new PrintAreaSpec.None(),
        ExcelPrintOrientation.PORTRAIT,
        new PrintScalingSpec.Automatic(),
        new PrintTitleRowsSpec.None(),
        new PrintTitleColumnsSpec.None(),
        HeaderFooterTextSpec.blank(),
        HeaderFooterTextSpec.blank(),
        new PrintSetupSpec(
            new PrintMarginsSpec(0.35d, 0.55d, 0.60d, 0.45d, 0.3d, 0.3d),
            true,
            true,
            true,
            8,
            true,
            true,
            2,
            true,
            4,
            List.of(6),
            List.of(3)));
  }

  static CellStylePatch advancedThemedStyleInput() {
    return new CellStylePatch(
        java.util.Optional.empty(),
        java.util.Optional.empty(),
        java.util.Optional.of(
            new CellFontPatch(
                java.util.Optional.empty(),
                java.util.Optional.of(true),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.of(ColorSpec.theme(6, -0.35d)),
                java.util.Optional.empty(),
                java.util.Optional.empty())),
        java.util.Optional.of(
            CellFillSpec.patternForeground(ExcelFillPattern.SOLID, ColorSpec.theme(3, 0.30d))),
        java.util.Optional.of(
            new CellBorderPatch(
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.of(
                    new BorderSidePatch(
                        ExcelBorderStyle.THIN,
                        ColorSpec.indexed(Short.toUnsignedInt(IndexedColors.DARK_RED.getIndex())))),
                java.util.Optional.empty())),
        java.util.Optional.empty());
  }

  static CellStylePatch advancedGradientStyleInput() {
    return new CellStylePatch(
        java.util.Optional.empty(),
        java.util.Optional.empty(),
        java.util.Optional.empty(),
        java.util.Optional.of(
            CellFillSpec.gradient(
                CellGradientFillSpec.linear(
                    java.util.Optional.of(45.0d),
                    List.of(
                        new CellGradientStopSpec(0.0d, ColorSpec.rgb(new RgbHex("#1F497D"))),
                        new CellGradientStopSpec(1.0d, ColorSpec.theme(4, 0.45d)))))),
        java.util.Optional.empty(),
        java.util.Optional.empty());
  }

  static CommentSpec advancedCommentInput() {
    return comment(
        "Lead review scheduled",
        "GridGrind",
        false,
        List.of(
            richTextRun(
                "Lead",
                new CellFontPatch(
                    java.util.Optional.of(true),
                    java.util.Optional.empty(),
                    java.util.Optional.empty(),
                    java.util.Optional.empty(),
                    java.util.Optional.of(ColorSpec.theme(4, -0.20d)),
                    java.util.Optional.empty(),
                    java.util.Optional.empty())),
            richTextRun(" ", null),
            richTextRun(
                "review scheduled",
                new CellFontPatch(
                    java.util.Optional.empty(),
                    java.util.Optional.of(true),
                    java.util.Optional.empty(),
                    java.util.Optional.empty(),
                    java.util.Optional.of(
                        ColorSpec.indexed(
                            Short.toUnsignedInt(IndexedColors.DARK_GREEN.getIndex()))),
                    java.util.Optional.empty(),
                    java.util.Optional.empty()))),
        new CommentAnchorSpec(4, 1, 8, 7));
  }

  static DrawingAnchorSpec.TwoCell twoCellAnchorInput(
      int fromColumn, int fromRow, int toColumn, int toRow, ExcelDrawingAnchorBehavior behavior) {
    return new DrawingAnchorSpec.TwoCell(
        new DrawingMarkerSpec(fromColumn, fromRow, 0, 0),
        new DrawingMarkerSpec(toColumn, toRow, 0, 0),
        behavior);
  }

  static TableSpec advancedTableInput() {
    return new TableSpec(
        new TableName("AdvancedTable"),
        new dev.erst.gridgrind.core.vocabulary.SheetName("Advanced"),
        new RangeRef("H1:J5"),
        true,
        false,
        new TableStyleSpec.None(),
        inlineText("Parity advanced table"),
        true,
        true,
        true,
        "ParityHeader",
        "ParityData",
        "ParityTotals",
        List.of(
            new TableColumnSpec(0, "", "Total", "", ""),
            new TableColumnSpec(1, "", "", "SUM", "[@Sales]*2"),
            new TableColumnSpec(2, "owner-unique", "", "", "")));
  }

  static List<PendingMutation> advancedConditionalFormattingMutations() {
    PendingMutation colorScale =
        mutate(
            new RangeSelector.ByRange(new SheetName("Advanced"), new RangeRef("L2:L5")),
            new StructuredMutationAction.SetConditionalFormatting(
                new ConditionalFormattingSpec(
                    List.of(
                        new ConditionalFormattingRuleSpec.ColorScaleRule(
                            false,
                            List.of(
                                new ConditionalFormattingThresholdSpec.Min(),
                                new ConditionalFormattingThresholdSpec.Percentile(50.0d),
                                new ConditionalFormattingThresholdSpec.Max()),
                            List.of(
                                ColorSpec.rgb(new RgbHex("#AA2211")),
                                ColorSpec.rgb(new RgbHex("#FFDD55")),
                                ColorSpec.rgb(new RgbHex("#11CC66"))))))));
    PendingMutation dataBar =
        mutate(
            new RangeSelector.ByRange(new SheetName("Advanced"), new RangeRef("M2:M5")),
            new StructuredMutationAction.SetConditionalFormatting(
                new ConditionalFormattingSpec(
                    List.of(
                        new ConditionalFormattingRuleSpec.DataBarRule(
                            false,
                            ColorSpec.rgb(new RgbHex("#123456")),
                            true,
                            10,
                            90,
                            new ConditionalFormattingThresholdSpec.Min(),
                            new ConditionalFormattingThresholdSpec.Max())))));
    PendingMutation iconSet =
        mutate(
            new RangeSelector.ByRange(new SheetName("Advanced"), new RangeRef("N2:N5")),
            new StructuredMutationAction.SetConditionalFormatting(
                new ConditionalFormattingSpec(
                    List.of(
                        new ConditionalFormattingRuleSpec.IconSetRule(
                            false,
                            ExcelConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS,
                            true,
                            true,
                            List.of(
                                new ConditionalFormattingThresholdSpec.Percent(0.0d),
                                new ConditionalFormattingThresholdSpec.Percent(33.0d),
                                new ConditionalFormattingThresholdSpec.Percent(67.0d)))))));
    PendingMutation top10 =
        mutate(
            new RangeSelector.ByRange(new SheetName("Advanced"), new RangeRef("K2:K5")),
            new StructuredMutationAction.SetConditionalFormatting(
                new ConditionalFormattingSpec(
                    List.of(
                        new ConditionalFormattingRuleSpec.Top10Rule(
                            true, 10, false, false, java.util.Optional.empty())))));
    return List.of(colorScale, dataBar, iconSet, top10);
  }

  static ProbeResult pass(String detail) {
    return new ProbeResult(XlsxParityLedger.ExpectedOutcome.PASS, detail);
  }

  static ProbeResult fail(String detail) {
    return new ProbeResult(XlsxParityLedger.ExpectedOutcome.FAIL, detail);
  }

  static boolean sortConditionsMatch(
      AutofilterSortConditionReport observed, XlsxParityOracle.SortConditionSnapshot direct) {
    return observed.range().value().equals(direct.range())
        && observed.descending() == direct.descending()
        && switch (observed) {
          case AutofilterSortConditionReport.Value _ -> direct.sortBy().isEmpty();
          case AutofilterSortConditionReport.CellColor _ -> "cellColor".equals(direct.sortBy());
          case AutofilterSortConditionReport.FontColor _ -> "fontColor".equals(direct.sortBy());
          case AutofilterSortConditionReport.Icon _ -> "icon".equals(direct.sortBy());
        };
  }

  static TableEntryReport findTable(List<TableEntryReport> tables, String name) {
    return tables.stream()
        .filter(table -> name.equals(table.name().value()))
        .findFirst()
        .orElse(null);
  }

  static TableEntryReport findTableOnSheet(List<TableEntryReport> tables, String sheetName) {
    return tables.stream()
        .filter(table -> sheetName.equals(table.sheetName().value()))
        .findFirst()
        .orElse(null);
  }

  static boolean tableMatchesReadReport(
      TableEntryReport observed,
      TableEntryReport expected,
      String expectedName,
      String expectedSheetName) {
    return observed != null
        && (expectedName == null || observed.name().value().equals(expectedName))
        && (expectedSheetName == null || observed.sheetName().value().equals(expectedSheetName))
        && observed.range().equals(expected.range())
        && observed.structure().headerRowCount() == expected.structure().headerRowCount()
        && observed.structure().totalsRowCount() == expected.structure().totalsRowCount()
        && observed.structure().columnNames().equals(expected.structure().columnNames())
        && normalizedTableColumns(observed.structure().columns())
            .equals(normalizedTableColumns(expected.structure().columns()))
        && observed.style().equals(expected.style())
        && observed.behavior().hasAutofilter() == expected.behavior().hasAutofilter()
        && observed.presentation().comment().equals(expected.presentation().comment())
        && observed.behavior().published() == expected.behavior().published()
        && observed.behavior().insertRow() == expected.behavior().insertRow()
        && observed.behavior().insertRowShift() == expected.behavior().insertRowShift()
        && observed
            .presentation()
            .headerRowCellStyle()
            .equals(expected.presentation().headerRowCellStyle())
        && observed.presentation().dataCellStyle().equals(expected.presentation().dataCellStyle())
        && observed
            .presentation()
            .totalsRowCellStyle()
            .equals(expected.presentation().totalsRowCellStyle());
  }

  static List<String> normalizedTableColumns(List<TableColumnReport> columns) {
    return columns.stream()
        .map(
            column ->
                "%s|%s|%s|%s|%s"
                    .formatted(
                        column.name(),
                        column.uniqueName(),
                        column.totalsRowLabel(),
                        column.totalsRowFunction(),
                        column.calculatedColumnFormula()))
        .toList();
  }

  static boolean matchesAnchor(
      DrawingAnchorReport anchor, XlsxParityOracle.DrawingAnchorSnapshot direct) {
    if (!(anchor instanceof DrawingAnchorReport.TwoCell twoCell)) {
      return false;
    }
    return twoCell.from().columnIndex() == direct.col1()
        && twoCell.from().rowIndex() == direct.row1()
        && twoCell.to().columnIndex() == direct.col2()
        && twoCell.to().rowIndex() == direct.row2();
  }

  static boolean chartDrawingMatches(
      DrawingObjectReport.Chart observed, XlsxParityOracle.ChartDrawingObjectSnapshot direct) {
    return observed.name().equals(direct.name())
        && matchesAnchor(observed.anchor(), direct.anchor())
        && observed.supported() == direct.supported()
        && observed.plotTypeTokens().equals(direct.plotTypeTokens())
        && observed.title().equals(direct.title());
  }

  static boolean hasSinglePlot(ChartReport chart, Class<? extends ChartReport.Plot> plotType) {
    return chart.plots().size() == 1 && plotType.isInstance(chart.plots().getFirst());
  }

  static <T extends ChartReport.Plot> T onlyPlot(ChartReport chart, Class<T> plotType) {
    if (!hasSinglePlot(chart, plotType)) {
      throw new IllegalStateException(
          "Expected one " + plotType.getSimpleName() + " plot on chart " + chart.name());
    }
    return plotType.cast(chart.plots().getFirst());
  }

  static List<String> plotTypeTokens(ChartReport chart) {
    return chart.plots().stream()
        .map(
            plot ->
                switch (plot) {
                  case ChartReport.Area _ -> "AREA";
                  case ChartReport.Area3D _ -> "AREA_3D";
                  case ChartReport.Bar _ -> "BAR";
                  case ChartReport.Bar3D _ -> "BAR_3D";
                  case ChartReport.Doughnut _ -> "DOUGHNUT";
                  case ChartReport.Line _ -> "LINE";
                  case ChartReport.Line3D _ -> "LINE_3D";
                  case ChartReport.Pie _ -> "PIE";
                  case ChartReport.Pie3D _ -> "PIE_3D";
                  case ChartReport.Radar _ -> "RADAR";
                  case ChartReport.Scatter _ -> "SCATTER";
                  case ChartReport.Surface _ -> "SURFACE";
                  case ChartReport.Surface3D _ -> "SURFACE_3D";
                  case ChartReport.Unsupported unsupported -> unsupported.plotTypeToken();
                })
        .toList();
  }

  static <T extends XlsxParityOracle.DirectDrawingObjectSnapshot> T directObject(
      XlsxParityOracle.DrawingSheetSnapshot snapshot, String name, Class<T> type) {
    return snapshot.objects().stream()
        .filter(candidate -> candidate.name().equals(name))
        .findFirst()
        .filter(type::isInstance)
        .map(type::cast)
        .orElseThrow(
            () ->
                new IllegalStateException(
                    "Missing direct oracle drawing object "
                        + name
                        + " as "
                        + type.getSimpleName()));
  }

  static <T extends DrawingObjectReport> T drawingObjectReport(
      WorkbookAssetInspectionResult.DrawingObjectsResult result, String name, Class<T> type) {
    return result.drawingObjects().stream()
        .filter(candidate -> candidate.name().equals(name))
        .findFirst()
        .filter(type::isInstance)
        .map(type::cast)
        .orElseThrow(
            () ->
                new IllegalStateException(
                    "Missing GridGrind drawing report " + name + " as " + type.getSimpleName()));
  }

  static boolean matchesColorDescriptor(
      dev.erst.gridgrind.core.model.style.CellColorReport color, String descriptor) {
    if (descriptor == null || descriptor.isBlank()) {
      return color == null;
    }
    if (color == null) {
      return false;
    }
    List<String> parts = Arrays.asList(descriptor.split("\\|"));
    for (String part : parts) {
      if (part.startsWith("rgb=")) {
        String expected = "#" + part.substring("rgb=".length());
        if (!(color instanceof CellColorReport.Rgb rgb) || !expected.equals(rgb.rgb().value())) {
          return false;
        }
      } else if (part.startsWith("theme=")) {
        if (!(color instanceof CellColorReport.Theme theme)
            || !Integer.valueOf(part.substring("theme=".length())).equals(theme.theme())) {
          return false;
        }
      } else if (part.startsWith("indexed=")) {
        if (!(color instanceof CellColorReport.Indexed indexed)
            || !Integer.valueOf(part.substring("indexed=".length())).equals(indexed.indexed())) {
          return false;
        }
      } else if (part.startsWith("tint=")
          && (color.tint().isEmpty()
              || !approximatelyEquals(
                  Double.parseDouble(part.substring("tint=".length())),
                  color.tint().orElseThrow()))) {
        return false;
      }
    }
    return true;
  }

  static boolean approximatelyEquals(double left, double right) {
    return Math.abs(left - right) <= 0.000001d;
  }

  static CellValueSpec.Text text(String value) {
    return new CellValueSpec.Text(TextSourceSpec.inline(value));
  }

  static ChartTitleSpec.Text chartTitle(String value) {
    return new ChartTitleSpec.Text(TextSourceSpec.inline(value));
  }

  static RichTextRunSpec richTextRun(String value, CellFontPatch font) {
    return new RichTextRunSpec(TextSourceSpec.inline(value), java.util.Optional.ofNullable(font));
  }

  static CommentSpec comment(String text, String author, boolean visible) {
    return CommentSpec.plain(TextSourceSpec.inline(text), author, visible);
  }

  static CommentSpec comment(
      String text,
      String author,
      boolean visible,
      List<RichTextRunSpec> runs,
      CommentAnchorSpec anchor) {
    return new CommentSpec(
        TextSourceSpec.inline(text),
        author,
        visible,
        java.util.Optional.of(runs),
        java.util.Optional.ofNullable(anchor));
  }

  static TextSourceSpec inlineText(String value) {
    return TextSourceSpec.inline(value);
  }

  record ProbeResult(XlsxParityLedger.ExpectedOutcome outcome, String detail) {}

  /** Shared parity probe runtime context with scenario materialization and derived-file helpers. */
  static final class ProbeContext {
    private final Path temporaryRoot;
    private final Map<String, XlsxParityScenarios.MaterializedScenario> scenarioCache =
        new ConcurrentHashMap<>();

    ProbeContext(Path temporaryRoot) {
      this.temporaryRoot = temporaryRoot;
    }

    XlsxParityScenarios.MaterializedScenario scenario(String scenarioId) {
      XlsxParityScenarios.MaterializedScenario existing = scenarioCache.get(scenarioId);
      if (existing != null) {
        return existing;
      }
      XlsxParityScenarios.MaterializedScenario materialized =
          XlsxParityScenarios.materialize(scenarioId, temporaryRoot.resolve("corpus"));
      scenarioCache.put(scenarioId, materialized);
      return materialized;
    }

    XlsxParityScenarios.MaterializedScenario copiedScenario(String scenarioId, String copyName) {
      return XlsxParitySupport.call(
          "copy materialized parity scenario " + scenarioId,
          () -> {
            XlsxParityScenarios.MaterializedScenario source = scenario(scenarioId);
            Path directory =
                Files.createDirectories(temporaryRoot.resolve("scenario-copies").resolve(copyName));
            Path workbookCopy = directory.resolve(source.workbookPath().getFileName().toString());
            Files.copy(source.workbookPath(), workbookCopy, StandardCopyOption.REPLACE_EXISTING);
            Map<String, Path> attachmentCopies = new ConcurrentHashMap<>();
            for (Map.Entry<String, Path> entry : source.attachments().entrySet()) {
              Path attachmentCopy = directory.resolve(entry.getValue().getFileName().toString());
              Files.copy(entry.getValue(), attachmentCopy, StandardCopyOption.REPLACE_EXISTING);
              attachmentCopies.put(entry.getKey(), attachmentCopy);
            }
            return new XlsxParityScenarios.MaterializedScenario(workbookCopy, attachmentCopies);
          });
    }

    Path derivedWorkbook(String name) {
      Path directory = derivedDirectory("workbooks");
      return directory.resolve(name + ".xlsx");
    }

    Path derivedDirectory(String name) {
      return XlsxParitySupport.call(
          "create derived parity directory " + name,
          () -> Files.createDirectories(temporaryRoot.resolve(name)));
    }
  }
}
