package dev.erst.gridgrind.executor.parity;

import static dev.erst.gridgrind.executor.parity.ParityPlanSupport.mutate;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingIconSet;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureState;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import dev.erst.gridgrind.executor.parity.ParityPlanSupport.PendingMutation;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
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

  static FormulaEnvironmentInput externalFormulaEnvironment(Path referencedWorkbookPath) {
    return new FormulaEnvironmentInput(
        List.of(
            new FormulaExternalWorkbookInput(
                "referenced.xlsx", referencedWorkbookPath.toAbsolutePath().toString())),
        FormulaMissingWorkbookPolicy.ERROR,
        List.of());
  }

  static FormulaEnvironmentInput missingWorkbookCachedValueEnvironment() {
    return new FormulaEnvironmentInput(
        List.of(), FormulaMissingWorkbookPolicy.USE_CACHED_VALUE, List.of());
  }

  static FormulaEnvironmentInput udfFormulaEnvironment() {
    return new FormulaEnvironmentInput(
        List.of(),
        FormulaMissingWorkbookPolicy.ERROR,
        List.of(
            new FormulaUdfToolpackInput(
                "math", List.of(new FormulaUdfFunctionInput("DOUBLE", 1, null, "ARG1*2")))));
  }

  static OoxmlOpenSecurityInput encryptedOpenSecurity() {
    return new OoxmlOpenSecurityInput(XlsxParityScenarios.ENCRYPTION_PASSWORD);
  }

  static OoxmlSignatureInput signingInput(Path pkcs12Path, String description) {
    return new OoxmlSignatureInput(
        pkcs12Path.toAbsolutePath().toString(),
        XlsxParityScenarios.SIGNING_KEYSTORE_PASSWORD,
        XlsxParityScenarios.SIGNING_KEY_PASSWORD,
        XlsxParityScenarios.SIGNING_KEY_ALIAS,
        null,
        description);
  }

  static <T> T cast(Class<T> type, Object value) {
    return type.cast(value);
  }

  static boolean hasEncryptedAgilePackage(OoxmlPackageSecurityReport security) {
    return security.encryption().encrypted()
        && security.encryption().mode() == ExcelOoxmlEncryptionMode.AGILE
        && security.signatures().isEmpty();
  }

  static boolean hasSingleSignatureState(
      OoxmlPackageSecurityReport security, ExcelOoxmlSignatureState state) {
    return !security.encryption().encrypted()
        && security.signatures().size() == 1
        && security.signatures().getFirst().state() == state;
  }

  static String textCell(InspectionResult.CellsResult cells, String address) {
    return cast(
            GridGrindResponse.CellReport.TextReport.class, byAddress(cells.cells()).get(address))
        .stringValue();
  }

  static Map<String, GridGrindResponse.CellReport> byAddress(
      List<GridGrindResponse.CellReport> cells) {
    return cells.stream()
        .collect(
            Collectors.toUnmodifiableMap(
                GridGrindResponse.CellReport::address, Function.identity()));
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
      InspectionResult.SheetSummaryResult summary) {
    return switch (summary.sheet().protection()) {
      case GridGrindResponse.SheetProtectionReport.Unprotected _ -> null;
      case GridGrindResponse.SheetProtectionReport.Protected protectedReport ->
          protectedReport.settings();
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
        sheetName,
        new PrintAreaReport.None(),
        ExcelPrintOrientation.PORTRAIT,
        new PrintScalingReport.Automatic(),
        new PrintTitleRowsReport.None(),
        new PrintTitleColumnsReport.None(),
        new HeaderFooterTextReport("", "", ""),
        new HeaderFooterTextReport("", "", ""),
        PrintSetupReport.defaults());
  }

  static PrintLayoutInput advancedPrintLayoutInput() {
    return new PrintLayoutInput(
        new PrintAreaInput.None(),
        ExcelPrintOrientation.PORTRAIT,
        new PrintScalingInput.Automatic(),
        new PrintTitleRowsInput.None(),
        new PrintTitleColumnsInput.None(),
        HeaderFooterTextInput.blank(),
        HeaderFooterTextInput.blank(),
        new PrintSetupInput(
            new PrintMarginsInput(0.35d, 0.55d, 0.60d, 0.45d, 0.3d, 0.3d),
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

  static CellStyleInput advancedThemedStyleInput() {
    return new CellStyleInput(
        null,
        null,
        new CellFontInput(null, true, null, null, null, 6, null, -0.35d, null, null),
        new CellFillInput(
            ExcelFillPattern.SOLID, null, 3, null, 0.30d, null, null, null, null, null),
        new CellBorderInput(
            null,
            null,
            null,
            new CellBorderSideInput(
                ExcelBorderStyle.THIN,
                null,
                null,
                Short.toUnsignedInt(IndexedColors.DARK_RED.getIndex()),
                null),
            null),
        null);
  }

  static CellStyleInput advancedGradientStyleInput() {
    return new CellStyleInput(
        null,
        null,
        null,
        new CellFillInput(
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            new CellGradientFillInput(
                "LINEAR",
                45.0d,
                null,
                null,
                null,
                null,
                List.of(
                    new CellGradientStopInput(0.0d, new ColorInput("#1F497D")),
                    new CellGradientStopInput(1.0d, new ColorInput(null, 4, null, 0.45d))))),
        null,
        null);
  }

  static CommentInput advancedCommentInput() {
    return comment(
        "Lead review scheduled",
        "GridGrind",
        false,
        List.of(
            richTextRun(
                "Lead",
                new CellFontInput(true, null, null, null, null, 4, null, -0.20d, null, null)),
            richTextRun(" ", null),
            richTextRun(
                "review scheduled",
                new CellFontInput(
                    null,
                    true,
                    null,
                    null,
                    null,
                    null,
                    Short.toUnsignedInt(IndexedColors.DARK_GREEN.getIndex()),
                    null,
                    null,
                    null))),
        new CommentAnchorInput(4, 1, 8, 7));
  }

  static DrawingAnchorInput.TwoCell twoCellAnchorInput(
      int fromColumn, int fromRow, int toColumn, int toRow, ExcelDrawingAnchorBehavior behavior) {
    return new DrawingAnchorInput.TwoCell(
        new DrawingMarkerInput(fromColumn, fromRow, 0, 0),
        new DrawingMarkerInput(toColumn, toRow, 0, 0),
        behavior);
  }

  static TableInput advancedTableInput() {
    return new TableInput(
        "AdvancedTable",
        "Advanced",
        "H1:J5",
        true,
        false,
        new TableStyleInput.None(),
        inlineText("Parity advanced table"),
        true,
        true,
        true,
        "ParityHeader",
        "ParityData",
        "ParityTotals",
        List.of(
            new TableColumnInput(0, "", "Total", "", ""),
            new TableColumnInput(1, "", "", "SUM", "[@Sales]*2"),
            new TableColumnInput(2, "owner-unique", "", "", "")));
  }

  static List<PendingMutation> advancedConditionalFormattingMutations() {
    PendingMutation colorScale =
        mutate(
            new SheetSelector.ByName("Advanced"),
            new MutationAction.SetConditionalFormatting(
                new ConditionalFormattingBlockInput(
                    List.of("L2:L5"),
                    List.of(
                        new ConditionalFormattingRuleInput.ColorScaleRule(
                            false,
                            List.of(
                                new ConditionalFormattingThresholdInput(
                                    ExcelConditionalFormattingThresholdType.MIN, null, null),
                                new ConditionalFormattingThresholdInput(
                                    ExcelConditionalFormattingThresholdType.PERCENTILE,
                                    null,
                                    50.0d),
                                new ConditionalFormattingThresholdInput(
                                    ExcelConditionalFormattingThresholdType.MAX, null, null)),
                            List.of(
                                new ColorInput("#AA2211"),
                                new ColorInput("#FFDD55"),
                                new ColorInput("#11CC66")))))));
    PendingMutation dataBar =
        mutate(
            new SheetSelector.ByName("Advanced"),
            new MutationAction.SetConditionalFormatting(
                new ConditionalFormattingBlockInput(
                    List.of("M2:M5"),
                    List.of(
                        new ConditionalFormattingRuleInput.DataBarRule(
                            false,
                            new ColorInput("#123456"),
                            true,
                            10,
                            90,
                            new ConditionalFormattingThresholdInput(
                                ExcelConditionalFormattingThresholdType.MIN, null, null),
                            new ConditionalFormattingThresholdInput(
                                ExcelConditionalFormattingThresholdType.MAX, null, null))))));
    PendingMutation iconSet =
        mutate(
            new SheetSelector.ByName("Advanced"),
            new MutationAction.SetConditionalFormatting(
                new ConditionalFormattingBlockInput(
                    List.of("N2:N5"),
                    List.of(
                        new ConditionalFormattingRuleInput.IconSetRule(
                            false,
                            ExcelConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS,
                            true,
                            true,
                            List.of(
                                new ConditionalFormattingThresholdInput(
                                    ExcelConditionalFormattingThresholdType.PERCENT, null, 0.0d),
                                new ConditionalFormattingThresholdInput(
                                    ExcelConditionalFormattingThresholdType.PERCENT, null, 33.0d),
                                new ConditionalFormattingThresholdInput(
                                    ExcelConditionalFormattingThresholdType.PERCENT,
                                    null,
                                    67.0d)))))));
    PendingMutation top10 =
        mutate(
            new SheetSelector.ByName("Advanced"),
            new MutationAction.SetConditionalFormatting(
                new ConditionalFormattingBlockInput(
                    List.of("K2:K5"),
                    List.of(
                        new ConditionalFormattingRuleInput.Top10Rule(
                            false, 10, false, false, null)))));
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
    return observed.range().equals(direct.range())
        && observed.descending() == direct.descending()
        && observed.sortBy().equals(direct.sortBy());
  }

  static TableEntryReport findTable(List<TableEntryReport> tables, String name) {
    return tables.stream().filter(table -> name.equals(table.name())).findFirst().orElse(null);
  }

  static TableEntryReport findTableOnSheet(List<TableEntryReport> tables, String sheetName) {
    return tables.stream()
        .filter(table -> sheetName.equals(table.sheetName()))
        .findFirst()
        .orElse(null);
  }

  static boolean tableMatchesReadReport(
      TableEntryReport observed,
      TableEntryReport expected,
      String expectedName,
      String expectedSheetName) {
    return observed != null
        && (expectedName == null || observed.name().equals(expectedName))
        && (expectedSheetName == null || observed.sheetName().equals(expectedSheetName))
        && observed.range().equals(expected.range())
        && observed.headerRowCount() == expected.headerRowCount()
        && observed.totalsRowCount() == expected.totalsRowCount()
        && observed.columnNames().equals(expected.columnNames())
        && normalizedTableColumns(observed.columns())
            .equals(normalizedTableColumns(expected.columns()))
        && observed.style().equals(expected.style())
        && observed.hasAutofilter() == expected.hasAutofilter()
        && observed.comment().equals(expected.comment())
        && observed.published() == expected.published()
        && observed.insertRow() == expected.insertRow()
        && observed.insertRowShift() == expected.insertRowShift()
        && observed.headerRowCellStyle().equals(expected.headerRowCellStyle())
        && observed.dataCellStyle().equals(expected.dataCellStyle())
        && observed.totalsRowCellStyle().equals(expected.totalsRowCellStyle());
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
      InspectionResult.DrawingObjectsResult result, String name, Class<T> type) {
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
      dev.erst.gridgrind.contract.dto.CellColorReport color, String descriptor) {
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
        if (!expected.equals(color.rgb())) {
          return false;
        }
      } else if (part.startsWith("theme=")) {
        if (!Integer.valueOf(part.substring("theme=".length())).equals(color.theme())) {
          return false;
        }
      } else if (part.startsWith("indexed=")) {
        if (!Integer.valueOf(part.substring("indexed=".length())).equals(color.indexed())) {
          return false;
        }
      } else if (part.startsWith("tint=")
          && (color.tint() == null
              || !approximatelyEquals(
                  Double.parseDouble(part.substring("tint=".length())), color.tint()))) {
        return false;
      }
    }
    return true;
  }

  static boolean approximatelyEquals(double left, double right) {
    return Math.abs(left - right) <= 0.000001d;
  }

  static CellInput.Text text(String value) {
    return new CellInput.Text(TextSourceInput.inline(value));
  }

  static ChartInput.Title.Text chartTitle(String value) {
    return new ChartInput.Title.Text(TextSourceInput.inline(value));
  }

  static RichTextRunInput richTextRun(String value, CellFontInput font) {
    return new RichTextRunInput(TextSourceInput.inline(value), font);
  }

  static CommentInput comment(String text, String author, boolean visible) {
    return new CommentInput(TextSourceInput.inline(text), author, visible);
  }

  static CommentInput comment(
      String text,
      String author,
      boolean visible,
      List<RichTextRunInput> runs,
      CommentAnchorInput anchor) {
    return new CommentInput(TextSourceInput.inline(text), author, visible, runs, anchor);
  }

  static TextSourceInput inlineText(String value) {
    return TextSourceInput.inline(value);
  }

  /**
   * Incremental-build compatibility alias for the historical monolithic registry layout.
   *
   * <p>The parity probe implementation now lives in focused group classes, but keeping the former
   * nested record names here prevents stale compiled inner classes from surviving no-clean builds
   * after the split.
   */
  record CoreReadObservation(
      XlsxParityOracle.CoreWorkbookSnapshot direct,
      GridGrindResponse.WorkbookSummary.WithSheets summary,
      InspectionResult.SheetSummaryResult opsSummary,
      InspectionResult.SheetSummaryResult queueSummary,
      InspectionResult.CellsResult cells,
      InspectionResult.MergedRegionsResult merged,
      InspectionResult.HyperlinksResult links,
      InspectionResult.CommentsResult comments,
      InspectionResult.SheetLayoutResult layout,
      InspectionResult.PrintLayoutResult print,
      InspectionResult.DataValidationsResult validations,
      InspectionResult.ConditionalFormattingResult formatting,
      InspectionResult.AutofiltersResult filtersOps,
      InspectionResult.AutofiltersResult filtersQueue,
      InspectionResult.TablesResult tables) {}

  /** Incremental-build compatibility alias for the historical monolithic registry layout. */
  record SheetCopySourceObservation(
      XlsxParityScenarios.MaterializedScenario source,
      InspectionResult.CommentsResult sourceComments,
      InspectionResult.AutofiltersResult sourceAutofilters,
      InspectionResult.ConditionalFormattingResult sourceFormatting,
      InspectionResult.SheetSummaryResult sourceSummary,
      TableEntryReport sourceTable,
      XlsxParityOracle.CommentSnapshot sourceComment,
      XlsxParityOracle.AdvancedPrintSnapshot sourcePrint,
      XlsxParityOracle.AutofilterSnapshot sourceAutofilter,
      XlsxParityOracle.TableSnapshot sourceDirectTable) {}

  /** Incremental-build compatibility alias for the historical monolithic registry layout. */
  record SheetCopyCopiedObservation(
      Path copiedPath,
      InspectionResult.WorkbookSummaryResult copiedWorkbook,
      InspectionResult.SheetSummaryResult copiedSummary,
      InspectionResult.CommentsResult copiedComments,
      InspectionResult.AutofiltersResult copiedAutofilters,
      InspectionResult.ConditionalFormattingResult copiedFormatting,
      TableEntryReport copiedTable,
      XlsxParityOracle.CommentSnapshot copiedComment,
      XlsxParityOracle.AdvancedPrintSnapshot copiedPrint,
      XlsxParityOracle.AutofilterSnapshot copiedAutofilter,
      XlsxParityOracle.TableSnapshot copiedDirectTable) {}

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
