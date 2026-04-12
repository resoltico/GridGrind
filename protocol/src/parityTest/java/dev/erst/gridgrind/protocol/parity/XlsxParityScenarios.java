package dev.erst.gridgrind.protocol.parity;

import dev.erst.gridgrind.excel.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.ExcelFillPattern;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelPictureFormat;
import dev.erst.gridgrind.excel.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import dev.erst.gridgrind.protocol.dto.*;
import dev.erst.gridgrind.protocol.dto.GridGrindRequest;
import dev.erst.gridgrind.protocol.dto.GridGrindResponse;
import dev.erst.gridgrind.protocol.exec.DefaultGridGrindRequestExecutor;
import dev.erst.gridgrind.protocol.json.GridGrindJson;
import dev.erst.gridgrind.protocol.operation.WorkbookOperation;
import java.io.ByteArrayOutputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.KeyPair;
import java.security.KeyPairGenerator;
import java.security.SecureRandom;
import java.security.Security;
import java.security.cert.X509Certificate;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Base64;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.crypt.HashAlgorithm;
import org.apache.poi.poifs.crypt.dsig.SignatureConfig;
import org.apache.poi.poifs.crypt.dsig.SignatureInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IconMultiStateFormatting;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.BarDirection;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;
import org.bouncycastle.asn1.x500.X500Name;
import org.bouncycastle.cert.X509CertificateHolder;
import org.bouncycastle.cert.jcajce.JcaX509CertificateConverter;
import org.bouncycastle.cert.jcajce.JcaX509v3CertificateBuilder;
import org.bouncycastle.jce.provider.BouncyCastleProvider;
import org.bouncycastle.operator.ContentSigner;
import org.bouncycastle.operator.jcajce.JcaContentSignerBuilder;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidations;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCfType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STTotalsRowFunction;

/** Materializes the committed parity workbook corpus into deterministic temporary files. */
public final class XlsxParityScenarios {
  static final String CORE_WORKBOOK = "gridgrind-core-authoring";
  static final String ADVANCED_NONDRAWING = "poi-advanced-nondrawing";
  static final String EXTERNAL_FORMULA = "poi-external-formula";
  static final String UDF_FORMULA = "poi-udf-formula";
  static final String FORMULA_LIFECYCLE = "poi-formula-lifecycle";
  static final String DRAWING_AUTHORING = "gridgrind-drawing-authoring";
  static final String DRAWING_IMAGE = "poi-drawing-image";
  static final String DRAWING_COMMENTS = "poi-drawing-comments";
  static final String DRAWING_MERGED_IMAGE = "poi-drawing-merged-image";
  static final String CHART = "poi-chart";
  static final String CHART_AUTHORING = "gridgrind-chart-authoring";
  static final String CHART_UNSUPPORTED = "poi-chart-unsupported";
  static final String PIVOT = "poi-pivot";
  static final String PIVOT_AUTHORING = "gridgrind-pivot-authoring";
  static final String EMBEDDED_OBJECT = "poi-embedded-object";
  static final String LARGE_SHEET = "poi-large-sheet";
  static final String SIGNED_WORKBOOK = "poi-signed-workbook";
  static final String ENCRYPTED_WORKBOOK = "poi-encrypted-workbook";
  static final String WORKBOOK_PROTECTION_PASSWORD = "gridgrind-phase1-workbook";
  static final String SHEET_PROTECTION_PASSWORD = "gridgrind-phase1-sheet";
  static final String ENCRYPTION_PASSWORD = "gridgrind-phase1-encrypted";
  static final String SIGNING_SUBJECT = "CN=GridGrind Phase 1 Parity";

  private XlsxParityScenarios() {}

  /** Materializes one committed parity scenario under the supplied temporary root directory. */
  public static MaterializedScenario materialize(String scenarioId, Path temporaryRoot) {
    Objects.requireNonNull(scenarioId, "scenarioId must not be null");
    Objects.requireNonNull(temporaryRoot, "temporaryRoot must not be null");
    return XlsxParitySupport.call(
        "materialize parity scenario " + scenarioId,
        () ->
            switch (scenarioId) {
              case CORE_WORKBOOK -> materializeCoreWorkbook(temporaryRoot);
              case ADVANCED_NONDRAWING -> materializeAdvancedNonDrawingWorkbook(temporaryRoot);
              case EXTERNAL_FORMULA -> materializeExternalFormulaWorkbook(temporaryRoot);
              case UDF_FORMULA -> materializeUdfFormulaWorkbook(temporaryRoot);
              case FORMULA_LIFECYCLE -> materializeFormulaLifecycleWorkbook(temporaryRoot);
              case DRAWING_AUTHORING -> materializeDrawingAuthoringWorkbook(temporaryRoot);
              case DRAWING_IMAGE -> materializeDrawingWorkbook(temporaryRoot);
              case DRAWING_COMMENTS -> materializeDrawingCommentsWorkbook(temporaryRoot);
              case DRAWING_MERGED_IMAGE -> materializeDrawingMergedImageWorkbook(temporaryRoot);
              case CHART -> materializeChartWorkbook(temporaryRoot);
              case CHART_AUTHORING -> materializeChartAuthoringWorkbook(temporaryRoot);
              case CHART_UNSUPPORTED -> materializeUnsupportedChartWorkbook(temporaryRoot);
              case PIVOT -> materializePivotWorkbook(temporaryRoot);
              case PIVOT_AUTHORING -> materializePivotAuthoringWorkbook(temporaryRoot);
              case EMBEDDED_OBJECT -> materializeEmbeddedObjectWorkbook(temporaryRoot);
              case LARGE_SHEET -> materializeLargeSheetWorkbook(temporaryRoot);
              case SIGNED_WORKBOOK -> materializeSignedWorkbook(temporaryRoot);
              case ENCRYPTED_WORKBOOK -> materializeEncryptedWorkbook(temporaryRoot);
              default ->
                  throw new IllegalArgumentException("Unknown parity scenario id: " + scenarioId);
            });
  }

  /** One materialized corpus workbook plus any auxiliary files needed by its probe. */
  public record MaterializedScenario(Path workbookPath, Map<String, Path> attachments) {
    public MaterializedScenario {
      Objects.requireNonNull(workbookPath, "workbookPath must not be null");
      attachments = Map.copyOf(Objects.requireNonNull(attachments, "attachments must not be null"));
    }

    /** Returns the attachment path with the given stable key. */
    public Path attachment(String key) {
      Objects.requireNonNull(key, "key must not be null");
      Path path = attachments.get(key);
      if (path == null) {
        throw new IllegalArgumentException("Unknown scenario attachment: " + key);
      }
      return path;
    }
  }

  private static MaterializedScenario materializeCoreWorkbook(Path temporaryRoot) {
    return XlsxParitySupport.call(
        "materialize core parity workbook",
        () -> {
          Path scenarioDirectory = Files.createDirectories(temporaryRoot.resolve(CORE_WORKBOOK));
          Path workbookPath = scenarioDirectory.resolve("core.xlsx");
          GridGrindRequest request = coreWorkbookRequest(workbookPath);
          GridGrindResponse response = new DefaultGridGrindRequestExecutor().execute(request);
          if (!(response instanceof GridGrindResponse.Success)) {
            throw new IllegalStateException(
                "Core parity workbook request must succeed: " + response);
          }
          return new MaterializedScenario(workbookPath, Map.of());
        });
  }

  private static MaterializedScenario materializeAdvancedNonDrawingWorkbook(Path temporaryRoot) {
    return XlsxParitySupport.call(
        "materialize advanced non-drawing parity workbook",
        () -> {
          Path scenarioDirectory =
              Files.createDirectories(temporaryRoot.resolve(ADVANCED_NONDRAWING));
          Path workbookPath = scenarioDirectory.resolve("advanced-nondrawing.xlsx");

          try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Advanced");
            seedMatrix(sheet, 12, 14);

            workbook.lockStructure();
            workbook.lockWindows();
            workbook.lockRevision();
            workbook.setWorkbookPassword(WORKBOOK_PROTECTION_PASSWORD, HashAlgorithm.sha512);

            sheet.protectSheet(SHEET_PROTECTION_PASSWORD);
            sheet.setSheetPassword(SHEET_PROTECTION_PASSWORD, HashAlgorithm.sha512);

            sheet.setMargin(XSSFSheet.LeftMargin, 0.35d);
            sheet.setMargin(XSSFSheet.RightMargin, 0.55d);
            sheet.setMargin(XSSFSheet.TopMargin, 0.60d);
            sheet.setMargin(XSSFSheet.BottomMargin, 0.45d);
            sheet.setHorizontallyCenter(true);
            sheet.setVerticallyCenter(true);
            sheet.setDisplayGridlines(false);
            sheet.setRowBreak(6);
            sheet.setColumnBreak(3);
            sheet.getPrintSetup().setPaperSize((short) 8);
            sheet.getPrintSetup().setDraft(true);
            sheet.getPrintSetup().setNoColor(true);
            sheet.getPrintSetup().setCopies((short) 2);
            sheet.getPrintSetup().setUsePage(true);
            sheet.getPrintSetup().setPageStart((short) 4);

            XSSFComment comment = advancedComment(workbook, sheet);
            sheet.getRow(1).getCell(4).setCellComment(comment);

            NameSupport.addFormulaDefinedName(
                workbook, "WorkbookFormulaBudget", null, "SUM(Advanced!$B$2:$B$5)");
            NameSupport.addFormulaDefinedName(
                workbook,
                "SheetScopedFormulaBudget",
                workbook.getSheetIndex(sheet),
                "SUM(Advanced!$C$2:$C$5)");

            createAdvancedStyles(workbook, sheet);
            createAdvancedDataValidation(sheet);
            createAdvancedAutofilter(sheet);
            createAdvancedConditionalFormatting(sheet);
            createAdvancedTable(sheet);

            try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
              workbook.write(outputStream);
            }
          }

          return new MaterializedScenario(workbookPath, Map.of());
        });
  }

  private static MaterializedScenario materializeExternalFormulaWorkbook(Path temporaryRoot) {
    return XlsxParitySupport.call(
        "materialize external-formula parity workbook",
        () -> {
          Path scenarioDirectory = Files.createDirectories(temporaryRoot.resolve(EXTERNAL_FORMULA));
          Path referencedWorkbookPath = scenarioDirectory.resolve("referenced.xlsx");
          Path workbookPath = scenarioDirectory.resolve("external-formula.xlsx");

          try (XSSFWorkbook referencedWorkbook = new XSSFWorkbook()) {
            XSSFSheet referencedSheet = referencedWorkbook.createSheet("Rates");
            referencedSheet.createRow(0).createCell(0).setCellValue(7.5d);
            try (OutputStream outputStream = Files.newOutputStream(referencedWorkbookPath)) {
              referencedWorkbook.write(outputStream);
            }
          }

          try (Workbook referencedWorkbook =
                  WorkbookFactory.create(referencedWorkbookPath.toFile());
              XSSFWorkbook workbook = new XSSFWorkbook()) {
            workbook.linkExternalWorkbook(
                referencedWorkbookPath.getFileName().toString(), referencedWorkbook);
            XSSFSheet sheet = workbook.createSheet("Ops");
            sheet.createRow(0).createCell(0).setCellValue("External");
            workbook.setCellFormulaValidation(false);
            sheet
                .getRow(0)
                .createCell(1)
                .setCellFormula("[" + referencedWorkbookPath.getFileName() + "]Rates!$A$1");
            var evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.setupReferencedWorkbooks(
                Map.of(
                    referencedWorkbookPath.getFileName().toString(),
                    referencedWorkbook.getCreationHelper().createFormulaEvaluator()));
            evaluator.evaluateFormulaCell(sheet.getRow(0).getCell(1));
            try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
              workbook.write(outputStream);
            }
          }

          return new MaterializedScenario(
              workbookPath, Map.of("referencedWorkbook", referencedWorkbookPath));
        });
  }

  private static MaterializedScenario materializeFormulaLifecycleWorkbook(Path temporaryRoot) {
    return XlsxParitySupport.call(
        "materialize formula-lifecycle parity workbook",
        () -> {
          Path scenarioDirectory =
              Files.createDirectories(temporaryRoot.resolve(FORMULA_LIFECYCLE));
          Path workbookPath = scenarioDirectory.resolve("formula-lifecycle.xlsx");

          try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Budget");
            sheet.createRow(0).createCell(0).setCellValue(2.0d);
            sheet.getRow(0).createCell(1).setCellFormula("A1*2");
            sheet.getRow(0).createCell(2).setCellFormula("A1*3");
            workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
            sheet.getRow(0).getCell(0).setCellValue(4.0d);
            try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
              workbook.write(outputStream);
            }
          }

          return new MaterializedScenario(workbookPath, Map.of());
        });
  }

  private static MaterializedScenario materializeUdfFormulaWorkbook(Path temporaryRoot) {
    return XlsxParitySupport.call(
        "materialize UDF parity workbook",
        () -> {
          Path scenarioDirectory = Files.createDirectories(temporaryRoot.resolve(UDF_FORMULA));
          Path workbookPath = scenarioDirectory.resolve("udf-formula.xlsx");

          try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            workbook.setCellFormulaValidation(false);
            XSSFSheet sheet = workbook.createSheet("Ops");
            Row row = sheet.createRow(0);
            row.createCell(0).setCellValue(21d);
            row.createCell(1).setCellFormula("DOUBLE(A1)");
            try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
              workbook.write(outputStream);
            }
          }

          return new MaterializedScenario(workbookPath, Map.of());
        });
  }

  private static MaterializedScenario materializeDrawingWorkbook(Path temporaryRoot) {
    return XlsxParitySupport.call(
        "materialize drawing parity workbook",
        () -> {
          Path scenarioDirectory = Files.createDirectories(temporaryRoot.resolve(DRAWING_IMAGE));
          Path workbookPath = scenarioDirectory.resolve("drawing.xlsx");

          try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Ops");
            seedMatrix(sheet, 4, 2);
            int pictureIndex = workbook.addPicture(PNG_PIXEL_BYTES, Workbook.PICTURE_TYPE_PNG);
            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 3, 1, 6, 8);
            drawing
                .createPicture(anchor, pictureIndex)
                .getCTPicture()
                .getNvPicPr()
                .getCNvPr()
                .setName("OpsPicture");
            try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
              workbook.write(outputStream);
            }
          }

          return new MaterializedScenario(workbookPath, Map.of());
        });
  }

  private static MaterializedScenario materializeDrawingCommentsWorkbook(Path temporaryRoot) {
    return XlsxParitySupport.call(
        "materialize drawing-comments parity workbook",
        () -> {
          Path scenarioDirectory = Files.createDirectories(temporaryRoot.resolve(DRAWING_COMMENTS));
          Path workbookPath = scenarioDirectory.resolve("drawing-comments.xlsx");

          try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Ops");
            seedMatrix(sheet, 4, 4);
            int pictureIndex = workbook.addPicture(PNG_PIXEL_BYTES, Workbook.PICTURE_TYPE_PNG);
            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor pictureAnchor = drawing.createAnchor(0, 0, 0, 0, 3, 1, 6, 8);
            drawing
                .createPicture(pictureAnchor, pictureIndex)
                .getCTPicture()
                .getNvPicPr()
                .getCNvPr()
                .setName("OpsPicture");
            XSSFClientAnchor commentAnchor = drawing.createAnchor(64, 24, 448, 96, 0, 0, 3, 3);
            XSSFComment comment = drawing.createCellComment(commentAnchor);
            comment.setString(new XSSFRichTextString("Pinned"));
            comment.setAuthor("GridGrind");
            sheet.getRow(0).getCell(0).setCellComment(comment);

            try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
              workbook.write(outputStream);
            }
          }

          return new MaterializedScenario(workbookPath, Map.of());
        });
  }

  private static MaterializedScenario materializeDrawingMergedImageWorkbook(Path temporaryRoot) {
    return XlsxParitySupport.call(
        "materialize drawing-merged-image parity workbook",
        () -> {
          Path scenarioDirectory =
              Files.createDirectories(temporaryRoot.resolve(DRAWING_MERGED_IMAGE));
          Path workbookPath = scenarioDirectory.resolve("drawing-merged-image.xlsx");

          try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Ops");
            seedMatrix(sheet, 8, 6);
            sheet.addMergedRegion(CellRangeAddress.valueOf("B2:D4"));
            int pictureIndex = workbook.addPicture(PNG_PIXEL_BYTES, Workbook.PICTURE_TYPE_PNG);
            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 1, 1, 5, 7);
            drawing
                .createPicture(anchor, pictureIndex)
                .getCTPicture()
                .getNvPicPr()
                .getCNvPr()
                .setName("OpsPicture");
            try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
              workbook.write(outputStream);
            }
          }

          return new MaterializedScenario(workbookPath, Map.of());
        });
  }

  private static MaterializedScenario materializeChartWorkbook(Path temporaryRoot) {
    return XlsxParitySupport.call(
        "materialize chart parity workbook",
        () -> {
          Path scenarioDirectory = Files.createDirectories(temporaryRoot.resolve(CHART));
          Path workbookPath = scenarioDirectory.resolve("chart.xlsx");

          try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Chart");
            seedChartData(sheet);

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 3, 0, 10, 18);
            var chart = drawing.createChart(anchor);
            chart.getGraphicFrame().setName("OpsChart");
            chart.setTitleText("Roadmap");
            chart.getOrAddLegend().setPosition(LegendPosition.TOP_RIGHT);
            XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
            valueAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            var categories =
                XDDFDataSourcesFactory.fromStringCellRange(
                    sheet, CellRangeAddress.valueOf("A2:A4"));
            XDDFNumericalDataSource<Double> values =
                XDDFDataSourcesFactory.fromNumericCellRange(
                    sheet, CellRangeAddress.valueOf("B2:B4"));
            XDDFBarChartData chartData =
                (XDDFBarChartData) chart.createData(ChartTypes.BAR, categoryAxis, valueAxis);
            chartData.setBarDirection(BarDirection.COL);
            chartData.setVaryColors(true);
            var series = chartData.addSeries(categories, values);
            series.setTitle("B1", null);
            chart.displayBlanksAs(org.apache.poi.xddf.usermodel.chart.DisplayBlanks.SPAN);
            chart.setPlotOnlyVisibleCells(false);
            chart.plot(chartData);

            try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
              workbook.write(outputStream);
            }
          }

          return new MaterializedScenario(workbookPath, Map.of());
        });
  }

  private static MaterializedScenario materializeChartAuthoringWorkbook(Path temporaryRoot) {
    return XlsxParitySupport.call(
        "materialize chart authoring parity workbook",
        () -> {
          Path scenarioDirectory = Files.createDirectories(temporaryRoot.resolve(CHART_AUTHORING));
          Path workbookPath = scenarioDirectory.resolve("chart-authoring.xlsx");

          GridGrindResponse response =
              new DefaultGridGrindRequestExecutor().execute(chartAuthoringRequest(workbookPath));
          if (!(response instanceof GridGrindResponse.Success)) {
            throw new IllegalStateException(
                "Chart authoring parity workbook request must succeed: " + response);
          }

          return new MaterializedScenario(workbookPath, Map.of());
        });
  }

  private static MaterializedScenario materializeUnsupportedChartWorkbook(Path temporaryRoot) {
    return XlsxParitySupport.call(
        "materialize unsupported chart parity workbook",
        () -> {
          Path scenarioDirectory =
              Files.createDirectories(temporaryRoot.resolve(CHART_UNSUPPORTED));
          Path workbookPath = scenarioDirectory.resolve("chart-unsupported.xlsx");

          try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Chart");
            seedChartData(sheet);

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 1, 11, 16);
            var chart = drawing.createChart(anchor);
            chart.getGraphicFrame().setName("ComboChart");
            chart.getOrAddLegend().setPosition(LegendPosition.TOP_RIGHT);
            XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
            valueAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            var categories =
                XDDFDataSourcesFactory.fromStringCellRange(
                    sheet, CellRangeAddress.valueOf("A2:A4"));
            var barValues =
                XDDFDataSourcesFactory.fromNumericCellRange(
                    sheet, CellRangeAddress.valueOf("B2:B4"));
            var lineValues =
                XDDFDataSourcesFactory.fromNumericCellRange(
                    sheet, CellRangeAddress.valueOf("C2:C4"));
            XDDFBarChartData barData =
                (XDDFBarChartData) chart.createData(ChartTypes.BAR, categoryAxis, valueAxis);
            barData.addSeries(categories, barValues).setTitle("Plan", null);
            chart.plot(barData);
            var lineData = chart.createData(ChartTypes.LINE, categoryAxis, valueAxis);
            lineData.addSeries(categories, lineValues).setTitle("Actual", null);
            chart.plot(lineData);

            try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
              workbook.write(outputStream);
            }
          }

          return new MaterializedScenario(workbookPath, Map.of());
        });
  }

  private static MaterializedScenario materializePivotWorkbook(Path temporaryRoot) {
    return XlsxParitySupport.call(
        "materialize pivot parity workbook",
        () -> {
          Path scenarioDirectory = Files.createDirectories(temporaryRoot.resolve(PIVOT));
          Path workbookPath = scenarioDirectory.resolve("pivot.xlsx");

          try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet dataSheet = workbook.createSheet("Data");
            dataSheet.createRow(0).createCell(0).setCellValue("Region");
            dataSheet.getRow(0).createCell(1).setCellValue("Stage");
            dataSheet.getRow(0).createCell(2).setCellValue("Owner");
            dataSheet.getRow(0).createCell(3).setCellValue("Amount");
            dataSheet.createRow(1).createCell(0).setCellValue("North");
            dataSheet.getRow(1).createCell(1).setCellValue("Plan");
            dataSheet.getRow(1).createCell(2).setCellValue("Ada");
            dataSheet.getRow(1).createCell(3).setCellValue(10d);
            dataSheet.createRow(2).createCell(0).setCellValue("South");
            dataSheet.getRow(2).createCell(1).setCellValue("Do");
            dataSheet.getRow(2).createCell(2).setCellValue("Lin");
            dataSheet.getRow(2).createCell(3).setCellValue(14d);
            dataSheet.createRow(3).createCell(0).setCellValue("North");
            dataSheet.getRow(3).createCell(1).setCellValue("Do");
            dataSheet.getRow(3).createCell(2).setCellValue("Ada");
            dataSheet.getRow(3).createCell(3).setCellValue(7d);
            dataSheet.createRow(4).createCell(0).setCellValue("South");
            dataSheet.getRow(4).createCell(1).setCellValue("Plan");
            dataSheet.getRow(4).createCell(2).setCellValue("Lin");
            dataSheet.getRow(4).createCell(3).setCellValue(12d);

            XSSFSheet pivotSheet = workbook.createSheet("Pivot");
            AreaReference source = new AreaReference("A1:D5", SpreadsheetVersion.EXCEL2007);
            XSSFPivotTable pivotTable =
                pivotSheet.createPivotTable(source, new CellReference("C5"), dataSheet);
            pivotTable.getCTPivotTableDefinition().setName("POI Pivot");
            pivotTable.addRowLabel(0);
            pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 3, "Total Amount");

            try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
              workbook.write(outputStream);
            }
          }

          return new MaterializedScenario(workbookPath, Map.of());
        });
  }

  private static MaterializedScenario materializePivotAuthoringWorkbook(Path temporaryRoot) {
    return XlsxParitySupport.call(
        "materialize pivot authoring parity workbook",
        () -> {
          Path scenarioDirectory = Files.createDirectories(temporaryRoot.resolve(PIVOT_AUTHORING));
          Path workbookPath = scenarioDirectory.resolve("pivot-authoring.xlsx");

          GridGrindResponse response =
              new DefaultGridGrindRequestExecutor().execute(pivotAuthoringRequest(workbookPath));
          if (!(response instanceof GridGrindResponse.Success)) {
            throw new IllegalStateException(
                "Pivot authoring parity workbook request must succeed: " + response);
          }

          return new MaterializedScenario(workbookPath, Map.of());
        });
  }

  private static MaterializedScenario materializeEmbeddedObjectWorkbook(Path temporaryRoot) {
    return XlsxParitySupport.call(
        "materialize embedded-object parity workbook",
        () -> {
          Path scenarioDirectory = Files.createDirectories(temporaryRoot.resolve(EMBEDDED_OBJECT));
          Path workbookPath = scenarioDirectory.resolve("embedded-object.xlsx");

          try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Objects");
            seedMatrix(sheet, 4, 4);
            int pictureIndex = workbook.addPicture(PNG_PIXEL_BYTES, Workbook.PICTURE_TYPE_PNG);
            int storageId =
                workbook.addOlePackage(
                    "GridGrind embedded payload".getBytes(StandardCharsets.UTF_8),
                    "Parity Payload",
                    "payload.txt",
                    "payload.txt");
            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 1, 1, 5, 10);
            drawing
                .createObjectData(anchor, storageId, pictureIndex)
                .getCTShape()
                .getNvSpPr()
                .getCNvPr()
                .setName("OpsEmbed");
            try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
              workbook.write(outputStream);
            }
          }

          return new MaterializedScenario(workbookPath, Map.of());
        });
  }

  private static MaterializedScenario materializeDrawingAuthoringWorkbook(Path temporaryRoot) {
    return XlsxParitySupport.call(
        "materialize GridGrind drawing authoring workbook",
        () -> {
          Path scenarioDirectory =
              Files.createDirectories(temporaryRoot.resolve(DRAWING_AUTHORING));
          Path workbookPath = scenarioDirectory.resolve("drawing-authoring.xlsx");
          GridGrindResponse response =
              new DefaultGridGrindRequestExecutor().execute(drawingAuthoringRequest(workbookPath));
          if (!(response instanceof GridGrindResponse.Success)) {
            throw new IllegalStateException(
                "GridGrind drawing authoring parity workbook request must succeed: " + response);
          }
          return new MaterializedScenario(workbookPath, Map.of());
        });
  }

  private static MaterializedScenario materializeLargeSheetWorkbook(Path temporaryRoot) {
    return XlsxParitySupport.call(
        "materialize large-sheet parity workbook",
        () -> {
          Path scenarioDirectory = Files.createDirectories(temporaryRoot.resolve(LARGE_SHEET));
          Path workbookPath = scenarioDirectory.resolve("large.xlsx");

          try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Large");
            for (int rowIndex = 0; rowIndex < 2_000; rowIndex++) {
              Row row = sheet.createRow(rowIndex);
              for (int columnIndex = 0; columnIndex < 20; columnIndex++) {
                row.createCell(columnIndex).setCellValue("R" + rowIndex + "C" + columnIndex);
              }
            }
            try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
              workbook.write(outputStream);
            }
          }

          return new MaterializedScenario(workbookPath, Map.of());
        });
  }

  private static MaterializedScenario materializeSignedWorkbook(Path temporaryRoot) {
    return XlsxParitySupport.call(
        "materialize signed parity workbook",
        () -> {
          Path scenarioDirectory = Files.createDirectories(temporaryRoot.resolve(SIGNED_WORKBOOK));
          Path workbookPath = scenarioDirectory.resolve("signed.xlsx");

          try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Signed");
            sheet.createRow(0).createCell(0).setCellValue("Signed workbook");
            sheet.getRow(0).createCell(1).setCellValue(2026d);
            try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
              workbook.write(outputStream);
            }
          }

          SelfSignedMaterial keyMaterial = createSelfSignedCertificate();
          try (OPCPackage pkg = OPCPackage.open(workbookPath.toFile(), PackageAccess.READ_WRITE)) {
            SignatureConfig signatureConfig = new SignatureConfig();
            signatureConfig.setExecutionTime(
                java.util.Date.from(Instant.parse("2026-04-10T00:00:00Z")));
            signatureConfig.setKey(keyMaterial.keyPair().getPrivate());
            signatureConfig.setSigningCertificateChain(List.of(keyMaterial.certificate()));

            SignatureInfo signatureInfo = new SignatureInfo();
            signatureInfo.setSignatureConfig(signatureConfig);
            signatureInfo.setOpcPackage(pkg);
            signatureInfo.confirmSignature();
            if (!signatureInfo.verifySignature()) {
              throw new IllegalStateException("Signed parity workbook must validate immediately");
            }
          }

          return new MaterializedScenario(workbookPath, Map.of());
        });
  }

  private static MaterializedScenario materializeEncryptedWorkbook(Path temporaryRoot) {
    return XlsxParitySupport.call(
        "materialize encrypted parity workbook",
        () -> {
          Path scenarioDirectory =
              Files.createDirectories(temporaryRoot.resolve(ENCRYPTED_WORKBOOK));
          Path workbookPath = scenarioDirectory.resolve("encrypted.xlsx");

          byte[] workbookBytes;
          try (XSSFWorkbook workbook = new XSSFWorkbook();
              ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            XSSFSheet sheet = workbook.createSheet("Encrypted");
            sheet.createRow(0).createCell(0).setCellValue("Encrypted workbook");
            sheet.getRow(0).createCell(1).setCellValue(17d);
            workbook.write(outputStream);
            workbookBytes = outputStream.toByteArray();
          }

          EncryptionInfo encryptionInfo = new EncryptionInfo(EncryptionMode.agile);
          Encryptor encryptor = encryptionInfo.getEncryptor();
          encryptor.confirmPassword(ENCRYPTION_PASSWORD);
          try (POIFSFileSystem fileSystem = new POIFSFileSystem()) {
            try (OutputStream encryptedStream = encryptor.getDataStream(fileSystem)) {
              encryptedStream.write(workbookBytes);
            }
            try (OutputStream fileOutputStream = Files.newOutputStream(workbookPath)) {
              fileSystem.writeFilesystem(fileOutputStream);
            }
          }

          return new MaterializedScenario(workbookPath, Map.of());
        });
  }

  private static XSSFComment advancedComment(XSSFWorkbook workbook, XSSFSheet sheet) {
    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    XSSFClientAnchor anchor = drawing.createAnchor(64, 24, 448, 96, 4, 1, 8, 7);
    XSSFComment comment = drawing.createCellComment(anchor);
    XSSFRichTextString richText = new XSSFRichTextString("Lead review scheduled");
    XSSFFont boldFont = workbook.createFont();
    boldFont.setBold(true);
    XSSFColor boldColor = new XSSFColor();
    boldColor.setTheme(4);
    boldColor.setTint(-0.20d);
    boldFont.getCTFont().addNewColor().set(boldColor.getCTColor());
    XSSFFont italicFont = workbook.createFont();
    italicFont.setItalic(true);
    XSSFColor italicColor = new XSSFColor();
    italicColor.setIndexed(IndexedColors.DARK_GREEN.getIndex());
    italicFont.getCTFont().addNewColor().set(italicColor.getCTColor());
    richText.applyFont(0, 4, boldFont);
    richText.applyFont(5, richText.length(), italicFont);
    comment.setString(richText);
    comment.setAuthor("GridGrind");
    comment.setVisible(false);
    return comment;
  }

  private static void createAdvancedStyles(XSSFWorkbook workbook, XSSFSheet sheet) {
    XlsxParitySupport.call(
        "create advanced parity styles",
        () -> {
          sheet.getRow(2).getCell(0).setCellValue("ThemeTintStyle");
          sheet.getRow(3).getCell(0).setCellValue("GradientFillStyle");

          XSSFCellStyle themedStyle = workbook.createCellStyle();
          XSSFFont themedFont = workbook.createFont();
          themedFont.setItalic(true);
          var themedFontColor = themedFont.getCTFont().addNewColor();
          themedFontColor.setTheme(6L);
          themedFontColor.setTint(-0.35d);
          themedStyle.setFont(themedFont);
          themedStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
          XSSFColor themedFillColor = new XSSFColor();
          themedFillColor.setTheme(3);
          themedFillColor.setTint(0.30d);
          themedStyle.setFillForegroundColor(themedFillColor);
          themedStyle.setBorderBottom(BorderStyle.THIN);
          XSSFColor indexedBorderColor = new XSSFColor();
          indexedBorderColor.setIndexed(IndexedColors.DARK_RED.getIndex());
          themedStyle.setBottomBorderColor(indexedBorderColor);
          sheet.getRow(2).getCell(0).setCellStyle(themedStyle);

          StylesTable stylesTable = workbook.getStylesSource();
          XSSFCellStyle gradientStyle = workbook.createCellStyle();
          CTFill gradientFill = CTFill.Factory.newInstance();
          var gradient = gradientFill.addNewGradientFill();
          gradient.setDegree(45d);
          var firstStop = gradient.addNewStop();
          firstStop.setPosition(0d);
          firstStop.addNewColor().setRgb(new byte[] {(byte) 0xFF, 0x1F, 0x49, 0x7D});
          var secondStop = gradient.addNewStop();
          secondStop.setPosition(1d);
          var secondStopColor = secondStop.addNewColor();
          secondStopColor.setTheme(4L);
          secondStopColor.setTint(0.45d);
          int gradientFillId =
              stylesTable.putFill(new XSSFCellFill(gradientFill, stylesTable.getIndexedColors()));
          gradientStyle.getCoreXf().setApplyFill(true);
          gradientStyle.getCoreXf().setFillId(gradientFillId);
          sheet.getRow(3).getCell(0).setCellStyle(gradientStyle);
          return null;
        });
  }

  private static void createAdvancedDataValidation(XSSFSheet sheet) {
    CTDataValidations dataValidations =
        sheet.getCTWorksheet().isSetDataValidations()
            ? sheet.getCTWorksheet().getDataValidations()
            : sheet.getCTWorksheet().addNewDataValidations();

    CTDataValidation invalidExplicitList = dataValidations.addNewDataValidation();
    invalidExplicitList.setType(STDataValidationType.LIST);
    invalidExplicitList.setSqref(List.of("E2"));
    invalidExplicitList.setFormula1("\"\"");

    CTDataValidation missingFormula = dataValidations.addNewDataValidation();
    missingFormula.setType(STDataValidationType.LIST);
    missingFormula.setSqref(List.of("F2"));

    dataValidations.setCount((long) dataValidations.sizeOfDataValidationArray());
  }

  private static void createAdvancedAutofilter(XSSFSheet sheet) {
    XlsxParitySupport.call(
        "create advanced parity autofilter",
        () -> {
          sheet.setAutoFilter(CellRangeAddress.valueOf("A1:C5"));
          CTAutoFilter autoFilter = sheet.getCTWorksheet().getAutoFilter();
          var filterColumn = autoFilter.addNewFilterColumn();
          filterColumn.setColId(0L);
          filterColumn.addNewFilters().addNewFilter().setVal("R1C0");
          var sortState = autoFilter.addNewSortState();
          sortState.setRef("A2:C5");
          sortState.addNewSortCondition().setRef("B2:B5");
          sortState.getSortConditionArray(0).setDescending(true);
          return null;
        });
  }

  private static void createAdvancedConditionalFormatting(XSSFSheet sheet) {
    var formatting = sheet.getSheetConditionalFormatting();
    XSSFConditionalFormattingRule colorScaleRule =
        formatting.createConditionalFormattingColorScaleRule();
    var colorScale = colorScaleRule.getColorScaleFormatting();
    colorScale.setNumControlPoints(3);
    colorScale.getThresholds()[0].setRangeType(ConditionalFormattingThreshold.RangeType.MIN);
    colorScale.getThresholds()[1].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENTILE);
    colorScale.getThresholds()[1].setValue(50d);
    colorScale.getThresholds()[2].setRangeType(ConditionalFormattingThreshold.RangeType.MAX);
    colorScale.setColors(
        new XSSFColor[] {
          new XSSFColor(new byte[] {(byte) 0xAA, 0x22, 0x11}),
          new XSSFColor(new byte[] {(byte) 0xFF, (byte) 0xDD, 0x55}),
          new XSSFColor(new byte[] {0x11, (byte) 0xCC, 0x66})
        });
    formatting.addConditionalFormatting(
        new CellRangeAddress[] {CellRangeAddress.valueOf("L2:L5")}, colorScaleRule);

    XSSFConditionalFormattingRule dataBarRule =
        formatting.createConditionalFormattingRule(new XSSFColor(new byte[] {0x12, 0x34, 0x56}));
    var dataBar = dataBarRule.getDataBarFormatting();
    dataBar.setIconOnly(true);
    dataBar.setLeftToRight(false);
    dataBar.setWidthMin(10);
    dataBar.setWidthMax(90);
    dataBar.getMinThreshold().setRangeType(ConditionalFormattingThreshold.RangeType.MIN);
    dataBar.getMaxThreshold().setRangeType(ConditionalFormattingThreshold.RangeType.MAX);
    formatting.addConditionalFormatting(
        new CellRangeAddress[] {CellRangeAddress.valueOf("M2:M5")}, dataBarRule);

    XSSFConditionalFormattingRule iconSetRule =
        formatting.createConditionalFormattingRule(
            IconMultiStateFormatting.IconSet.GYR_3_TRAFFIC_LIGHTS);
    var iconSet = iconSetRule.getMultiStateFormatting();
    iconSet.setIconOnly(true);
    iconSet.setReversed(true);
    iconSet.getThresholds()[0].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENT);
    iconSet.getThresholds()[0].setValue(0d);
    iconSet.getThresholds()[1].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENT);
    iconSet.getThresholds()[1].setValue(33d);
    iconSet.getThresholds()[2].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENT);
    iconSet.getThresholds()[2].setValue(67d);
    formatting.addConditionalFormatting(
        new CellRangeAddress[] {CellRangeAddress.valueOf("N2:N5")}, iconSetRule);

    formatting.addConditionalFormatting(
        new CellRangeAddress[] {CellRangeAddress.valueOf("K2:K5")},
        formatting.createConditionalFormattingRule("K2>0"));
    sheet
        .getCTWorksheet()
        .getConditionalFormattingArray(
            sheet.getCTWorksheet().sizeOfConditionalFormattingArray() - 1)
        .getCfRuleArray(0)
        .setType(STCfType.TOP_10);
  }

  private static void createAdvancedTable(XSSFSheet sheet) {
    for (int rowIndex = 0; rowIndex <= 4; rowIndex++) {
      Row row = sheet.getRow(rowIndex);
      row.getCell(7)
          .setCellValue(
              switch (rowIndex) {
                case 0 -> "Region";
                case 1, 3 -> "North";
                case 2, 4 -> "South";
                default -> "Region";
              });
      if (rowIndex == 0) {
        row.getCell(8).setCellValue("Sales");
      } else {
        row.getCell(8)
            .setCellValue(
                switch (rowIndex) {
                  case 1 -> 10d;
                  case 2 -> 18d;
                  case 3 -> 12d;
                  case 4 -> 16d;
                  default -> 0d;
                });
      }
      row.getCell(9)
          .setCellValue(
              switch (rowIndex) {
                case 0 -> "Owner";
                case 1, 4 -> "Ada";
                case 2 -> "Grace";
                case 3 -> "Linus";
                default -> "Owner";
              });
    }

    XSSFTable table = sheet.createTable(new AreaReference("H1:J5", SpreadsheetVersion.EXCEL2007));
    table.setName("AdvancedTable");
    table.setDisplayName("AdvancedTable");
    table.getCTTable().setComment("Parity advanced table");
    table.getCTTable().setPublished(true);
    table.getCTTable().setTotalsRowCount(1);
    table.getCTTable().setTotalsRowShown(true);
    table.getCTTable().setHeaderRowCellStyle("ParityHeader");
    table.getCTTable().setDataCellStyle("ParityData");
    table.getCTTable().setTotalsRowCellStyle("ParityTotals");
    table.getCTTable().setInsertRow(true);
    table.getCTTable().setInsertRowShift(true);
    table.getCTTable().getTableColumns().getTableColumnArray(0).setTotalsRowLabel("Total");
    table
        .getCTTable()
        .getTableColumns()
        .getTableColumnArray(1)
        .setTotalsRowFunction(STTotalsRowFunction.SUM);
    table
        .getCTTable()
        .getTableColumns()
        .getTableColumnArray(1)
        .addNewCalculatedColumnFormula()
        .setStringValue("[@Sales]*2");
    table.getCTTable().getTableColumns().getTableColumnArray(2).setUniqueName("owner-unique");
  }

  private static SelfSignedMaterial createSelfSignedCertificate() {
    return XlsxParitySupport.call(
        "create self-signed parity certificate",
        () -> {
          ensureBouncyCastleProvider();
          KeyPairGenerator keyPairGenerator = KeyPairGenerator.getInstance("RSA");
          keyPairGenerator.initialize(2048, new SecureRandom(new byte[] {1, 2, 3, 4, 5, 6, 7, 8}));
          KeyPair keyPair = keyPairGenerator.generateKeyPair();
          Instant notBefore = Instant.parse("2026-01-01T00:00:00Z");
          Instant notAfter = Instant.parse("2030-01-01T00:00:00Z");
          X500Name subject = new X500Name(SIGNING_SUBJECT);
          JcaX509v3CertificateBuilder builder =
              new JcaX509v3CertificateBuilder(
                  subject,
                  BigInteger.valueOf(17L),
                  java.util.Date.from(notBefore),
                  java.util.Date.from(notAfter),
                  subject,
                  keyPair.getPublic());
          ContentSigner signer =
              new JcaContentSignerBuilder("SHA256withRSA")
                  .setProvider(BouncyCastleProvider.PROVIDER_NAME)
                  .build(keyPair.getPrivate());
          X509CertificateHolder holder = builder.build(signer);
          X509Certificate certificate =
              new JcaX509CertificateConverter()
                  .setProvider(BouncyCastleProvider.PROVIDER_NAME)
                  .getCertificate(holder);
          certificate.verify(keyPair.getPublic());
          return new SelfSignedMaterial(keyPair, certificate);
        });
  }

  private static void ensureBouncyCastleProvider() {
    if (Security.getProvider(BouncyCastleProvider.PROVIDER_NAME) == null) {
      Security.addProvider(new BouncyCastleProvider());
    }
  }

  private static GridGrindRequest coreWorkbookRequest(Path workbookPath) {
    return roundTripRequest(
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.SaveAs(
                workbookPath.toAbsolutePath().normalize().toString()),
            List.of(
                new WorkbookOperation.EnsureSheet("Ops"),
                new WorkbookOperation.EnsureSheet("Queue"),
                new WorkbookOperation.SetActiveSheet("Ops"),
                new WorkbookOperation.SetSelectedSheets(List.of("Ops")),
                new WorkbookOperation.SetSheetVisibility("Queue", ExcelSheetVisibility.HIDDEN),
                new WorkbookOperation.SetSheetProtection(
                    "Ops",
                    new SheetProtectionSettings(
                        true, false, false, false, false, false, false, false, false, true, false,
                        false, true, true, true)),
                new WorkbookOperation.MergeCells("Ops", "A1:B1"),
                new WorkbookOperation.SetRange(
                    "Ops",
                    "A1:D4",
                    List.of(
                        List.of(
                            new CellInput.Text("Quarterly Ops"),
                            new CellInput.Blank(),
                            new CellInput.Text("Launch"),
                            new CellInput.Text("Link")),
                        List.of(
                            new CellInput.Text("Owner"),
                            new CellInput.Text("Hours"),
                            new CellInput.Text("Ready"),
                            new CellInput.Text("Score")),
                        List.of(
                            new CellInput.RichText(
                                List.of(
                                    new RichTextRunInput(
                                        "Ada",
                                        new CellFontInput(
                                            true, null, null, null, "#204060", null, null)),
                                    new RichTextRunInput(
                                        " Lovelace",
                                        new CellFontInput(
                                            null, true, null, null, "#4080A0", null, null)))),
                            new CellInput.Numeric(12.5d),
                            new CellInput.BooleanValue(true),
                            new CellInput.Formula("B3*2")),
                        List.of(
                            new CellInput.Date(LocalDate.of(2026, 4, 1)),
                            new CellInput.DateTime(LocalDateTime.of(2026, 4, 1, 9, 30, 0)),
                            new CellInput.Text("Docs"),
                            new CellInput.Numeric(99.0d)))),
                new WorkbookOperation.ApplyStyle(
                    "Ops",
                    "A1:D4",
                    new CellStyleInput(
                        null,
                        new CellAlignmentInput(
                            true,
                            ExcelHorizontalAlignment.CENTER,
                            ExcelVerticalAlignment.CENTER,
                            null,
                            null),
                        new CellFontInput(
                            true,
                            null,
                            "Aptos",
                            new FontHeightInput.Points(new BigDecimal("11.5")),
                            "#112233",
                            null,
                            null),
                        new CellFillInput(ExcelFillPattern.SOLID, "#D9EAF7", null),
                        new CellBorderInput(
                            new CellBorderSideInput(ExcelBorderStyle.THIN, "#5B7C99"),
                            null,
                            null,
                            null,
                            null),
                        new CellProtectionInput(true, null))),
                new WorkbookOperation.SetColumnWidth("Ops", 0, 1, 18.0d),
                new WorkbookOperation.SetRowHeight("Ops", 0, 3, 21.0d),
                new WorkbookOperation.GroupRows("Ops", new RowSpanInput.Band(1, 3), true),
                new WorkbookOperation.GroupColumns("Ops", new ColumnSpanInput.Band(0, 1), false),
                new WorkbookOperation.SetSheetPane("Ops", new PaneInput.Frozen(1, 1, 1, 1)),
                new WorkbookOperation.SetSheetZoom("Ops", 125),
                new WorkbookOperation.SetPrintLayout(
                    "Ops",
                    new PrintLayoutInput(
                        new PrintAreaInput.Range("A1:D20"),
                        ExcelPrintOrientation.LANDSCAPE,
                        new PrintScalingInput.Fit(1, 2),
                        new PrintTitleRowsInput.Band(0, 1),
                        new PrintTitleColumnsInput.Band(0, 0),
                        new HeaderFooterTextInput("Ops", "Quarterly", "2026"),
                        new HeaderFooterTextInput("GridGrind", "Parity", "1"))),
                new WorkbookOperation.SetHyperlink(
                    "Ops", "C4", new HyperlinkTarget.Url("https://example.com/docs")),
                new WorkbookOperation.SetComment(
                    "Ops", "C4", new CommentInput("Review docs", "GridGrind", true)),
                new WorkbookOperation.SetNamedRange(
                    "OpsScore",
                    new NamedRangeScope.Workbook(),
                    new NamedRangeTarget("Ops", "D3:D4")),
                new WorkbookOperation.SetNamedRange(
                    "QueueWindow",
                    new NamedRangeScope.Sheet("Queue"),
                    new NamedRangeTarget("Queue", "A1:B3")),
                new WorkbookOperation.SetDataValidation(
                    "Ops",
                    "B3:B10",
                    new DataValidationInput(
                        new DataValidationRuleInput.WholeNumber(
                            ExcelComparisonOperator.BETWEEN, "1", "40"),
                        false,
                        false,
                        new DataValidationPromptInput("Hours", "Enter 1-40", true),
                        new DataValidationErrorAlertInput(
                            ExcelDataValidationErrorStyle.STOP,
                            "Invalid",
                            "Hours must be 1-40",
                            true))),
                new WorkbookOperation.SetConditionalFormatting(
                    "Ops",
                    new ConditionalFormattingBlockInput(
                        List.of("D3:D4"),
                        List.of(
                            new ConditionalFormattingRuleInput.FormulaRule(
                                "D3>20",
                                false,
                                new DifferentialStyleInput(
                                    null, true, null, null, "#006100", null, null, "#C6EFCE",
                                    null)),
                            new ConditionalFormattingRuleInput.CellValueRule(
                                ExcelComparisonOperator.LESS_THAN,
                                "20",
                                null,
                                false,
                                new DifferentialStyleInput(
                                    null, null, null, null, "#9C0006", null, null, "#FFC7CE",
                                    null))))),
                new WorkbookOperation.SetAutofilter("Ops", "A2:D4"),
                new WorkbookOperation.SetRange(
                    "Queue",
                    "A1:B3",
                    List.of(
                        List.of(new CellInput.Text("Owner"), new CellInput.Text("Task")),
                        List.of(new CellInput.Text("Ada"), new CellInput.Text("Docs")),
                        List.of(new CellInput.Text("Grace"), new CellInput.Text("Review")))),
                new WorkbookOperation.SetTable(
                    new TableInput(
                        "QueueTable",
                        "Queue",
                        "A1:B3",
                        false,
                        new TableStyleInput.Named("TableStyleMedium2", false, false, true, false))),
                new WorkbookOperation.EvaluateFormulas(),
                new WorkbookOperation.ForceFormulaRecalculationOnOpen()),
            List.of()));
  }

  private static GridGrindRequest drawingAuthoringRequest(Path workbookPath) {
    return new GridGrindRequest(
        new GridGrindRequest.WorkbookSource.New(),
        new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
        null,
        List.of(
            new WorkbookOperation.EnsureSheet("Ops"),
            new WorkbookOperation.SetPicture(
                "Ops",
                new PictureInput(
                    "OpsPicture",
                    pictureDataInput(),
                    twoCellAnchorInput(1, 1, 4, 6, ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                    "Queue preview")),
            new WorkbookOperation.SetShape(
                "Ops",
                new ShapeInput(
                    "OpsShape",
                    ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                    twoCellAnchorInput(5, 1, 8, 5, ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                    "rect",
                    "Queue")),
            new WorkbookOperation.SetShape(
                "Ops",
                new ShapeInput(
                    "OpsConnector",
                    ExcelAuthoredDrawingShapeKind.CONNECTOR,
                    twoCellAnchorInput(9, 1, 11, 4, ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                    null,
                    null)),
            new WorkbookOperation.SetEmbeddedObject(
                "Ops",
                new EmbeddedObjectInput(
                    "OpsEmbed",
                    "Payload",
                    "payload.txt",
                    "payload.txt",
                    Base64.getEncoder()
                        .encodeToString(
                            "GridGrind embedded payload".getBytes(StandardCharsets.UTF_8)),
                    pictureDataInput(),
                    twoCellAnchorInput(12, 1, 15, 6, ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE))),
            new WorkbookOperation.SetDrawingObjectAnchor(
                "Ops",
                "OpsPicture",
                twoCellAnchorInput(6, 2, 9, 7, ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE)),
            new WorkbookOperation.DeleteDrawingObject("Ops", "OpsConnector")),
        List.of());
  }

  private static GridGrindRequest chartAuthoringRequest(Path workbookPath) {
    return new GridGrindRequest(
        new GridGrindRequest.WorkbookSource.New(),
        new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
        null,
        List.of(
            new WorkbookOperation.EnsureSheet("Chart"),
            new WorkbookOperation.SetCell("Chart", "A1", new CellInput.Text("Month")),
            new WorkbookOperation.SetCell("Chart", "B1", new CellInput.Text("Plan")),
            new WorkbookOperation.SetCell("Chart", "C1", new CellInput.Text("Actual")),
            new WorkbookOperation.SetCell("Chart", "A2", new CellInput.Text("Jan")),
            new WorkbookOperation.SetCell("Chart", "B2", new CellInput.Numeric(10d)),
            new WorkbookOperation.SetCell("Chart", "C2", new CellInput.Numeric(12d)),
            new WorkbookOperation.SetCell("Chart", "A3", new CellInput.Text("Feb")),
            new WorkbookOperation.SetCell("Chart", "B3", new CellInput.Numeric(18d)),
            new WorkbookOperation.SetCell("Chart", "C3", new CellInput.Numeric(16d)),
            new WorkbookOperation.SetCell("Chart", "A4", new CellInput.Text("Mar")),
            new WorkbookOperation.SetCell("Chart", "B4", new CellInput.Numeric(15d)),
            new WorkbookOperation.SetCell("Chart", "C4", new CellInput.Numeric(21d)),
            new WorkbookOperation.SetNamedRange(
                "ChartCategories",
                new NamedRangeScope.Workbook(),
                new NamedRangeTarget("Chart", "A2:A4")),
            new WorkbookOperation.SetNamedRange(
                "ChartActual",
                new NamedRangeScope.Workbook(),
                new NamedRangeTarget("Chart", "C2:C4")),
            new WorkbookOperation.SetChart(
                "Chart",
                new ChartInput.Bar(
                    "OpsChart",
                    twoCellAnchorInput(4, 1, 11, 16, ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                    new ChartInput.Title.Text("Roadmap"),
                    new ChartInput.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
                    ExcelChartDisplayBlanksAs.SPAN,
                    false,
                    true,
                    ExcelChartBarDirection.COLUMN,
                    List.of(
                        new ChartInput.Series(
                            new ChartInput.Title.Formula("B1"),
                            new ChartInput.DataSource("A2:A4"),
                            new ChartInput.DataSource("B2:B4")),
                        new ChartInput.Series(
                            new ChartInput.Title.Formula("C1"),
                            new ChartInput.DataSource("ChartCategories"),
                            new ChartInput.DataSource("ChartActual")))))),
        List.of());
  }

  private static GridGrindRequest pivotAuthoringRequest(Path workbookPath) {
    return roundTripRequest(
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
            null,
            List.of(
                new WorkbookOperation.EnsureSheet("Data"),
                new WorkbookOperation.EnsureSheet("RangeReport"),
                new WorkbookOperation.EnsureSheet("NamedReport"),
                new WorkbookOperation.EnsureSheet("TableReport"),
                new WorkbookOperation.SetCell("Data", "A1", new CellInput.Text("Region")),
                new WorkbookOperation.SetCell("Data", "B1", new CellInput.Text("Stage")),
                new WorkbookOperation.SetCell("Data", "C1", new CellInput.Text("Owner")),
                new WorkbookOperation.SetCell("Data", "D1", new CellInput.Text("Amount")),
                new WorkbookOperation.SetCell("Data", "A2", new CellInput.Text("North")),
                new WorkbookOperation.SetCell("Data", "B2", new CellInput.Text("Plan")),
                new WorkbookOperation.SetCell("Data", "C2", new CellInput.Text("Ada")),
                new WorkbookOperation.SetCell("Data", "D2", new CellInput.Numeric(10d)),
                new WorkbookOperation.SetCell("Data", "A3", new CellInput.Text("North")),
                new WorkbookOperation.SetCell("Data", "B3", new CellInput.Text("Do")),
                new WorkbookOperation.SetCell("Data", "C3", new CellInput.Text("Ada")),
                new WorkbookOperation.SetCell("Data", "D3", new CellInput.Numeric(15d)),
                new WorkbookOperation.SetCell("Data", "A4", new CellInput.Text("South")),
                new WorkbookOperation.SetCell("Data", "B4", new CellInput.Text("Plan")),
                new WorkbookOperation.SetCell("Data", "C4", new CellInput.Text("Lin")),
                new WorkbookOperation.SetCell("Data", "D4", new CellInput.Numeric(7d)),
                new WorkbookOperation.SetCell("Data", "A5", new CellInput.Text("South")),
                new WorkbookOperation.SetCell("Data", "B5", new CellInput.Text("Do")),
                new WorkbookOperation.SetCell("Data", "C5", new CellInput.Text("Lin")),
                new WorkbookOperation.SetCell("Data", "D5", new CellInput.Numeric(12d)),
                new WorkbookOperation.SetNamedRange(
                    "PivotSource",
                    new NamedRangeScope.Workbook(),
                    new NamedRangeTarget("Data", "A1:D5")),
                new WorkbookOperation.SetTable(
                    new TableInput(
                        "SalesTable", "Data", "A1:D5", false, new TableStyleInput.None())),
                new WorkbookOperation.SetPivotTable(
                    new PivotTableInput(
                        "Sales Pivot",
                        "RangeReport",
                        new PivotTableInput.Source.Range("Data", "A1:D5"),
                        new PivotTableInput.Anchor("C5"),
                        List.of("Region"),
                        List.of("Stage"),
                        List.of(),
                        List.of(defaultPivotDataFieldInput()))),
                new WorkbookOperation.SetPivotTable(
                    new PivotTableInput(
                        "Named Pivot",
                        "NamedReport",
                        new PivotTableInput.Source.NamedRange("PivotSource"),
                        new PivotTableInput.Anchor("A3"),
                        List.of("Region"),
                        List.of(),
                        List.of("Owner"),
                        List.of(defaultPivotDataFieldInput()))),
                new WorkbookOperation.SetPivotTable(
                    new PivotTableInput(
                        "Table Pivot",
                        "TableReport",
                        new PivotTableInput.Source.Table("SalesTable"),
                        new PivotTableInput.Anchor("F4"),
                        List.of("Stage"),
                        List.of(),
                        List.of(),
                        List.of(defaultPivotDataFieldInput())))),
            List.of()));
  }

  private static PivotTableInput.DataField defaultPivotDataFieldInput() {
    return new PivotTableInput.DataField(
        "Amount",
        dev.erst.gridgrind.excel.ExcelPivotDataConsolidateFunction.SUM,
        "Total Amount",
        "#,##0.00");
  }

  private static PictureDataInput pictureDataInput() {
    return new PictureDataInput(
        ExcelPictureFormat.PNG, Base64.getEncoder().encodeToString(PNG_PIXEL_BYTES));
  }

  private static DrawingAnchorInput.TwoCell twoCellAnchorInput(
      int fromColumn, int fromRow, int toColumn, int toRow, ExcelDrawingAnchorBehavior behavior) {
    return new DrawingAnchorInput.TwoCell(
        new DrawingMarkerInput(fromColumn, fromRow, 0, 0),
        new DrawingMarkerInput(toColumn, toRow, 0, 0),
        behavior);
  }

  private static GridGrindRequest roundTripRequest(GridGrindRequest request) {
    return XlsxParitySupport.call(
        "round-trip parity request through JSON",
        () -> GridGrindJson.readRequest(GridGrindJson.writeRequestBytes(request)));
  }

  private static void seedMatrix(XSSFSheet sheet, int rows, int columns) {
    for (int rowIndex = 0; rowIndex < rows; rowIndex++) {
      Row row = sheet.createRow(rowIndex);
      for (int columnIndex = 0; columnIndex < columns; columnIndex++) {
        row.createCell(columnIndex).setCellValue("R" + rowIndex + "C" + columnIndex);
      }
    }
  }

  private static void seedChartData(XSSFSheet sheet) {
    sheet.createRow(0).createCell(0).setCellValue("Month");
    sheet.getRow(0).createCell(1).setCellValue("Plan");
    sheet.getRow(0).createCell(2).setCellValue("Actual");
    sheet.createRow(1).createCell(0).setCellValue("Jan");
    sheet.getRow(1).createCell(1).setCellValue(10d);
    sheet.getRow(1).createCell(2).setCellValue(12d);
    sheet.createRow(2).createCell(0).setCellValue("Feb");
    sheet.getRow(2).createCell(1).setCellValue(18d);
    sheet.getRow(2).createCell(2).setCellValue(16d);
    sheet.createRow(3).createCell(0).setCellValue("Mar");
    sheet.getRow(3).createCell(1).setCellValue(15d);
    sheet.getRow(3).createCell(2).setCellValue(21d);
  }

  private static final byte[] PNG_PIXEL_BYTES =
      Base64.getDecoder()
          .decode(
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=");

  private record SelfSignedMaterial(KeyPair keyPair, X509Certificate certificate) {}

  /** Helper for authoring formula-defined names in direct-POI parity scenarios. */
  private static final class NameSupport {
    private NameSupport() {}

    private static void addFormulaDefinedName(
        XSSFWorkbook workbook, String name, Integer sheetIndex, String formula) {
      var definedName = workbook.createName();
      definedName.setNameName(name);
      if (sheetIndex != null) {
        definedName.setSheetIndex(sheetIndex);
      }
      definedName.setRefersToFormula(formula);
    }
  }
}
