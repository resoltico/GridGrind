package dev.erst.gridgrind.protocol.parity;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HexFormat;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.function.Function;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.dsig.SignatureConfig;
import org.apache.poi.poifs.crypt.dsig.SignatureInfo;
import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.poifs.filesystem.Ole10Native;
import org.apache.poi.poifs.filesystem.Ole10NativeException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.formula.OperationEvaluationContext;
import org.apache.poi.ss.formula.eval.EvaluationException;
import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.ss.formula.eval.OperandResolver;
import org.apache.poi.ss.formula.eval.ValueEval;
import org.apache.poi.ss.formula.functions.FreeRefFunction;
import org.apache.poi.ss.formula.udf.DefaultUDFFinder;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFConnector;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFGraphicFrame;
import org.apache.poi.xssf.usermodel.XSSFObjectData;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFShapeGroup;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTConditionalFormatting;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidations;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbookProtection;

/** Direct Apache POI oracle for parity tests. */
final class XlsxParityOracle {
  private XlsxParityOracle() {}

  static CoreWorkbookSnapshot coreWorkbook(Path workbookPath) {
    return withWorkbook(
        workbookPath,
        workbook -> {
          XSSFSheet ops = workbook.getSheet("Ops");
          XSSFSheet queue = workbook.getSheet("Queue");
          int queueIndex = workbook.getSheetIndex(queue);
          XSSFCell formulaCell = cell(ops, "D3");
          return new CoreWorkbookSnapshot(
              sheetNames(workbook),
              workbook.getSheetName(workbook.getActiveSheetIndex()),
              selectedSheetNames(workbook),
              workbook.getForceFormulaRecalculation(),
              workbook.getSheetVisibility(queueIndex) != SheetVisibility.VISIBLE,
              ops.getProtect(),
              ops.getMergedRegion(0).formatAsString(),
              cell(ops, "C4").getHyperlink().getAddress(),
              cell(ops, "C4").getCellComment().getString().getString(),
              cell(ops, "C4").getCellComment().getAuthor(),
              cell(ops, "C4").getCellComment().isVisible(),
              (int) ops.getCTWorksheet().getSheetViews().getSheetViewArray(0).getZoomScale(),
              ops.getPaneInformation() == null
                  ? "NONE"
                  : "%s:%d:%d"
                      .formatted(
                          ops.getPaneInformation().isFreezePane() ? "FROZEN" : "SPLIT",
                          ops.getPaneInformation().getHorizontalSplitPosition(),
                          ops.getPaneInformation().getVerticalSplitPosition()),
              ops.getColumnWidthInPixels(0) > 0.0d,
              ops.getRow(0).getHeightInPoints(),
              formulaCell.getCellFormula(),
              workbook
                  .getCreationHelper()
                  .createFormulaEvaluator()
                  .evaluate(formulaCell)
                  .getNumberValue(),
              ops.getCTWorksheet().getAutoFilter().getRef(),
              queue.getTables().getFirst().getName(),
              queue.getTables().getFirst().getColumns().stream()
                  .map(column -> column.getName())
                  .toList(),
              ops.getSheetConditionalFormatting().getNumConditionalFormattings(),
              dataValidationKinds(ops));
        });
  }

  static WorkbookProtectionSnapshot workbookProtection(Path workbookPath) {
    return withWorkbook(
        workbookPath,
        workbook -> {
          var protection = workbook.getCTWorkbook().getWorkbookProtection();
          return new WorkbookProtectionSnapshot(
              workbook.isStructureLocked(),
              workbook.isWindowsLocked(),
              workbook.isRevisionLocked(),
              workbookPasswordHashPresent(protection),
              revisionsPasswordHashPresent(protection),
              protection != null
                  && workbook.validateWorkbookPassword(
                      XlsxParityScenarios.WORKBOOK_PROTECTION_PASSWORD));
        });
  }

  private static boolean workbookPasswordHashPresent(CTWorkbookProtection protection) {
    return protection != null
        && (protection.isSetWorkbookPassword()
            || protection.isSetWorkbookHashValue()
            || protection.isSetWorkbookSaltValue()
            || protection.isSetWorkbookSpinCount()
            || protection.isSetWorkbookAlgorithmName());
  }

  private static boolean revisionsPasswordHashPresent(CTWorkbookProtection protection) {
    return protection != null
        && (protection.isSetRevisionsPassword()
            || protection.isSetRevisionsHashValue()
            || protection.isSetRevisionsSaltValue()
            || protection.isSetRevisionsSpinCount()
            || protection.isSetRevisionsAlgorithmName());
  }

  static AdvancedPrintSnapshot advancedPrint(Path workbookPath) {
    return advancedPrint(workbookPath, "Advanced");
  }

  static AdvancedPrintSnapshot advancedPrint(Path workbookPath, String sheetName) {
    return withWorkbook(
        workbookPath,
        workbook -> {
          XSSFSheet sheet = workbook.getSheet(sheetName);
          return new AdvancedPrintSnapshot(
              sheet.getMargin(XSSFSheet.LeftMargin),
              sheet.getMargin(XSSFSheet.RightMargin),
              sheet.getMargin(XSSFSheet.TopMargin),
              sheet.getMargin(XSSFSheet.BottomMargin),
              sheet.getHorizontallyCenter(),
              sheet.getVerticallyCenter(),
              sheet.getPrintSetup().getPaperSize(),
              sheet.getPrintSetup().getDraft(),
              sheet.getPrintSetup().getNoColor(),
              sheet.getPrintSetup().getCopies(),
              sheet.getPrintSetup().getUsePage(),
              sheet.getPrintSetup().getPageStart(),
              java.util.Arrays.stream(sheet.getRowBreaks()).boxed().toList(),
              java.util.Arrays.stream(sheet.getColumnBreaks()).boxed().toList());
        });
  }

  static CommentSnapshot comment(Path workbookPath, String sheetName, String address) {
    return withWorkbook(
        workbookPath,
        workbook -> {
          XSSFComment comment = cell(workbook.getSheet(sheetName), address).getCellComment();
          XSSFClientAnchor anchor = (XSSFClientAnchor) comment.getClientAnchor();
          return new CommentSnapshot(
              comment.getAuthor(),
              comment.isVisible(),
              richTextRunCount(comment.getString()),
              anchor.getCol1(),
              anchor.getRow1(),
              anchor.getCol2(),
              anchor.getRow2());
        });
  }

  static StyleSnapshot style(Path workbookPath, String sheetName, String address) {
    return withWorkbook(
        workbookPath,
        workbook -> {
          XSSFCellStyle style = cell(workbook.getSheet(sheetName), address).getCellStyle();
          StylesFillSnapshot fillSnapshot = fillSnapshot(workbook, style);
          XSSFColor fontColor = style.getFont().getXSSFColor();
          XSSFColor fillColor = style.getFillForegroundColorColor();
          XSSFColor borderColor = style.getBottomBorderXSSFColor();
          return new StyleSnapshot(
              colorDescriptor(fontColor),
              colorDescriptor(fillColor),
              colorDescriptor(borderColor),
              fillSnapshot.gradientFill());
        });
  }

  static List<NamedRangeSnapshot> namedRanges(Path workbookPath) {
    return withWorkbook(
        workbookPath,
        workbook -> {
          List<NamedRangeSnapshot> snapshots = new ArrayList<>();
          for (Name name : workbook.getAllNames()) {
            snapshots.add(
                new NamedRangeSnapshot(
                    name.getNameName(),
                    name.getSheetIndex() < 0
                        ? "WORKBOOK"
                        : workbook.getSheetName(name.getSheetIndex()),
                    name.getRefersToFormula(),
                    isFormulaDefined(name)));
          }
          snapshots.sort(
              Comparator.comparing(NamedRangeSnapshot::name)
                  .thenComparing(NamedRangeSnapshot::scope)
                  .thenComparing(NamedRangeSnapshot::refersToFormula));
          return List.copyOf(snapshots);
        });
  }

  static List<String> dataValidationKinds(Path workbookPath, String sheetName) {
    return withWorkbook(
        workbookPath, workbook -> dataValidationKinds(workbook.getSheet(sheetName)));
  }

  static AutofilterSnapshot autofilter(Path workbookPath, String sheetName) {
    return withWorkbook(
        workbookPath,
        workbook -> {
          XSSFSheet sheet = workbook.getSheet(sheetName);
          if (!sheet.getCTWorksheet().isSetAutoFilter()) {
            return new AutofilterSnapshot("", List.of(), null);
          }
          var autoFilter = sheet.getCTWorksheet().getAutoFilter();
          return new AutofilterSnapshot(
              autoFilter.getRef(),
              java.util.Arrays.stream(autoFilter.getFilterColumnArray())
                  .map(
                      filterColumn ->
                          new FilterColumnSnapshot(
                              filterColumn.getColId(),
                              filterColumn.isSetFilters()
                                  ? java.util.Arrays.stream(
                                          filterColumn.getFilters().getFilterArray())
                                      .map(
                                          filter -> Objects.requireNonNullElse(filter.getVal(), ""))
                                      .toList()
                                  : List.of(),
                              filterColumn.isSetFilters() && filterColumn.getFilters().getBlank()))
                  .toList(),
              autoFilter.isSetSortState()
                  ? new SortStateSnapshot(
                      autoFilter.getSortState().getRef(),
                      autoFilter.getSortState().isSetCaseSensitive()
                          && autoFilter.getSortState().getCaseSensitive(),
                      autoFilter.getSortState().isSetColumnSort()
                          && autoFilter.getSortState().getColumnSort(),
                      autoFilter.getSortState().isSetSortMethod()
                          ? autoFilter.getSortState().getSortMethod().toString()
                          : "",
                      java.util.Arrays.stream(autoFilter.getSortState().getSortConditionArray())
                          .map(
                              condition ->
                                  new SortConditionSnapshot(
                                      condition.getRef(),
                                      condition.isSetDescending() && condition.getDescending(),
                                      condition.isSetSortBy()
                                          ? condition.getSortBy().toString()
                                          : ""))
                          .toList())
                  : null);
        });
  }

  static TableSnapshot table(Path workbookPath, String tableName) {
    return withWorkbook(
        workbookPath,
        workbook -> {
          for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            for (XSSFTable table : workbook.getSheetAt(sheetIndex).getTables()) {
              if (table.getName().equals(tableName)) {
                var ctTable = table.getCTTable();
                var amountColumn = ctTable.getTableColumns().getTableColumnArray(1);
                var ownerColumn = ctTable.getTableColumns().getTableColumnArray(2);
                return new TableSnapshot(
                    ctTable.getComment(),
                    ctTable.getPublished(),
                    ctTable.getInsertRow(),
                    ctTable.getInsertRowShift(),
                    ctTable.getHeaderRowCellStyle(),
                    ctTable.getDataCellStyle(),
                    ctTable.getTotalsRowCellStyle(),
                    ctTable.getTableColumns().getTableColumnArray(0).getTotalsRowLabel(),
                    amountColumn.getTotalsRowFunction().toString(),
                    amountColumn.getCalculatedColumnFormula().getStringValue(),
                    ownerColumn.getUniqueName());
              }
            }
          }
          throw new IllegalArgumentException("Unknown table " + tableName);
        });
  }

  static List<String> conditionalFormattingKinds(Path workbookPath, String sheetName) {
    return withWorkbook(
        workbookPath,
        workbook -> {
          XSSFSheet sheet = workbook.getSheet(sheetName);
          List<String> kinds = new ArrayList<>();
          for (CTConditionalFormatting block :
              sheet.getCTWorksheet().getConditionalFormattingArray()) {
            for (var rule : block.getCfRuleArray()) {
              kinds.add(rule.getType().toString());
            }
          }
          kinds.sort(Comparator.naturalOrder());
          return List.copyOf(kinds);
        });
  }

  static double evaluateFormula(Path workbookPath, String sheetName, String address) {
    return withWorkbook(
        workbookPath,
        workbook ->
            workbook
                .getCreationHelper()
                .createFormulaEvaluator()
                .evaluate(cell(workbook.getSheet(sheetName), address))
                .getNumberValue());
  }

  static String cachedFormulaRawValue(Path workbookPath, String sheetName, String address) {
    return withWorkbook(
        workbookPath,
        workbook -> {
          XSSFCell cell = cell(workbook.getSheet(sheetName), address);
          return cell.getCTCell().isSetV() ? cell.getCTCell().getV() : null;
        });
  }

  static double evaluateExternalFormula(Path workbookPath, Path referencedWorkbookPath) {
    return XlsxParitySupport.call(
        "evaluate external-formula parity workbook",
        () -> {
          try (InputStream referencedStream = Files.newInputStream(referencedWorkbookPath);
              Workbook referencedWorkbook = WorkbookFactory.create(referencedStream);
              InputStream workbookStream = Files.newInputStream(workbookPath);
              XSSFWorkbook workbook = new XSSFWorkbook(workbookStream)) {
            workbook.linkExternalWorkbook(
                referencedWorkbookPath.getFileName().toString(), referencedWorkbook);
            var evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.setupReferencedWorkbooks(
                java.util.Map.of(
                    referencedWorkbookPath.getFileName().toString(),
                    referencedWorkbook.getCreationHelper().createFormulaEvaluator()));
            return evaluator
                .evaluate(workbook.getSheet("Ops").getRow(0).getCell(1))
                .getNumberValue();
          }
        });
  }

  static boolean externalFormulaFailsWithoutBinding(Path workbookPath) {
    return XlsxParitySupport.call(
        "evaluate external-formula parity workbook without bindings",
        () -> {
          try (InputStream workbookStream = Files.newInputStream(workbookPath);
              XSSFWorkbook workbook = new XSSFWorkbook(workbookStream)) {
            var evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.setIgnoreMissingWorkbooks(false);
            evaluator.evaluate(workbook.getSheet("Ops").getRow(0).getCell(1));
            return false;
          } catch (RuntimeException exception) {
            return true;
          }
        });
  }

  static double evaluateExternalFormulaUsingCachedValue(Path workbookPath) {
    return XlsxParitySupport.call(
        "evaluate external-formula parity workbook with cached-value fallback",
        () -> {
          try (InputStream workbookStream = Files.newInputStream(workbookPath);
              XSSFWorkbook workbook = new XSSFWorkbook(workbookStream)) {
            var evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.setIgnoreMissingWorkbooks(true);
            return evaluator
                .evaluate(workbook.getSheet("Ops").getRow(0).getCell(1))
                .getNumberValue();
          }
        });
  }

  static double evaluateUdfFormula(Path workbookPath) {
    return XlsxParitySupport.call(
        "evaluate UDF parity workbook",
        () -> {
          try (InputStream workbookStream = Files.newInputStream(workbookPath);
              XSSFWorkbook workbook = new XSSFWorkbook(workbookStream)) {
            workbook.addToolPack(
                new DefaultUDFFinder(
                    new String[] {"DOUBLE"},
                    new FreeRefFunction[] {
                      (ValueEval[] args, OperationEvaluationContext ec) -> {
                        try {
                          ValueEval value =
                              OperandResolver.getSingleValue(
                                  args[0], ec.getRowIndex(), ec.getColumnIndex());
                          return new NumberEval(OperandResolver.coerceValueToDouble(value) * 2.0d);
                        } catch (EvaluationException exception) {
                          return exception.getErrorEval();
                        }
                      }
                    }));
            return workbook
                .getCreationHelper()
                .createFormulaEvaluator()
                .evaluate(workbook.getSheet("Ops").getRow(0).getCell(1))
                .getNumberValue();
          }
        });
  }

  static DrawingSnapshot drawing(Path workbookPath) {
    DrawingSheetSnapshot snapshot = drawingSheet(workbookPath, "Ops");
    PictureDrawingObjectSnapshot picture =
        snapshot.objects().stream()
            .filter(PictureDrawingObjectSnapshot.class::isInstance)
            .map(PictureDrawingObjectSnapshot.class::cast)
            .findFirst()
            .orElseThrow(() -> new IllegalStateException("Expected picture shape"));
    return new DrawingSnapshot(
        snapshot.objects().size(),
        picture.anchor().col1(),
        picture.anchor().row1(),
        picture.anchor().col2(),
        picture.anchor().row2(),
        picture.pictureDigest());
  }

  static EmbeddedObjectSnapshot embeddedObject(Path workbookPath) {
    DrawingSheetSnapshot snapshot = drawingSheet(workbookPath, "Objects");
    EmbeddedObjectDrawingObjectSnapshot objectData =
        snapshot.objects().stream()
            .filter(EmbeddedObjectDrawingObjectSnapshot.class::isInstance)
            .map(EmbeddedObjectDrawingObjectSnapshot.class::cast)
            .findFirst()
            .orElseThrow(() -> new IllegalStateException("Expected embedded object"));
    return new EmbeddedObjectSnapshot(
        snapshot.objects().size(), objectData.fileName(), objectData.objectDigest());
  }

  static DrawingSheetSnapshot drawingSheet(Path workbookPath, String sheetName) {
    return withWorkbook(
        workbookPath,
        workbook -> {
          XSSFSheet sheet = workbook.getSheet(sheetName);
          if (sheet == null) {
            throw new IllegalStateException("Expected sheet " + sheetName);
          }
          XSSFDrawing drawing = sheet.getDrawingPatriarch();
          List<DirectDrawingObjectSnapshot> objects = new ArrayList<>();
          if (drawing != null) {
            for (XSSFShape shape : drawing.getShapes()) {
              objects.add(drawingObjectSnapshot(drawing, shape));
            }
          }

          List<String> mergedRegions = new ArrayList<>();
          for (int index = 0; index < sheet.getNumMergedRegions(); index++) {
            mergedRegions.add(sheet.getMergedRegion(index).formatAsString());
          }

          List<CellCommentSnapshot> comments =
              sheet.getCellComments().entrySet().stream()
                  .sorted(Comparator.comparing(entry -> entry.getKey().formatAsString()))
                  .map(
                      entry ->
                          new CellCommentSnapshot(
                              entry.getKey().formatAsString(),
                              entry.getValue().getString().getString(),
                              entry.getValue().getAuthor(),
                              entry.getValue().isVisible()))
                  .toList();

          return new DrawingSheetSnapshot(
              List.copyOf(objects), List.copyOf(mergedRegions), comments);
        });
  }

  static ChartSnapshot chart(Path workbookPath) {
    return withWorkbook(
        workbookPath,
        workbook -> {
          XSSFDrawing drawing = workbook.getSheet("Chart").createDrawingPatriarch();
          XSSFChart chart = drawing.getCharts().getFirst();
          return new ChartSnapshot(
              drawing.getCharts().size(),
              chart.getTitleText() == null ? "" : chart.getTitleText().getString(),
              chart.getCTChart().getPlotArea().sizeOfBarChartArray());
        });
  }

  static PivotSnapshot pivot(Path workbookPath) {
    return withWorkbook(
        workbookPath,
        workbook -> {
          XSSFSheet pivotSheet = workbook.getSheet("Pivot");
          XSSFPivotTable pivotTable =
              pivotSheet.getRelations().stream()
                  .filter(XSSFPivotTable.class::isInstance)
                  .map(XSSFPivotTable.class::cast)
                  .findFirst()
                  .orElseThrow(() -> new IllegalStateException("Expected pivot table relation"));
          return new PivotSnapshot(
              1,
              pivotTable
                  .getPivotCacheDefinition()
                  .getCTPivotCacheDefinition()
                  .getCacheSource()
                  .getWorksheetSource()
                  .getRef(),
              pivotTable.getRowLabelColumns());
        });
  }

  static int eventModelRowCount(Path workbookPath) {
    return XlsxParitySupport.call(
        "count event-model rows in parity workbook",
        () -> {
          try (OPCPackage pkg = OPCPackage.open(workbookPath.toFile(), PackageAccess.READ)) {
            XSSFReader reader = new XSSFReader(pkg);
            Iterator<InputStream> sheets = reader.getSheetsData();
            if (!sheets.hasNext()) {
              throw new IllegalStateException("Expected at least one sheet");
            }
            String xml = new String(sheets.next().readAllBytes(), StandardCharsets.UTF_8);
            return countOccurrences(xml, "<row");
          }
        });
  }

  static Path sxssfWriteWorkbook(Path temporaryRoot) {
    return XlsxParitySupport.call(
        "write SXSSF parity workbook",
        () -> {
          Path workbookPath = temporaryRoot.resolve("sxssf-written.xlsx");
          try (SXSSFWorkbook workbook = new SXSSFWorkbook(50);
              var outputStream = Files.newOutputStream(workbookPath)) {
            var sheet = workbook.createSheet("Streamed");
            for (int rowIndex = 0; rowIndex < 1500; rowIndex++) {
              var row = sheet.createRow(rowIndex);
              row.createCell(0).setCellValue("R" + rowIndex);
              row.createCell(1).setCellValue(rowIndex);
            }
            workbook.write(outputStream);
          }
          return workbookPath;
        });
  }

  static boolean encryptedWorkbookOpens(Path workbookPath) {
    return XlsxParitySupport.call(
        "open encrypted parity workbook through POI",
        () -> {
          try (POIFSFileSystem fileSystem = new POIFSFileSystem(workbookPath.toFile());
              InputStream decryptedStream = decryptedWorkbookStream(fileSystem);
              Workbook workbook = WorkbookFactory.create(decryptedStream)) {
            return "Encrypted workbook"
                .equals(workbook.getSheet("Encrypted").getRow(0).getCell(0).getStringCellValue());
          }
        });
  }

  static boolean signatureValid(Path workbookPath) {
    return XlsxParitySupport.call(
        "verify parity workbook signature",
        () -> {
          try (OPCPackage pkg = OPCPackage.open(workbookPath.toFile(), PackageAccess.READ)) {
            SignatureInfo signatureInfo = new SignatureInfo();
            signatureInfo.setSignatureConfig(new SignatureConfig());
            signatureInfo.setOpcPackage(pkg);
            return signatureInfo.verifySignature();
          }
        });
  }

  private static List<String> dataValidationKinds(XSSFSheet sheet) {
    List<String> kinds = new ArrayList<>();
    CTDataValidations dataValidations = sheet.getCTWorksheet().getDataValidations();
    if (dataValidations == null) {
      return List.of();
    }
    for (CTDataValidation validation : dataValidations.getDataValidationArray()) {
      if (validation.getType() == null) {
        kinds.add("UNKNOWN");
        continue;
      }
      if ("list".equalsIgnoreCase(validation.getType().toString())) {
        String formula1 = validation.isSetFormula1() ? validation.getFormula1() : "";
        if ("\"\"".equals(formula1)) {
          kinds.add("EMPTY_EXPLICIT_LIST");
          continue;
        }
        if (formula1.isBlank()) {
          kinds.add("MISSING_FORMULA");
          continue;
        }
      }
      kinds.add(validation.getType().toString().toUpperCase(Locale.ROOT));
    }
    kinds.sort(Comparator.naturalOrder());
    return List.copyOf(kinds);
  }

  private static InputStream decryptedWorkbookStream(POIFSFileSystem fileSystem) {
    return XlsxParitySupport.call(
        "decrypt parity workbook stream",
        () -> {
          EncryptionInfo encryptionInfo = new EncryptionInfo(fileSystem);
          Decryptor decryptor = Decryptor.getInstance(encryptionInfo);
          if (!decryptor.verifyPassword(XlsxParityScenarios.ENCRYPTION_PASSWORD)) {
            throw new IllegalStateException(
                "Failed to verify parity workbook encryption password.");
          }
          return decryptor.getDataStream(fileSystem);
        });
  }

  private static StylesFillSnapshot fillSnapshot(XSSFWorkbook workbook, XSSFCellStyle style) {
    long fillId = style.getCoreXf().getFillId();
    XSSFCellFill fill = workbook.getStylesSource().getFillAt((int) fillId);
    return new StylesFillSnapshot(fill.getCTFill().isSetGradientFill());
  }

  private static String colorDescriptor(XSSFColor color) {
    if (color == null) {
      return "none";
    }
    List<String> parts = new ArrayList<>();
    if (color.isRGB()) {
      parts.add("rgb=" + HexFormat.of().formatHex(color.getRGB()));
    }
    if (color.isThemed()) {
      parts.add("theme=" + color.getTheme());
    }
    if (color.isIndexed()) {
      parts.add("indexed=" + color.getIndexed());
    }
    if (color.hasTint()) {
      parts.add("tint=" + color.getTint());
    }
    return String.join("|", parts);
  }

  private static XSSFCell cell(XSSFSheet sheet, String address) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    var reference = new org.apache.poi.ss.util.CellReference(address);
    return sheet.getRow(reference.getRow()).getCell(reference.getCol());
  }

  private static boolean isFormulaDefined(Name name) {
    String formula = name.getRefersToFormula();
    return formula != null && formula.contains("(");
  }

  private static List<String> sheetNames(XSSFWorkbook workbook) {
    List<String> names = new ArrayList<>();
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      names.add(workbook.getSheetName(sheetIndex));
    }
    return List.copyOf(names);
  }

  private static List<String> selectedSheetNames(XSSFWorkbook workbook) {
    List<String> selected = new ArrayList<>();
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      if (workbook.getSheetAt(sheetIndex).isSelected()) {
        selected.add(workbook.getSheetName(sheetIndex));
      }
    }
    return List.copyOf(selected);
  }

  private static int countOccurrences(String text, String token) {
    int count = 0;
    int index = 0;
    index = text.indexOf(token, index);
    while (index >= 0) {
      count++;
      index += token.length();
      index = text.indexOf(token, index);
    }
    return count;
  }

  private static String sha256(byte[] bytes) {
    return XlsxParitySupport.call(
        "compute SHA-256 digest for parity snapshot",
        () -> HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(bytes)));
  }

  private static int richTextRunCount(XSSFRichTextString richText) {
    int xmlRunCount = richText.getCTRst().sizeOfRArray();
    return xmlRunCount == 0 && !richText.getString().isEmpty() ? 1 : xmlRunCount;
  }

  private static <T> T withWorkbook(Path workbookPath, Function<XSSFWorkbook, T> function) {
    return XlsxParitySupport.call(
        "open parity workbook " + workbookPath,
        () -> {
          try (XSSFWorkbook workbook =
              (XSSFWorkbook) WorkbookFactory.create(workbookPath.toFile())) {
            return function.apply(workbook);
          }
        });
  }

  record CoreWorkbookSnapshot(
      List<String> sheetNames,
      String activeSheetName,
      List<String> selectedSheetNames,
      boolean forceFormulaRecalculation,
      boolean queueHidden,
      boolean opsProtected,
      String mergedRegion,
      String hyperlinkTarget,
      String commentText,
      String commentAuthor,
      boolean commentVisible,
      int zoomPercent,
      String paneDescriptor,
      boolean hasColumnSizing,
      float rowZeroHeightPoints,
      String formulaText,
      double formulaValue,
      String sheetAutofilterRange,
      String tableName,
      List<String> tableColumns,
      int conditionalFormattingBlockCount,
      List<String> dataValidationKinds) {}

  record WorkbookProtectionSnapshot(
      boolean structureLocked,
      boolean windowsLocked,
      boolean revisionLocked,
      boolean workbookPasswordHashPresent,
      boolean revisionsPasswordHashPresent,
      boolean passwordMatches) {}

  record AdvancedPrintSnapshot(
      double leftMargin,
      double rightMargin,
      double topMargin,
      double bottomMargin,
      boolean horizontallyCentered,
      boolean verticallyCentered,
      short paperSize,
      boolean draft,
      boolean noColor,
      short copies,
      boolean usePage,
      short pageStart,
      List<Integer> rowBreaks,
      List<Integer> columnBreaks) {}

  record CommentSnapshot(
      String author, boolean visible, int runCount, int col1, int row1, int col2, int row2) {}

  record StyleSnapshot(
      String fontColorDescriptor,
      String fillColorDescriptor,
      String borderColorDescriptor,
      boolean gradientFill) {}

  record NamedRangeSnapshot(
      String name, String scope, String refersToFormula, boolean formulaDefined) {}

  record AutofilterSnapshot(
      String range, List<FilterColumnSnapshot> filterColumns, SortStateSnapshot sortState) {
    int filterColumnCount() {
      return filterColumns.size();
    }

    boolean hasSortState() {
      return sortState != null;
    }

    String sortStateRange() {
      return sortState == null ? "" : sortState.range();
    }
  }

  record FilterColumnSnapshot(long columnId, List<String> values, boolean includeBlank) {}

  record SortStateSnapshot(
      String range,
      boolean caseSensitive,
      boolean columnSort,
      String sortMethod,
      List<SortConditionSnapshot> conditions) {}

  record SortConditionSnapshot(String range, boolean descending, String sortBy) {}

  record TableSnapshot(
      String comment,
      boolean published,
      boolean insertRow,
      boolean insertRowShift,
      String headerRowCellStyle,
      String dataCellStyle,
      String totalsRowCellStyle,
      String totalsRowLabel,
      String totalsRowFunction,
      String calculatedColumnFormula,
      String uniqueName) {}

  record DrawingSnapshot(
      int shapeCount, int col1, int row1, int col2, int row2, String pictureDigest) {}

  record EmbeddedObjectSnapshot(int shapeCount, String fileName, String objectDigest) {}

  record DrawingSheetSnapshot(
      List<DirectDrawingObjectSnapshot> objects,
      List<String> mergedRegions,
      List<CellCommentSnapshot> comments) {}

  /** Canonical direct-POI snapshot of one authored drawing object. */
  sealed interface DirectDrawingObjectSnapshot
      permits PictureDrawingObjectSnapshot,
          ShapeDrawingObjectSnapshot,
          EmbeddedObjectDrawingObjectSnapshot {
    String name();

    DrawingAnchorSnapshot anchor();
  }

  record PictureDrawingObjectSnapshot(
      String name, DrawingAnchorSnapshot anchor, String pictureDigest)
      implements DirectDrawingObjectSnapshot {}

  record ShapeDrawingObjectSnapshot(
      String name,
      DrawingAnchorSnapshot anchor,
      String kind,
      String presetGeometryToken,
      String text,
      int childCount)
      implements DirectDrawingObjectSnapshot {}

  record EmbeddedObjectDrawingObjectSnapshot(
      String name, DrawingAnchorSnapshot anchor, String fileName, String objectDigest)
      implements DirectDrawingObjectSnapshot {}

  record DrawingAnchorSnapshot(int col1, int row1, int col2, int row2) {}

  record CellCommentSnapshot(String address, String text, String author, boolean visible) {}

  record ChartSnapshot(int chartCount, String title, int barChartCount) {}

  record PivotSnapshot(int pivotCount, String sourceRef, List<Integer> rowLabelColumns) {}

  private record StylesFillSnapshot(boolean gradientFill) {}

  private static DirectDrawingObjectSnapshot drawingObjectSnapshot(
      XSSFDrawing drawing, XSSFShape shape) {
    DrawingAnchorSnapshot anchor = drawingAnchor(shape);
    if (shape instanceof XSSFPicture picture) {
      return new PictureDrawingObjectSnapshot(
          picture.getShapeName(), anchor, sha256(picture.getPictureData().getData()));
    }
    if (shape instanceof XSSFObjectData objectData) {
      byte[] objectBytes =
          XlsxParitySupport.call(
              "read embedded-object bytes from parity workbook", objectData::getObjectData);
      EmbeddedObjectOracleReadback readback = embeddedObjectReadback(objectData, objectBytes);
      return new EmbeddedObjectDrawingObjectSnapshot(
          objectData.getShapeName(), anchor, readback.fileName(), sha256(readback.payloadBytes()));
    }
    if (shape instanceof XSSFConnector connector) {
      return new ShapeDrawingObjectSnapshot(
          connector.getShapeName(), anchor, "CONNECTOR", null, null, 0);
    }
    if (shape instanceof XSSFShapeGroup group) {
      return new ShapeDrawingObjectSnapshot(
          group.getShapeName(), anchor, "GROUP", null, null, drawing.getShapes(group).size());
    }
    if (shape instanceof XSSFGraphicFrame graphicFrame) {
      return new ShapeDrawingObjectSnapshot(
          graphicFrame.getShapeName(), anchor, "GRAPHIC_FRAME", null, null, 0);
    }
    if (shape instanceof XSSFSimpleShape simpleShape) {
      String presetGeometryToken =
          simpleShape.getCTShape().getSpPr() == null
                  || !simpleShape.getCTShape().getSpPr().isSetPrstGeom()
              ? null
              : simpleShape.getCTShape().getSpPr().getPrstGeom().getPrst().toString();
      return new ShapeDrawingObjectSnapshot(
          simpleShape.getShapeName(),
          anchor,
          "SIMPLE_SHAPE",
          presetGeometryToken,
          simpleShape.getText(),
          0);
    }
    throw new IllegalStateException(
        "Unsupported drawing shape type: " + shape.getClass().getName());
  }

  private static DrawingAnchorSnapshot drawingAnchor(XSSFShape shape) {
    if (!(shape.getAnchor() instanceof XSSFClientAnchor anchor)) {
      throw new IllegalStateException(
          "Unsupported drawing anchor type: "
              + (shape.getAnchor() == null ? "null" : shape.getAnchor().getClass().getName()));
    }
    return new DrawingAnchorSnapshot(
        anchor.getCol1(), anchor.getRow1(), anchor.getCol2(), anchor.getRow2());
  }

  private static EmbeddedObjectOracleReadback embeddedObjectReadback(
      XSSFObjectData objectData, byte[] objectBytes) {
    byte[] payloadBytes = objectBytes;
    String fileName = objectData.getFileName();
    try {
      if (looksLikeOle2Storage(objectBytes)) {
        Ole10Native nativeData = ole10Native(objectBytes);
        payloadBytes = nativeData.getDataBuffer();
        fileName = firstNonBlank(nativeData.getFileName2(), nativeData.getFileName(), fileName);
      }
    } catch (Ole10NativeException | RuntimeException exception) {
      // Preserve the truthful raw-package fallback when POI exposes a non-Ole10Native object.
      payloadBytes = objectBytes;
    } catch (java.io.IOException exception) {
      throw new IllegalStateException(
          "Failed to inspect direct POI embedded object payload", exception);
    }
    return new EmbeddedObjectOracleReadback(fileName, payloadBytes);
  }

  private static boolean looksLikeOle2Storage(byte[] bytes) throws java.io.IOException {
    try (ByteArrayInputStream input = new ByteArrayInputStream(bytes)) {
      return FileMagic.valueOf(FileMagic.prepareToCheckMagic(input)) == FileMagic.OLE2;
    }
  }

  private static Ole10Native ole10Native(byte[] bytes)
      throws java.io.IOException, Ole10NativeException {
    try (POIFSFileSystem filesystem = new POIFSFileSystem(new ByteArrayInputStream(bytes))) {
      DirectoryNode directory = filesystem.getRoot();
      return Ole10Native.createFromEmbeddedOleObject(directory);
    }
  }

  private static String firstNonBlank(String first, String second, String fallback) {
    if (first != null && !first.isBlank()) {
      return first;
    }
    if (second != null && !second.isBlank()) {
      return second;
    }
    return fallback;
  }

  /** Immutable direct-POI readback of one embedded-object payload and filename. */
  private static final class EmbeddedObjectOracleReadback {
    private final String fileName;
    private final byte[] payloadBytes;

    private EmbeddedObjectOracleReadback(String fileName, byte[] payloadBytes) {
      this.fileName = fileName;
      this.payloadBytes = payloadBytes.clone();
    }

    private String fileName() {
      return fileName;
    }

    private byte[] payloadBytes() {
      return payloadBytes.clone();
    }
  }
}
