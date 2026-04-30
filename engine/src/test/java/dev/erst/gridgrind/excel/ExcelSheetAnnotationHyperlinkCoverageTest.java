package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.ExcelStyleTestAccess.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** ExcelSheet comment, hyperlink, and blank-cell snapshot coverage. */
class ExcelSheetAnnotationHyperlinkCoverageTest extends ExcelSheetTestSupport {
  @Test
  void interpretsHyperlinksCommentsAndPreviewHelpersAcrossAllVariants() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      Row row = poiSheet.createRow(0);
      Cell blankCell = row.createCell(0);
      Cell urlCell = row.createCell(1);
      Cell emailCell = row.createCell(2);
      Cell fileCell = row.createCell(3);
      Cell documentCell = row.createCell(4);
      Cell commentCell = row.createCell(5);
      Cell missingStringCommentCell = row.createCell(6);
      Cell blankCommentCell = row.createCell(7);

      urlCell.setHyperlink(hyperlink(poiWorkbook, HyperlinkType.URL, "https://example.com/report"));
      emailCell.setHyperlink(
          hyperlink(poiWorkbook, HyperlinkType.EMAIL, "mailto:team@example.com"));
      fileCell.setHyperlink(hyperlink(poiWorkbook, HyperlinkType.FILE, "/tmp/report.xlsx"));
      documentCell.setHyperlink(hyperlink(poiWorkbook, HyperlinkType.DOCUMENT, "Budget!B4"));
      commentCell.setCellComment(comment(poiWorkbook, poiSheet, "Review", "GridGrind", false));
      missingStringCommentCell.setCellComment(emptyComment(poiWorkbook, poiSheet));
      blankCommentCell.setCellComment(comment(poiWorkbook, poiSheet, " ", " ", false));

      assertFalse(ExcelSheet.shouldPreview(null));
      assertFalse(ExcelSheet.shouldPreview(blankCell));
      assertTrue(ExcelSheet.shouldPreview(urlCell));
      assertTrue(ExcelSheet.shouldPreview(commentCell));
      assertFalse(ExcelSheet.shouldPreview(CellType.BLANK, (short) 0, false, false));
      assertTrue(ExcelSheet.shouldPreview(CellType.BLANK, (short) 1, false, false));
      assertTrue(ExcelSheet.shouldPreview(CellType.BLANK, (short) 0, true, false));
      assertTrue(ExcelSheet.shouldPreview(CellType.BLANK, (short) 0, false, true));
      assertTrue(ExcelSheet.shouldPreview(CellType.STRING, (short) 0, false, false));

      assertNull(ExcelSheet.hyperlink(blankCell));
      assertEquals(
          new ExcelHyperlink.Url("https://example.com/report"), ExcelSheet.hyperlink(urlCell));
      assertEquals(new ExcelHyperlink.Email("team@example.com"), ExcelSheet.hyperlink(emailCell));
      assertEquals(new ExcelHyperlink.File("/tmp/report.xlsx"), ExcelSheet.hyperlink(fileCell));
      assertEquals(new ExcelHyperlink.Document("Budget!B4"), ExcelSheet.hyperlink(documentCell));
      assertNull(ExcelSheet.hyperlink(HyperlinkType.NONE, "ignored"));
      assertNull(ExcelSheet.hyperlink((HyperlinkType) null, "ignored"));
      assertNull(ExcelSheet.hyperlink(HyperlinkType.URL, null));
      assertNull(ExcelSheet.hyperlink(HyperlinkType.URL, " "));
      assertNull(ExcelSheet.hyperlink(HyperlinkType.EMAIL, "mailto:"));
      assertNull(ExcelSheet.hyperlink(HyperlinkType.FILE, "https://example.com/report"));

      assertEquals(new ExcelComment("Review", "GridGrind", false), ExcelSheet.comment(commentCell));
      assertNull(ExcelSheet.comment(blankCell));
      assertNull(ExcelSheet.comment(missingStringCommentCell));
      assertNull(ExcelSheet.comment(blankCommentCell));
      Cell nullAuthorCommentCell = row.createCell(11);
      nullAuthorCommentCell.setCellComment(comment(poiWorkbook, poiSheet, "Review", null, false));
      assertNull(ExcelSheet.comment(nullAuthorCommentCell));
      assertNull(ExcelSheet.comment((String) null, "GridGrind", false));
      Cell blankAuthorCommentCell = row.createCell(12);
      blankAuthorCommentCell.setCellComment(comment(poiWorkbook, poiSheet, "Review", " ", false));
      assertNull(ExcelSheet.comment(blankAuthorCommentCell));

      assertEquals(HyperlinkType.URL, ExcelSheet.toPoi(ExcelHyperlinkType.URL));
      assertEquals(HyperlinkType.EMAIL, ExcelSheet.toPoi(ExcelHyperlinkType.EMAIL));
      assertEquals(HyperlinkType.FILE, ExcelSheet.toPoi(ExcelHyperlinkType.FILE));
      assertEquals(HyperlinkType.DOCUMENT, ExcelSheet.toPoi(ExcelHyperlinkType.DOCUMENT));

      assertEquals(
          "https://example.com/report",
          ExcelSheet.toPoiTarget(new ExcelHyperlink.Url("https://example.com/report")));
      assertEquals(
          "mailto:team@example.com",
          ExcelSheet.toPoiTarget(new ExcelHyperlink.Email("team@example.com")));
      assertEquals(
          Path.of("/tmp/report.xlsx").toUri().toASCIIString(),
          ExcelSheet.toPoiTarget(new ExcelHyperlink.File("/tmp/report.xlsx")));
      assertEquals(
          "support/budget%20backup.xlsx",
          ExcelSheet.toPoiTarget(new ExcelHyperlink.File("support/budget backup.xlsx")));
      assertEquals("Budget!B4", ExcelSheet.toPoiTarget(new ExcelHyperlink.Document("Budget!B4")));
    }
  }

  @Test
  void commentSnapshotPreservesRichRunsAndAnchorMetadata() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      Row row = poiSheet.createRow(0);
      Cell cell = row.createCell(0);
      Comment comment = emptyComment(poiWorkbook, poiSheet);
      org.apache.poi.xssf.usermodel.XSSFRichTextString richText =
          new org.apache.poi.xssf.usermodel.XSSFRichTextString();
      richText.append("Hi ");
      var accentFont = poiWorkbook.createFont();
      accentFont.setBold(true);
      richText.append("there", accentFont);
      comment.setString(richText);
      comment.setAuthor("GridGrind");
      comment.setVisible(true);
      cell.setCellComment(comment);

      ExcelCommentSnapshot snapshot = ExcelSheet.commentSnapshot(cell);

      assertEquals("Hi there", snapshot.text());
      assertEquals("GridGrind", snapshot.author());
      assertEquals(new ExcelCommentAnchorSnapshot(0, 0, 3, 3), snapshot.anchor());
      assertNotNull(snapshot.runs());
      assertEquals(2, snapshot.runs().runs().size());
      assertEquals("Hi ", snapshot.runs().runs().get(0).text());
      assertEquals("there", snapshot.runs().runs().get(1).text());
    }
  }

  @Test
  void commentSnapshotReturnsNullForMissingIncompleteAndNonXssfComments() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      Row row = poiSheet.createRow(0);
      Cell blankCell = row.createCell(0);
      Cell emptyStringCell = row.createCell(1);
      emptyStringCell.setCellComment(emptyComment(poiWorkbook, poiSheet));
      Cell blankAuthorCell = row.createCell(2);
      blankAuthorCell.setCellComment(comment(poiWorkbook, poiSheet, "Review", " ", false));
      Cell validCommentCell = row.createCell(3);
      validCommentCell.setCellComment(comment(poiWorkbook, poiSheet, "Review", "GridGrind", true));
      Cell blankTextCell = row.createCell(4);
      blankTextCell.setCellComment(comment(poiWorkbook, poiSheet, " ", "GridGrind", false));
      Cell nullAuthorCell = row.createCell(5);
      nullAuthorCell.setCellComment(comment(poiWorkbook, poiSheet, "Review", null, false));

      assertNull(ExcelSheet.commentSnapshot((Cell) null));
      assertNull(ExcelSheet.commentSnapshot((Comment) null));
      assertNull(ExcelSheet.commentSnapshot(blankCell));
      assertNull(ExcelSheet.commentSnapshot(emptyStringCell));
      assertNull(ExcelSheet.commentSnapshot(blankAuthorCell.getCellComment()));
      assertNull(ExcelSheet.commentSnapshot(blankTextCell.getCellComment()));
      assertNull(ExcelSheet.commentSnapshot(nullAuthorCell.getCellComment()));

      ExcelCommentSnapshot directSnapshot =
          ExcelSheet.commentSnapshot(validCommentCell.getCellComment());
      assertEquals("Review", directSnapshot.text());
      assertEquals("GridGrind", directSnapshot.author());
      assertTrue(directSnapshot.visible());
      assertNull(directSnapshot.runs());

      ExcelCommentSnapshot noAnchorSnapshot =
          ExcelSheet.commentSnapshot(
              commentWithoutAnchor(poiSheet, validCommentCell.getCellComment()));
      assertEquals("Review", noAnchorSnapshot.text());
      assertNull(noAnchorSnapshot.anchor());
    }

    try (var poiWorkbook = new org.apache.poi.hssf.usermodel.HSSFWorkbook()) {
      var poiSheet = poiWorkbook.createSheet("Legacy");
      var row = poiSheet.createRow(0);
      var cell = row.createCell(0);
      var drawing = poiSheet.createDrawingPatriarch();
      var comment = drawing.createCellComment(new org.apache.poi.hssf.usermodel.HSSFClientAnchor());
      comment.setString(new org.apache.poi.hssf.usermodel.HSSFRichTextString("Legacy"));
      comment.setAuthor("GridGrind");
      cell.setCellComment(comment);

      assertNull(ExcelSheet.commentSnapshot(comment));
      assertNull(ExcelSheet.commentSnapshot(cell));
    }
  }

  @Test
  void derivesHyperlinkHealthFindingsAcrossMalformedAndDocumentTargets() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      poiWorkbook.createSheet("Quarter 1");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      Path reachableFile =
          ExcelTempFiles.createManagedTempFile("gridgrind-hyperlink-health-", ".xlsx");
      Cell validFileCell = poiSheet.createRow(0).createCell(0);
      validFileCell.setHyperlink(
          hyperlink(poiWorkbook, HyperlinkType.FILE, reachableFile.toString()));
      Cell missingSheetCell = poiSheet.createRow(1).createCell(0);
      missingSheetCell.setHyperlink(hyperlink(poiWorkbook, HyperlinkType.DOCUMENT, "Missing!A1"));
      Cell invalidDocumentCell = poiSheet.createRow(2).createCell(0);
      invalidDocumentCell.setHyperlink(hyperlink(poiWorkbook, HyperlinkType.DOCUMENT, "Budget!"));
      Cell invalidDocumentRangeCell = poiSheet.createRow(3).createCell(0);
      invalidDocumentRangeCell.setHyperlink(
          hyperlink(poiWorkbook, HyperlinkType.DOCUMENT, "Budget!A1:"));
      Cell quotedDocumentCell = poiSheet.createRow(4).createCell(0);
      quotedDocumentCell.setHyperlink(
          hyperlink(poiWorkbook, HyperlinkType.DOCUMENT, "'Quarter 1'!A1"));

      List<WorkbookAnalysis.AnalysisFinding> findings;
      try {
        findings = sheet.hyperlinkHealthFindings();
      } finally {
        Files.deleteIfExists(reachableFile);
      }

      assertEquals(5, sheet.hyperlinkCount());
      assertTrue(
          findings.stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList()
              .containsAll(
                  List.of(
                      AnalysisFindingCode.HYPERLINK_MISSING_DOCUMENT_SHEET,
                      AnalysisFindingCode.HYPERLINK_INVALID_DOCUMENT_TARGET)));
      assertTrue(
          findings.stream()
              .noneMatch(
                  finding ->
                      finding.location() instanceof WorkbookAnalysis.AnalysisLocation.Cell cell
                          && "A5".equals(cell.address())));
    }
  }

  @Test
  void fileHyperlinkFindingsCoverMalformedMissingAndUnresolvedTargets() {
    WorkbookAnalysis.AnalysisLocation.Cell location =
        new WorkbookAnalysis.AnalysisLocation.Cell("Budget", "A1");
    WorkbookLocation storedWorkbook =
        new WorkbookLocation.StoredWorkbook(
            Path.of("tmp", "file-hyperlink-findings", "Budget.xlsx").toAbsolutePath());

    List<WorkbookAnalysis.AnalysisFinding> malformed =
        ExcelSheet.fileHyperlinkFindings(location, "https://example.com/report", storedWorkbook);
    assertEquals(1, malformed.size());
    assertEquals(AnalysisFindingCode.HYPERLINK_MALFORMED_TARGET, malformed.getFirst().code());

    List<WorkbookAnalysis.AnalysisFinding> unresolved =
        ExcelSheet.fileHyperlinkFindings(
            location, "reports/q1.xlsx", new WorkbookLocation.UnsavedWorkbook());
    assertEquals(1, unresolved.size());
    assertEquals(
        AnalysisFindingCode.HYPERLINK_UNRESOLVED_FILE_TARGET, unresolved.getFirst().code());

    List<WorkbookAnalysis.AnalysisFinding> missing =
        ExcelSheet.fileHyperlinkFindings(location, "reports/q1.xlsx", storedWorkbook);
    assertEquals(1, missing.size());
    assertEquals(AnalysisFindingCode.HYPERLINK_MISSING_FILE_TARGET, missing.getFirst().code());
    assertTrue(missing.getFirst().message().contains("workbook directory"));
    assertTrue(
        missing
            .getFirst()
            .evidence()
            .contains(storedWorkbook.baseDirectory().orElseThrow().toString()));

    String absoluteMissingTarget =
        Path.of("tmp", "file-hyperlink-findings", "missing-report.xlsx")
            .toAbsolutePath()
            .toString();
    List<WorkbookAnalysis.AnalysisFinding> missingAbsolute =
        ExcelSheet.fileHyperlinkFindings(
            location, absoluteMissingTarget, new WorkbookLocation.UnsavedWorkbook());
    assertEquals(1, missingAbsolute.size());
    assertEquals(
        AnalysisFindingCode.HYPERLINK_MISSING_FILE_TARGET, missingAbsolute.getFirst().code());
    assertFalse(missingAbsolute.getFirst().message().contains("workbook directory"));

    List<WorkbookAnalysis.AnalysisFinding> missingAbsoluteStored =
        ExcelSheet.fileHyperlinkFindings(location, absoluteMissingTarget, storedWorkbook);
    assertEquals(1, missingAbsoluteStored.size());
    assertEquals(
        AnalysisFindingCode.HYPERLINK_MISSING_FILE_TARGET, missingAbsoluteStored.getFirst().code());
    assertFalse(missingAbsoluteStored.getFirst().message().contains("workbook directory"));
    assertFalse(
        missingAbsoluteStored
            .getFirst()
            .evidence()
            .contains(storedWorkbook.baseDirectory().orElseThrow().toString()));
  }

  @Test
  void externalHyperlinkFindingsCoverMalformedUrlAndEmailTargets() {
    WorkbookAnalysis.AnalysisLocation.Cell location =
        new WorkbookAnalysis.AnalysisLocation.Cell("Budget", "A1");

    List<WorkbookAnalysis.AnalysisFinding> malformedUrl =
        ExcelSheet.externalHyperlinkFindings(location, "example.com/report", "URL", false);
    assertEquals(1, malformedUrl.size());
    assertEquals(AnalysisFindingCode.HYPERLINK_MALFORMED_TARGET, malformedUrl.getFirst().code());

    List<WorkbookAnalysis.AnalysisFinding> malformedEmail =
        ExcelSheet.externalHyperlinkFindings(location, "mailto:", "EMAIL", false);
    assertEquals(1, malformedEmail.size());
    assertEquals(AnalysisFindingCode.HYPERLINK_MALFORMED_TARGET, malformedEmail.getFirst().code());
  }

  @Test
  void normalizesSparseWindowsAndInvalidHyperlinkHelpers() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Sparse");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.setCell("A1", ExcelCellValue.text("first"));
      sheet.setCell("A3", ExcelCellValue.text("third"));

      WorkbookSheetResult.Window window = sheet.window("A1", 3, 1);
      assertEquals("A2", window.rows().get(1).cells().getFirst().address());
      assertEquals("BLANK", window.rows().get(1).cells().getFirst().effectiveType());

      WorkbookSheetResult.SheetLayout layout = sheet.layout();
      assertEquals(3, layout.rows().size());
      assertEquals(poiSheet.getDefaultRowHeightInPoints(), layout.rows().get(1).heightPoints());

      assertNull(ExcelSheet.hyperlink(HyperlinkType.URL, "example.com/report"));
    }
  }

  @Test
  void helperMethodsHandleFormulaAndHyperlinkEdgeCases() throws Exception {
    assertTrue(ExcelSheet.containsExternalWorkbookReference("[Book.xlsx]Sheet1!A1"));
    assertFalse(ExcelSheet.containsExternalWorkbookReference("SUM(A1:A2)"));
    assertFalse(ExcelSheet.containsExternalWorkbookReference("[Book.xlsx"));
    assertFalse(ExcelSheet.containsExternalWorkbookReference("Book.xlsx]"));

    assertEquals(
        List.of("NOW", "INDIRECT"), ExcelSheet.volatileFunctions("NOW()+INDIRECT(\"A1\")"));
    assertEquals(List.of(), ExcelSheet.volatileFunctions("SUM(A1:A2)"));

    assertEquals("Quarter 1", ExcelSheet.unquoteSheetName("'Quarter 1'"));
    assertEquals("Budget", ExcelSheet.unquoteSheetName("Budget"));
    assertEquals("'", ExcelSheet.unquoteSheetName("'"));
    assertEquals("'Budget", ExcelSheet.unquoteSheetName("'Budget"));
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      ExcelSheet sheet =
          new ExcelSheet(
              poiWorkbook.createSheet("Helpers"),
              new WorkbookStyleRegistry(poiWorkbook),
              poiWorkbook.getCreationHelper().createFormulaEvaluator());
      assertEquals("IllegalStateException", sheet.exceptionMessage(new IllegalStateException()));
      assertEquals(
          "display failure", sheet.exceptionMessage(new IllegalStateException("display failure")));
      WorkbookAnalysis.AnalysisLocation.Cell location =
          new WorkbookAnalysis.AnalysisLocation.Cell("Helpers", "A1");
      assertEquals(
          AnalysisFindingCode.HYPERLINK_MALFORMED_TARGET,
          sheet
              .hyperlinkTargetFindings(
                  location, HyperlinkType.URL, " ", new WorkbookLocation.UnsavedWorkbook())
              .getFirst()
              .code());
      assertEquals(
          List.of(),
          sheet.hyperlinkTargetFindings(
              location,
              HyperlinkType.EMAIL,
              "team@example.com",
              new WorkbookLocation.UnsavedWorkbook()));
      assertEquals(
          List.of(),
          sheet.hyperlinkTargetFindings(
              location, HyperlinkType.NONE, "ignored", new WorkbookLocation.UnsavedWorkbook()));
    }

    assertNull(ExcelSheet.hyperlink((Cell) null));
    assertFalse(ExcelSheet.hasUsableHyperlink((org.apache.poi.ss.usermodel.Hyperlink) null));
    assertFalse(ExcelSheet.hasUsableHyperlinkType(null));
    assertFalse(ExcelSheet.hasUsableHyperlinkType(HyperlinkType.NONE));
    assertTrue(ExcelSheet.hasUsableHyperlinkType(HyperlinkType.URL));
    assertNull(ExcelSheet.hyperlink(hyperlinkWithNullType()));

    assertTrue(ExcelSheet.hasMissingHyperlinkTarget(null));
    assertTrue(ExcelSheet.hasMissingHyperlinkTarget(" "));
    assertFalse(ExcelSheet.hasMissingHyperlinkTarget("Budget!A1"));
    assertNull(ExcelSheet.comment((Cell) null));
    assertNull(ExcelSheet.comment((Comment) null));
  }

  @Test
  void validateDocumentHyperlinkTargetHandlesMalformedMissingAndRangeTargets() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      poiWorkbook.createSheet("Quarter 1");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);
      WorkbookAnalysis.AnalysisLocation.Cell location =
          new WorkbookAnalysis.AnalysisLocation.Cell("Budget", "A1");

      List<WorkbookAnalysis.AnalysisFinding> invalidStructure = new ArrayList<>();
      sheet.validateDocumentHyperlinkTarget(location, "!A1", invalidStructure);
      assertEquals(
          AnalysisFindingCode.HYPERLINK_INVALID_DOCUMENT_TARGET,
          invalidStructure.getFirst().code());

      List<WorkbookAnalysis.AnalysisFinding> missingSheet = new ArrayList<>();
      sheet.validateDocumentHyperlinkTarget(location, "Missing!A1", missingSheet);
      assertEquals(
          AnalysisFindingCode.HYPERLINK_MISSING_DOCUMENT_SHEET, missingSheet.getFirst().code());

      List<WorkbookAnalysis.AnalysisFinding> invalidRange = new ArrayList<>();
      sheet.validateDocumentHyperlinkTarget(location, "Budget!A1:", invalidRange);
      assertEquals(
          AnalysisFindingCode.HYPERLINK_INVALID_DOCUMENT_TARGET, invalidRange.getFirst().code());

      List<WorkbookAnalysis.AnalysisFinding> validRange = new ArrayList<>();
      sheet.validateDocumentHyperlinkTarget(location, "Budget!A1:B2", validRange);
      assertEquals(List.of(), validRange);

      List<WorkbookAnalysis.AnalysisFinding> quotedValid = new ArrayList<>();
      sheet.validateDocumentHyperlinkTarget(location, "'Quarter 1'!A1", quotedValid);
      assertEquals(List.of(), quotedValid);
    }
  }

  @Test
  void hyperlinkHealthTreatsEmptyCellsAndNoneLinksAsNonFindings() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      poiSheet.createRow(0).createCell(0).setCellValue("plain");
      Cell noneCell = poiSheet.createRow(1).createCell(0);
      noneCell.setHyperlink(hyperlink(poiWorkbook, HyperlinkType.NONE, "ignored"));
      Cell validUrlCell = poiSheet.createRow(2).createCell(0);
      validUrlCell.setHyperlink(hyperlink(poiWorkbook, HyperlinkType.URL, "https://example.com"));
      Cell validEmailCell = poiSheet.createRow(3).createCell(0);
      validEmailCell.setHyperlink(hyperlink(poiWorkbook, HyperlinkType.EMAIL, "team@example.com"));

      assertEquals(2, sheet.hyperlinkCount());
      assertEquals(List.of(), sheet.hyperlinkHealthFindings());
    }
  }

  @Test
  void snapshotCellReturnsBlankForUnwrittenCells() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.setCell("A1", ExcelCellValue.text("Header"));

      // A cell in a non-existent row returns blank.
      ExcelCellSnapshot blankInNewRow = sheet.snapshotCell("C5");
      assertInstanceOf(ExcelCellSnapshot.BlankSnapshot.class, blankInNewRow);
      assertEquals("C5", blankInNewRow.address());

      // A cell in a row that exists but in a column that has not been written returns blank.
      ExcelCellSnapshot blankInExistingRow = sheet.snapshotCell("B1");
      assertInstanceOf(ExcelCellSnapshot.BlankSnapshot.class, blankInExistingRow);
      assertEquals("B1", blankInExistingRow.address());
    }
  }
}
