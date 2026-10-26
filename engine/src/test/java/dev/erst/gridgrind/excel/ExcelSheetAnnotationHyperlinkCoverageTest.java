package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.ExcelStyleTestAccess.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.CommentAnchorReport;
import dev.erst.gridgrind.core.model.drawing.CommentReport;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.layout.SheetLayoutReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisLocationReport;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
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

      assertFalse(ExcelSheetStructureSupport.shouldPreview(null));
      assertFalse(ExcelSheetStructureSupport.shouldPreview(blankCell));
      assertTrue(ExcelSheetStructureSupport.shouldPreview(urlCell));
      assertTrue(ExcelSheetStructureSupport.shouldPreview(commentCell));
      assertFalse(
          ExcelSheetStructureSupport.shouldPreview(CellType.BLANK, (short) 0, false, false));
      assertTrue(ExcelSheetStructureSupport.shouldPreview(CellType.BLANK, (short) 1, false, false));
      assertTrue(ExcelSheetStructureSupport.shouldPreview(CellType.BLANK, (short) 0, true, false));
      assertTrue(ExcelSheetStructureSupport.shouldPreview(CellType.BLANK, (short) 0, false, true));
      assertTrue(
          ExcelSheetStructureSupport.shouldPreview(CellType.STRING, (short) 0, false, false));

      assertEquals(Optional.empty(), ExcelSheetAnnotationSupport.hyperlink(blankCell));
      assertEquals(
          Optional.of(new HyperlinkTarget.Url("https://example.com/report")),
          ExcelSheetAnnotationSupport.hyperlink(urlCell));
      assertEquals(
          Optional.of(new HyperlinkTarget.Email("team@example.com")),
          ExcelSheetAnnotationSupport.hyperlink(emailCell));
      assertEquals(
          Optional.of(new HyperlinkTarget.File("/tmp/report.xlsx")),
          ExcelSheetAnnotationSupport.hyperlink(fileCell));
      assertEquals(
          Optional.of(new HyperlinkTarget.Document("Budget!B4")),
          ExcelSheetAnnotationSupport.hyperlink(documentCell));
      assertEquals(
          Optional.empty(), ExcelSheetAnnotationSupport.hyperlink(HyperlinkType.NONE, "ignored"));
      assertEquals(
          Optional.empty(), ExcelSheetAnnotationSupport.hyperlink((HyperlinkType) null, "ignored"));
      assertEquals(
          Optional.empty(), ExcelSheetAnnotationSupport.hyperlink(HyperlinkType.URL, null));
      assertEquals(Optional.empty(), ExcelSheetAnnotationSupport.hyperlink(HyperlinkType.URL, " "));
      assertEquals(
          Optional.empty(), ExcelSheetAnnotationSupport.hyperlink(HyperlinkType.EMAIL, "mailto:"));
      assertEquals(
          Optional.empty(),
          ExcelSheetAnnotationSupport.hyperlink(HyperlinkType.FILE, "https://example.com/report"));

      assertEquals(
          Optional.of(new ResolvedComment("Review", "GridGrind", false)),
          ExcelSheetAnnotationSupport.comment(commentCell));
      assertEquals(Optional.empty(), ExcelSheetAnnotationSupport.comment(blankCell));
      assertEquals(Optional.empty(), ExcelSheetAnnotationSupport.comment(missingStringCommentCell));
      assertEquals(Optional.empty(), ExcelSheetAnnotationSupport.comment(blankCommentCell));
      Cell nullAuthorCommentCell = row.createCell(11);
      nullAuthorCommentCell.setCellComment(comment(poiWorkbook, poiSheet, "Review", null, false));
      assertEquals(Optional.empty(), ExcelSheetAnnotationSupport.comment(nullAuthorCommentCell));
      assertEquals(
          Optional.empty(), ExcelSheetAnnotationSupport.comment((String) null, "GridGrind", false));
      Cell blankAuthorCommentCell = row.createCell(12);
      blankAuthorCommentCell.setCellComment(comment(poiWorkbook, poiSheet, "Review", " ", false));
      assertEquals(Optional.empty(), ExcelSheetAnnotationSupport.comment(blankAuthorCommentCell));

      assertEquals(
          HyperlinkType.URL,
          ExcelSheetAnnotationSupport.toPoi(new HyperlinkTarget.Url("https://example.com")));
      assertEquals(
          HyperlinkType.EMAIL,
          ExcelSheetAnnotationSupport.toPoi(new HyperlinkTarget.Email("team@example.com")));
      assertEquals(
          HyperlinkType.FILE,
          ExcelSheetAnnotationSupport.toPoi(new HyperlinkTarget.File("/tmp/report.xlsx")));
      assertEquals(
          HyperlinkType.DOCUMENT,
          ExcelSheetAnnotationSupport.toPoi(new HyperlinkTarget.Document("Budget!A1")));

      assertEquals(
          "https://example.com/report",
          ExcelSheetAnnotationSupport.toPoiTarget(
              new HyperlinkTarget.Url("https://example.com/report")));
      assertEquals(
          "mailto:team@example.com",
          ExcelSheetAnnotationSupport.toPoiTarget(new HyperlinkTarget.Email("team@example.com")));
      assertEquals(
          Path.of("/tmp/report.xlsx").toUri().toASCIIString(),
          ExcelSheetAnnotationSupport.toPoiTarget(new HyperlinkTarget.File("/tmp/report.xlsx")));
      assertEquals(
          "support/budget%20backup.xlsx",
          ExcelSheetAnnotationSupport.toPoiTarget(
              new HyperlinkTarget.File("support/budget backup.xlsx")));
      assertEquals(
          "Budget!B4",
          ExcelSheetAnnotationSupport.toPoiTarget(new HyperlinkTarget.Document("Budget!B4")));
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

      CommentReport snapshot = ExcelSheetAnnotationSupport.commentSnapshot(cell).orElseThrow();

      assertEquals("Hi there", snapshot.text());
      assertEquals("GridGrind", snapshot.author());
      assertEquals(new CommentAnchorReport(0, 0, 3, 3), snapshot.anchor().orElseThrow());
      assertTrue(snapshot.runs().isPresent());
      assertEquals(2, snapshot.runs().orElseThrow().size());
      assertEquals("Hi ", snapshot.runs().orElseThrow().get(0).text());
      assertEquals("there", snapshot.runs().orElseThrow().get(1).text());
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

      assertTrue(ExcelSheetAnnotationSupport.commentSnapshot((Cell) null).isEmpty());
      assertTrue(ExcelSheetAnnotationSupport.commentSnapshot((Comment) null).isEmpty());
      assertTrue(ExcelSheetAnnotationSupport.commentSnapshot(blankCell).isEmpty());
      assertTrue(ExcelSheetAnnotationSupport.commentSnapshot(emptyStringCell).isEmpty());
      assertTrue(
          ExcelSheetAnnotationSupport.commentSnapshot(blankAuthorCell.getCellComment()).isEmpty());
      assertTrue(
          ExcelSheetAnnotationSupport.commentSnapshot(blankTextCell.getCellComment()).isEmpty());
      assertTrue(
          ExcelSheetAnnotationSupport.commentSnapshot(nullAuthorCell.getCellComment()).isEmpty());

      CommentReport directSnapshot =
          ExcelSheetAnnotationSupport.commentSnapshot(validCommentCell.getCellComment())
              .orElseThrow();
      assertEquals("Review", directSnapshot.text());
      assertEquals("GridGrind", directSnapshot.author());
      assertTrue(directSnapshot.visible());
      assertTrue(directSnapshot.runs().isEmpty());

      CommentReport noAnchorSnapshot =
          ExcelSheetAnnotationSupport.commentSnapshot(
                  commentWithoutAnchor(poiSheet, validCommentCell.getCellComment()))
              .orElseThrow();
      assertEquals("Review", noAnchorSnapshot.text());
      assertTrue(noAnchorSnapshot.anchor().isEmpty());
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

      assertTrue(ExcelSheetAnnotationSupport.commentSnapshot(comment).isEmpty());
      assertTrue(ExcelSheetAnnotationSupport.commentSnapshot(cell).isEmpty());
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

      List<AnalysisFindingReport> findings;
      try {
        findings = sheet.diagnostics().hyperlinkHealthFindings();
      } finally {
        Files.deleteIfExists(reachableFile);
      }

      assertEquals(5, sheet.diagnostics().hyperlinkCount());
      assertTrue(
          findings.stream()
              .map(AnalysisFindingReport::code)
              .toList()
              .containsAll(
                  List.of(
                      AnalysisFindingCode.HYPERLINK_MISSING_DOCUMENT_SHEET,
                      AnalysisFindingCode.HYPERLINK_INVALID_DOCUMENT_TARGET)));
      assertTrue(
          findings.stream()
              .noneMatch(
                  finding ->
                      finding.location() instanceof AnalysisLocationReport.Cell cell
                          && "A5".equals(cell.address())));
    }
  }

  @Test
  void fileHyperlinkFindingsCoverMalformedMissingAndUnresolvedTargets() {
    AnalysisLocationReport.Cell location =
        new AnalysisLocationReport.Cell(
            new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"), "A1");
    WorkbookLocation storedWorkbook =
        new WorkbookLocation.StoredWorkbook(
            Path.of("tmp", "file-hyperlink-findings", "Budget.xlsx").toAbsolutePath());

    List<AnalysisFindingReport> malformed =
        ExcelSheetAnalysisSupport.fileHyperlinkFindings(
            location, "https://example.com/report", storedWorkbook);
    assertEquals(1, malformed.size());
    assertEquals(AnalysisFindingCode.HYPERLINK_MALFORMED_TARGET, malformed.getFirst().code());

    List<AnalysisFindingReport> unresolved =
        ExcelSheetAnalysisSupport.fileHyperlinkFindings(
            location, "reports/q1.xlsx", new WorkbookLocation.UnsavedWorkbook());
    assertEquals(1, unresolved.size());
    assertEquals(
        AnalysisFindingCode.HYPERLINK_UNRESOLVED_FILE_TARGET, unresolved.getFirst().code());

    List<AnalysisFindingReport> missing =
        ExcelSheetAnalysisSupport.fileHyperlinkFindings(
            location, "reports/q1.xlsx", storedWorkbook);
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
    List<AnalysisFindingReport> missingAbsolute =
        ExcelSheetAnalysisSupport.fileHyperlinkFindings(
            location, absoluteMissingTarget, new WorkbookLocation.UnsavedWorkbook());
    assertEquals(1, missingAbsolute.size());
    assertEquals(
        AnalysisFindingCode.HYPERLINK_MISSING_FILE_TARGET, missingAbsolute.getFirst().code());
    assertFalse(missingAbsolute.getFirst().message().contains("workbook directory"));

    List<AnalysisFindingReport> missingAbsoluteStored =
        ExcelSheetAnalysisSupport.fileHyperlinkFindings(
            location, absoluteMissingTarget, storedWorkbook);
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
    AnalysisLocationReport.Cell location =
        new AnalysisLocationReport.Cell(
            new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"), "A1");

    List<AnalysisFindingReport> malformedUrl =
        ExcelSheetAnalysisSupport.externalHyperlinkFindings(
            location, "example.com/report", "URL", false);
    assertEquals(1, malformedUrl.size());
    assertEquals(AnalysisFindingCode.HYPERLINK_MALFORMED_TARGET, malformedUrl.getFirst().code());

    List<AnalysisFindingReport> malformedEmail =
        ExcelSheetAnalysisSupport.externalHyperlinkFindings(location, "mailto:", "EMAIL", false);
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

      sheet.cells().setCell("A1", new CellValue.Text("first"));
      sheet.cells().setCell("A3", new CellValue.Text("third"));

      WorkbookSheetResult.Window window = sheet.cells().window("A1", 3, 1);
      assertEquals("A2", window.rows().get(1).cells().getFirst().address().value());
      assertEquals("BLANK", window.rows().get(1).cells().getFirst().type());

      SheetLayoutReport layout = sheet.layout().snapshot();
      assertEquals(3, layout.rows().size());
      assertEquals(poiSheet.getDefaultRowHeightInPoints(), layout.rows().get(1).heightPoints());

      assertEquals(
          Optional.empty(),
          ExcelSheetAnnotationSupport.hyperlink(HyperlinkType.URL, "example.com/report"));
    }
  }

  @Test
  void helperMethodsHandleFormulaAndHyperlinkEdgeCases() throws Exception {
    assertTrue(ExcelSheetAnalysisSupport.containsExternalWorkbookReference("[Book.xlsx]Sheet1!A1"));
    assertFalse(ExcelSheetAnalysisSupport.containsExternalWorkbookReference("SUM(A1:A2)"));
    assertFalse(ExcelSheetAnalysisSupport.containsExternalWorkbookReference("[Book.xlsx"));
    assertFalse(ExcelSheetAnalysisSupport.containsExternalWorkbookReference("Book.xlsx]"));

    assertEquals(
        List.of("NOW", "INDIRECT"),
        ExcelSheetAnalysisSupport.volatileFunctions("NOW()+INDIRECT(\"A1\")"));
    assertEquals(List.of(), ExcelSheetAnalysisSupport.volatileFunctions("SUM(A1:A2)"));

    assertEquals("Quarter 1", ExcelSheetAnalysisSupport.unquoteSheetName("'Quarter 1'"));
    assertEquals("Budget", ExcelSheetAnalysisSupport.unquoteSheetName("Budget"));
    assertEquals("'", ExcelSheetAnalysisSupport.unquoteSheetName("'"));
    assertEquals("'Budget", ExcelSheetAnalysisSupport.unquoteSheetName("'Budget"));
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      ExcelSheet sheet =
          new ExcelSheet(
              poiWorkbook.createSheet("Helpers"),
              new WorkbookStyleRegistry(poiWorkbook),
              poiWorkbook.getCreationHelper().createFormulaEvaluator());
      assertEquals("IllegalStateException", sheet.exceptionMessage(new IllegalStateException()));
      assertEquals(
          "display failure", sheet.exceptionMessage(new IllegalStateException("display failure")));
      AnalysisLocationReport.Cell location =
          new AnalysisLocationReport.Cell(
              new dev.erst.gridgrind.core.vocabulary.SheetLabel("Helpers"), "A1");
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

    assertEquals(Optional.empty(), ExcelSheetAnnotationSupport.hyperlink((Cell) null));
    assertFalse(
        ExcelSheetAnalysisSupport.hasUsableHyperlink((org.apache.poi.ss.usermodel.Hyperlink) null));
    assertFalse(ExcelSheetAnalysisSupport.hasUsableHyperlinkType(null));
    assertFalse(ExcelSheetAnalysisSupport.hasUsableHyperlinkType(HyperlinkType.NONE));
    assertTrue(ExcelSheetAnalysisSupport.hasUsableHyperlinkType(HyperlinkType.URL));
    assertEquals(Optional.empty(), ExcelSheetAnnotationSupport.hyperlink(hyperlinkWithNullType()));

    assertTrue(ExcelSheetAnalysisSupport.hasMissingHyperlinkTarget(null));
    assertTrue(ExcelSheetAnalysisSupport.hasMissingHyperlinkTarget(" "));
    assertFalse(ExcelSheetAnalysisSupport.hasMissingHyperlinkTarget("Budget!A1"));
    assertEquals(Optional.empty(), ExcelSheetAnnotationSupport.comment((Cell) null));
    assertEquals(Optional.empty(), ExcelSheetAnnotationSupport.comment((Comment) null));
  }

  @Test
  void validateDocumentHyperlinkTargetHandlesMalformedMissingAndRangeTargets() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      poiWorkbook.createSheet("Quarter 1");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);
      AnalysisLocationReport.Cell location =
          new AnalysisLocationReport.Cell(
              new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"), "A1");

      List<AnalysisFindingReport> invalidStructure = new ArrayList<>();
      sheet.validateDocumentHyperlinkTarget(location, "!A1", invalidStructure);
      assertEquals(
          AnalysisFindingCode.HYPERLINK_INVALID_DOCUMENT_TARGET,
          invalidStructure.getFirst().code());

      List<AnalysisFindingReport> missingSheet = new ArrayList<>();
      sheet.validateDocumentHyperlinkTarget(location, "Missing!A1", missingSheet);
      assertEquals(
          AnalysisFindingCode.HYPERLINK_MISSING_DOCUMENT_SHEET, missingSheet.getFirst().code());

      List<AnalysisFindingReport> invalidRange = new ArrayList<>();
      sheet.validateDocumentHyperlinkTarget(location, "Budget!A1:", invalidRange);
      assertEquals(
          AnalysisFindingCode.HYPERLINK_INVALID_DOCUMENT_TARGET, invalidRange.getFirst().code());

      List<AnalysisFindingReport> validRange = new ArrayList<>();
      sheet.validateDocumentHyperlinkTarget(location, "Budget!A1:B2", validRange);
      assertEquals(List.of(), validRange);

      List<AnalysisFindingReport> quotedValid = new ArrayList<>();
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

      assertEquals(2, sheet.diagnostics().hyperlinkCount());
      assertEquals(List.of(), sheet.diagnostics().hyperlinkHealthFindings());
    }
  }

  @Test
  void snapshotCellReturnsBlankForUnwrittenCells() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.cells().setCell("A1", new CellValue.Text("Header"));

      // A cell in a non-existent row returns blank.
      CellReport blankInNewRow = sheet.cells().snapshotCell("C5");
      assertInstanceOf(CellReport.BlankReport.class, blankInNewRow);
      assertEquals("C5", blankInNewRow.address().value());

      // A cell in a row that exists but in a column that has not been written returns blank.
      CellReport blankInExistingRow = sheet.cells().snapshotCell("B1");
      assertInstanceOf(CellReport.BlankReport.class, blankInExistingRow);
      assertEquals("B1", blankInExistingRow.address().value());
    }
  }
}
