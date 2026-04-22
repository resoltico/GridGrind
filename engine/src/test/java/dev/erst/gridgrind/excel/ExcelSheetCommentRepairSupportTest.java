package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Method;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipFile;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.model.CommentsTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFFactory;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTComments;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

/** Verifies authoritative comment repair after column structure mutations and sheet copies. */
@SuppressWarnings("StringConcatToTextBlock")
class ExcelSheetCommentRepairSupportTest {
  @Test
  void commentsAfterShiftColumnsKeepsMovedCommentWhenTargetsCollide() {
    List<WorkbookReadResult.CellComment> shifted =
        ExcelSheetCommentRepairSupport.commentsAfterShiftColumns(
            List.of(
                new WorkbookReadResult.CellComment(
                    "A2", new ExcelCommentSnapshot("stationary", "GridGrind", true, null, null)),
                new WorkbookReadResult.CellComment(
                    "B2", new ExcelCommentSnapshot("moving", "GridGrind", true, null, null))),
            new ExcelColumnSpan(1, 1),
            -1);

    assertEquals(
        List.of(
            new WorkbookReadResult.CellComment(
                "A2", new ExcelCommentSnapshot("moving", "GridGrind", true, null, null))),
        shifted);
  }

  @Test
  void hasPersistedCommentsAndRawCommentAddressesHandleEmptyState() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Ops");

      ExcelSheetCommentRepairSupport support =
          new ExcelSheetCommentRepairSupport(workbook.sheet("Ops").xssfSheet());

      assertFalse(support.hasPersistedComments());
      support.replaceComments(List.of());
      assertFalse(support.hasPersistedComments());
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("RawOnly");
      CommentsTable rawOnlyTable =
          (CommentsTable)
              sheet.createRelationship(XSSFRelation.SHEET_COMMENTS, XSSFFactory.getInstance());
      rawOnlyTable.setSheet(sheet);
      rawOnlyTable.newComment(new CellAddress("A1"));

      ExcelSheetCommentRepairSupport support = new ExcelSheetCommentRepairSupport(sheet);

      assertTrue(sheet.getCellComments().isEmpty());
      assertTrue(support.hasPersistedComments());
    }

    CommentsTable commentsTable =
        new CommentsTable() {
          @Override
          public CTComments getCTComments() {
            return CTComments.Factory.newInstance();
          }
        };

    assertEquals(List.of(), invokeRawCommentAddresses(commentsTable));
  }

  @Test
  void insertDeleteAndShiftColumnHelpersRewriteReadComments() {
    List<WorkbookReadResult.CellComment> inserted =
        ExcelSheetCommentRepairSupport.commentsAfterInsertColumns(
            List.of(comment("A2", "Stationary"), comment("C2", "Shifted")), 1, 1);

    assertEquals(List.of(comment("A2", "Stationary"), comment("D2", "Shifted")), inserted);

    List<WorkbookReadResult.CellComment> deleted =
        ExcelSheetCommentRepairSupport.commentsAfterDeleteColumns(
            List.of(comment("A2", "Stationary"), comment("D2", "Shifted")),
            new ExcelColumnSpan(1, 2));

    assertEquals(List.of(comment("A2", "Stationary"), comment("B2", "Shifted")), deleted);

    List<WorkbookReadResult.CellComment> shiftedRight =
        ExcelSheetCommentRepairSupport.commentsAfterShiftColumns(
            List.of(
                comment("A2", "Stationary"),
                comment("B2", "Moving"),
                comment("C2", "Overwritten"),
                comment("E2", "Far")),
            new ExcelColumnSpan(1, 1),
            1);

    assertEquals(
        List.of(comment("A2", "Stationary"), comment("C2", "Moving"), comment("E2", "Far")),
        shiftedRight);

    List<WorkbookReadResult.CellComment> unchanged =
        ExcelSheetCommentRepairSupport.commentsAfterShiftColumns(
            List.of(comment("A2", "Stationary"), comment("B2", "Moving")),
            new ExcelColumnSpan(1, 1),
            0);

    assertEquals(List.of(comment("A2", "Stationary"), comment("B2", "Moving")), unchanged);

    List<WorkbookReadResult.CellComment> shiftedLeft =
        ExcelSheetCommentRepairSupport.commentsAfterShiftColumns(
            List.of(comment("B2", "Moving"), comment("D2", "Stationary")),
            new ExcelColumnSpan(1, 1),
            -1);

    assertEquals(List.of(comment("A2", "Moving"), comment("D2", "Stationary")), shiftedLeft);

    List<WorkbookReadResult.CellComment> shiftedLeftOutsideWindow =
        ExcelSheetCommentRepairSupport.commentsAfterShiftColumns(
            List.of(comment("C2", "Stationary"), comment("F2", "Moving")),
            new ExcelColumnSpan(5, 5),
            -2);

    assertEquals(
        List.of(comment("C2", "Stationary"), comment("D2", "Moving")), shiftedLeftOutsideWindow);
  }

  @Test
  void replaceCommentsSupportsRawPlainAndRichComments() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook.sheet("Ops").setCell("C3", ExcelCellValue.text("Existing"));

      ExcelSheetCommentRepairSupport support =
          new ExcelSheetCommentRepairSupport(workbook.sheet("Ops").xssfSheet());
      ExcelRichTextSnapshot richText =
          new ExcelRichTextSnapshot(
              List.of(new ExcelRichTextRunSnapshot("Raw rich", baseFontWithoutColor())));
      ExcelRichTextSnapshot readbackRichText =
          new ExcelRichTextSnapshot(
              List.of(
                  new ExcelRichTextRunSnapshot(
                      "Raw rich",
                      new ExcelCellFontSnapshot(
                          false,
                          false,
                          "Aptos",
                          new ExcelFontHeight(220),
                          new ExcelColorSnapshot(null, null, 8, null),
                          false,
                          false))));
      ExcelCommentAnchorSnapshot explicitAnchor = new ExcelCommentAnchorSnapshot(2, 2, 6, 7);

      support.replaceComments(
          List.of(
              new ExcelSheetCommentRepairSupport.CommentRewriteSnapshot(
                  "B4", "Plain raw", "", true, null, null),
              new ExcelSheetCommentRepairSupport.CommentRewriteSnapshot(
                  "C3", null, "", false, richText, explicitAnchor)));

      assertEquals(
          Map.of(
              "B4",
                  new ExcelSheetCommentRepairSupport.CommentRewriteSnapshot(
                      "B4",
                      "Plain raw",
                      "",
                      true,
                      null,
                      new ExcelCommentAnchorSnapshot(1, 3, 4, 6)),
              "C3",
                  new ExcelSheetCommentRepairSupport.CommentRewriteSnapshot(
                      "C3", "Raw rich", "", false, readbackRichText, explicitAnchor)),
          rawCommentSnapshots(workbook, "Ops"));
    }
  }

  @Test
  void commentRewriteSnapshotHandlesAnchorlessPoiCommentsAndNormalization() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      CommentsTable commentsTable = new CommentsTable();
      commentsTable.setSheet(sheet);

      XSSFComment anchorless =
          new XSSFComment(commentsTable, commentsTable.newComment(new CellAddress("A1")), null);
      anchorless.setAuthor("GridGrind");
      anchorless.setVisible(false);

      ExcelSheetCommentRepairSupport.CommentRewriteSnapshot snapshot =
          ExcelSheetCommentRepairSupport.CommentRewriteSnapshot.from(
              new CellAddress("A1"), anchorless, workbook);

      assertEquals("A1", snapshot.address());
      assertEquals("", snapshot.text());
      assertEquals("GridGrind", snapshot.author());
      assertFalse(snapshot.visible());
      assertNull(snapshot.runs());
      assertNull(snapshot.anchor());
    }

    ExcelSheetCommentRepairSupport.CommentRewriteSnapshot normalized =
        new ExcelSheetCommentRepairSupport.CommentRewriteSnapshot(
            "B2", null, null, true, null, null);
    ExcelSheetCommentRepairSupport.CommentRewriteSnapshot missingAuthor =
        new ExcelSheetCommentRepairSupport.CommentRewriteSnapshot(
            "B2", "Text", "", true, null, null);
    ExcelSheetCommentRepairSupport.CommentRewriteSnapshot compatible =
        new ExcelSheetCommentRepairSupport.CommentRewriteSnapshot(
            "B2", "Text", "GridGrind", true, null, null);

    assertEquals("", normalized.text());
    assertEquals("", normalized.author());
    assertFalse(normalized.isAuthoringCompatible());
    assertFalse(missingAuthor.isAuthoringCompatible());
    assertTrue(compatible.isAuthoringCompatible());
    assertEquals(
        new ExcelComment("Text", "GridGrind", true, null, null), compatible.toAuthoringComment());
  }

  @Test
  void deleteColumnsRoundTripKeepsShiftedCommentAndCanonicalCommentParts() throws Exception {
    Path workbookPath = Files.createTempFile("gridgrind-comment-delete-columns-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("LL");
      workbook
          .sheet("LL")
          .setComment("E2", new ExcelComment("Note BudgetTotal", "GridGrind", true));
      workbook
          .sheet("LL")
          .setComment("A2", new ExcelComment("Note Report_Value", "GridGrind", true));
      workbook.getOrCreateSheet("LL");
      workbook.sheet("LL").deleteColumns(new ExcelColumnSpan(1, 3));
      workbook.sheet("LL").deleteColumns(new ExcelColumnSpan(0, 0));

      assertVisibleComments(
          workbook, "LL", Map.of("A2", new ExcelComment("Note BudgetTotal", "GridGrind", true)));
      workbook.save(workbookPath);
    }

    assertCanonicalCommentParts(workbookPath, Map.of("A2", "Note BudgetTotal"), Map.of("1:0", 1));

    try (ExcelWorkbook reopened = ExcelWorkbook.open(workbookPath)) {
      assertVisibleComments(
          reopened, "LL", Map.of("A2", new ExcelComment("Note BudgetTotal", "GridGrind", true)));
    }
  }

  @Test
  void insertColumnsRoundTripKeepsShiftedCommentsAndCanonicalCommentParts() throws Exception {
    Path workbookPath = Files.createTempFile("gridgrind-comment-insert-columns-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook.sheet("Ops").setComment("A2", new ExcelComment("Stationary", "GridGrind", true));
      workbook.sheet("Ops").setComment("C2", new ExcelComment("Shifted", "GridGrind", true));
      workbook.sheet("Ops").insertColumns(1, 1);

      assertVisibleComments(
          workbook,
          "Ops",
          Map.of(
              "A2", new ExcelComment("Stationary", "GridGrind", true),
              "D2", new ExcelComment("Shifted", "GridGrind", true)));
      workbook.save(workbookPath);
    }

    assertCanonicalCommentParts(
        workbookPath, Map.of("A2", "Stationary", "D2", "Shifted"), Map.of("1:0", 1, "1:3", 1));

    try (ExcelWorkbook reopened = ExcelWorkbook.open(workbookPath)) {
      assertVisibleComments(
          reopened,
          "Ops",
          Map.of(
              "A2", new ExcelComment("Stationary", "GridGrind", true),
              "D2", new ExcelComment("Shifted", "GridGrind", true)));
    }
  }

  @Test
  void shiftColumnsLeftRoundTripKeepsShiftedCommentAndCanonicalCommentParts() throws Exception {
    Path workbookPath = Files.createTempFile("gridgrind-comment-shift-columns-left-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook.sheet("Ops").setComment("A2", new ExcelComment("Stationary", "GridGrind", true));
      workbook.sheet("Ops").setComment("B2", new ExcelComment("Moving", "GridGrind", true));
      workbook.sheet("Ops").shiftColumns(new ExcelColumnSpan(1, 1), -1);

      assertVisibleComments(
          workbook, "Ops", Map.of("A2", new ExcelComment("Moving", "GridGrind", true)));
      workbook.save(workbookPath);
    }

    assertCanonicalCommentParts(workbookPath, Map.of("A2", "Moving"), Map.of("1:0", 1));

    try (ExcelWorkbook reopened = ExcelWorkbook.open(workbookPath)) {
      assertVisibleComments(
          reopened, "Ops", Map.of("A2", new ExcelComment("Moving", "GridGrind", true)));
    }
  }

  @Test
  void shiftColumnsRightRoundTripKeepsShiftedCommentAndCanonicalCommentParts() throws Exception {
    Path workbookPath = Files.createTempFile("gridgrind-comment-shift-columns-right-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook.sheet("Ops").setComment("A2", new ExcelComment("Moving", "GridGrind", true));
      workbook.sheet("Ops").setComment("B2", new ExcelComment("Stationary", "GridGrind", true));
      workbook.sheet("Ops").shiftColumns(new ExcelColumnSpan(0, 0), 1);

      assertVisibleComments(
          workbook, "Ops", Map.of("B2", new ExcelComment("Moving", "GridGrind", true)));
      workbook.save(workbookPath);
    }

    assertCanonicalCommentParts(workbookPath, Map.of("B2", "Moving"), Map.of("1:1", 1));

    try (ExcelWorkbook reopened = ExcelWorkbook.open(workbookPath)) {
      assertVisibleComments(
          reopened, "Ops", Map.of("B2", new ExcelComment("Moving", "GridGrind", true)));
    }
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static void assertVisibleComments(
      ExcelWorkbook workbook, String sheetName, Map<String, ExcelComment> expected) {
    Map<String, ExcelComment> actual = new java.util.LinkedHashMap<>();
    for (WorkbookReadResult.CellComment comment :
        workbook.sheet(sheetName).comments(new ExcelCellSelection.AllUsedCells())) {
      actual.put(comment.address(), comment.comment().toPlainComment());
    }
    assertEquals(expected, Map.copyOf(actual));
  }

  private static void assertCanonicalCommentParts(
      Path workbookPath, Map<String, String> expectedComments, Map<String, Integer> expectedNotes)
      throws IOException {
    try (ZipFile zipFile = new ZipFile(workbookPath.toFile())) {
      CommentParts commentParts = commentParts(zipFile);
      assertEquals(expectedComments.size(), commentParts.totalComments());
      assertEquals(expectedComments, commentParts.commentsByRef());
      assertEquals(expectedNotes, noteShapeCountsByCell(zipFile));
    }
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static CommentParts commentParts(ZipFile zipFile) throws IOException {
    Map<String, String> comments = new java.util.LinkedHashMap<>();
    int totalComments = 0;
    for (var entry : zipFile.stream().toList()) {
      if (!entry.getName().startsWith("xl/comments")) {
        continue;
      }
      try (InputStream inputStream = zipFile.getInputStream(entry)) {
        Document document = parseXml(inputStream);
        NodeList commentNodes =
            document.getElementsByTagNameNS(
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "comment");
        for (int index = 0; index < commentNodes.getLength(); index++) {
          totalComments++;
          Element comment = (Element) commentNodes.item(index);
          Element textNode =
              (Element)
                  comment
                      .getElementsByTagNameNS(
                          "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "t")
                      .item(0);
          comments.put(comment.getAttribute("ref"), textNode.getTextContent());
        }
      }
    }
    return new CommentParts(Map.copyOf(comments), totalComments);
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static Map<String, Integer> noteShapeCountsByCell(ZipFile zipFile) throws IOException {
    Map<String, Integer> counts = new java.util.LinkedHashMap<>();
    for (var entry : zipFile.stream().toList()) {
      if (!entry.getName().startsWith("xl/drawings/vmlDrawing")) {
        continue;
      }
      try (InputStream inputStream = zipFile.getInputStream(entry)) {
        Document document = parseXml(inputStream);
        NodeList clientDataNodes =
            document.getElementsByTagNameNS("urn:schemas-microsoft-com:office:excel", "ClientData");
        for (int index = 0; index < clientDataNodes.getLength(); index++) {
          Element clientData = (Element) clientDataNodes.item(index);
          if (!"Note".equals(clientData.getAttribute("ObjectType"))) {
            continue;
          }
          String row =
              clientData
                  .getElementsByTagNameNS("urn:schemas-microsoft-com:office:excel", "Row")
                  .item(0)
                  .getTextContent();
          String column =
              clientData
                  .getElementsByTagNameNS("urn:schemas-microsoft-com:office:excel", "Column")
                  .item(0)
                  .getTextContent();
          counts.merge(row + ":" + column, 1, Integer::sum);
        }
      }
    }
    return Map.copyOf(counts);
  }

  private static WorkbookReadResult.CellComment comment(String address, String text) {
    return new WorkbookReadResult.CellComment(
        address, new ExcelCommentSnapshot(text, "GridGrind", true, null, null));
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static Map<String, ExcelSheetCommentRepairSupport.CommentRewriteSnapshot>
      rawCommentSnapshots(ExcelWorkbook workbook, String sheetName) {
    Map<String, ExcelSheetCommentRepairSupport.CommentRewriteSnapshot> snapshots =
        new java.util.LinkedHashMap<>();
    for (var entry : workbook.sheet(sheetName).xssfSheet().getCellComments().entrySet()) {
      snapshots.put(
          entry.getKey().formatAsString(),
          ExcelSheetCommentRepairSupport.CommentRewriteSnapshot.from(
              entry.getKey(), entry.getValue(), workbook.xssfWorkbook()));
    }
    return Map.copyOf(snapshots);
  }

  @SuppressWarnings({"PMD.AvoidAccessibilityAlteration", "unchecked"})
  private static List<CellAddress> invokeRawCommentAddresses(CommentsTable commentsTable)
      throws ReflectiveOperationException {
    Method method =
        ExcelSheetCommentRepairSupport.class.getDeclaredMethod(
            "rawCommentAddresses", CommentsTable.class);
    method.setAccessible(true);
    return (List<CellAddress>) method.invoke(null, commentsTable);
  }

  private static ExcelCellFontSnapshot baseFontWithoutColor() {
    return new ExcelCellFontSnapshot(
        false, false, "Aptos", new ExcelFontHeight(220), null, false, false);
  }

  private static Document parseXml(InputStream inputStream) throws IOException {
    try {
      DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
      factory.setNamespaceAware(true);
      return factory.newDocumentBuilder().parse(inputStream);
    } catch (ParserConfigurationException | SAXException exception) {
      throw new IOException("failed to parse OOXML test fixture", exception);
    }
  }

  private record CommentParts(Map<String, String> commentsByRef, int totalComments) {}
}
