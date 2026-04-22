package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.function.UnaryOperator;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CommentsDocument;

/** Regression tests for persisted OOXML comment-package invariants. */
class XlsxCommentPackageInvariantSupportTest {
  @Test
  void requireCanonicalCommentPackageStateAcceptsCanonicalWorkbook() throws IOException {
    Path workbookPath = workbookWithComment("gridgrind-comment-package-canonical-");

    assertDoesNotThrow(
        () -> XlsxCommentPackageInvariantSupport.requireCanonicalCommentPackageState(workbookPath));
  }

  @Test
  void requireCanonicalCommentPackageStateRejectsDuplicateCommentRefs() throws IOException {
    Path workbookPath = workbookWithComment("gridgrind-comment-package-duplicate-");
    Path mutatedWorkbook =
        rewriteEntry(
            workbookPath,
            "xl/comments",
            XlsxCommentPackageInvariantSupportTest::duplicateCommentRef);

    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                XlsxCommentPackageInvariantSupport.requireCanonicalCommentPackageState(
                    mutatedWorkbook));
    assertTrue(failure.getMessage().contains("comment refs must be unique"));
  }

  private static Path workbookWithComment(String prefix) throws IOException {
    Path workbookPath = Files.createTempFile(prefix, ".xlsx");
    Files.deleteIfExists(workbookPath);
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.sheet("Budget").setComment("A1", new ExcelComment("Review", "GridGrind", true));
      workbook.save(workbookPath);
    }
    return workbookPath;
  }

  private static Path rewriteEntry(
      Path sourceWorkbook, String entryPrefix, UnaryOperator<byte[]> transformer)
      throws IOException {
    Path rewrittenWorkbook = Files.createTempFile("gridgrind-comment-package-mutated-", ".xlsx");
    try (ZipInputStream inputStream = new ZipInputStream(Files.newInputStream(sourceWorkbook));
        ZipOutputStream outputStream =
            new ZipOutputStream(Files.newOutputStream(rewrittenWorkbook))) {
      boolean rewritten = false;
      while (true) {
        ZipEntry entry = inputStream.getNextEntry();
        if (entry == null) {
          break;
        }
        byte[] bytes = inputStream.readAllBytes();
        outputStream.putNextEntry(copyOf(entry));
        if (entry.getName().startsWith(entryPrefix)) {
          outputStream.write(transformer.apply(bytes));
          rewritten = true;
        } else {
          outputStream.write(bytes);
        }
        outputStream.closeEntry();
      }
      if (!rewritten) {
        throw new IllegalStateException("missing OOXML entry with prefix " + entryPrefix);
      }
    }
    return rewrittenWorkbook;
  }

  private static byte[] duplicateCommentRef(byte[] bytes) {
    try {
      CommentsDocument commentsDocument =
          CommentsDocument.Factory.parse(new ByteArrayInputStream(bytes));
      var commentList = commentsDocument.getComments().getCommentList();
      commentList.addNewComment().set(commentList.getCommentArray(0));
      ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
      commentsDocument.save(outputStream);
      return outputStream.toByteArray();
    } catch (IOException | org.apache.xmlbeans.XmlException exception) {
      throw new IllegalStateException("failed to duplicate comment ref for test", exception);
    }
  }

  private static ZipEntry copyOf(ZipEntry entry) {
    return new ZipEntry(entry.getName());
  }
}
