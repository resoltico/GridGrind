package dev.erst.gridgrind.jazzer.support;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.Enumeration;
import java.util.HashSet;
import java.util.Objects;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CommentsDocument;

/** Verifies canonical persisted OOXML comment state inside saved `.xlsx` packages. */
final class XlsxCommentPackageInvariantSupport {
  private XlsxCommentPackageInvariantSupport() {}

  static void requireCanonicalCommentPackageState(Path workbookPath) throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");
    try (ZipFile zipFile = new ZipFile(workbookPath.toFile())) {
      Enumeration<? extends ZipEntry> entries = zipFile.entries();
      while (entries.hasMoreElements()) {
        ZipEntry entry = entries.nextElement();
        if (entry.getName().startsWith("xl/comments")) {
          requireCanonicalCommentEntry(zipFile, entry);
        }
      }
    }
  }

  private static void requireCanonicalCommentEntry(ZipFile zipFile, ZipEntry entry)
      throws IOException {
    CommentsDocument commentsDocument = parseComments(zipFile, entry);
    if (commentsDocument.getComments() == null
        || commentsDocument.getComments().getCommentList() == null) {
      return;
    }
    Set<String> refs = new HashSet<>();
    for (var comment : commentsDocument.getComments().getCommentList().getCommentArray()) {
      if (!refs.add(comment.getRef())) {
        throw new IllegalStateException(
            "comment refs must be unique in " + entry.getName() + ": " + comment.getRef());
      }
    }
  }

  private static CommentsDocument parseComments(ZipFile zipFile, ZipEntry entry)
      throws IOException {
    try (InputStream inputStream = zipFile.getInputStream(entry)) {
      return CommentsDocument.Factory.parse(inputStream);
    } catch (XmlException exception) {
      throw new IllegalStateException(
          "comment OOXML must parse successfully in " + entry.getName(), exception);
    }
  }
}
