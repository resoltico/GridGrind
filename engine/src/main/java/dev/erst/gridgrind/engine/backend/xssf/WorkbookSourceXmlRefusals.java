package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.core.problem.WorkbookContentIssue;
import dev.erst.gridgrind.core.problem.WorkbookContentLocation;
import java.io.IOException;
import java.util.Collections;
import java.util.IdentityHashMap;
import java.util.Optional;
import java.util.Set;
import org.apache.poi.ooxml.POIXMLException;
import org.apache.xmlbeans.XmlException;
import org.xml.sax.SAXParseException;

/** Classifies only qualified XML failures at an actual workbook source-load boundary. */
public final class WorkbookSourceXmlRefusals {
  private WorkbookSourceXmlRefusals() {}

  /**
   * Preserves source I/O and known XML refusals while keeping unknown provider failures internal.
   * The caller must invoke this only for native source loading, not arbitrary execution failures.
   */
  public static RuntimeException translate(Exception failure) throws IOException {
    if (failure instanceof WorkbookContentRefusalException refusal) return refusal;
    if (failure instanceof UnexpectedNativeContentException unexpected) return unexpected;
    Throwable terminal = qualifiedTerminal(failure);
    if (terminal instanceof Error fatal) throw fatal;
    if (terminal instanceof IOException io) throw io;
    if (terminal instanceof SAXParseException parse) {
      if (parse.getCause() instanceof Error fatal) throw fatal;
      if (parse.getCause() instanceof IOException io) throw io;
      if (parse.getCause() == null) {
        Optional<Integer> line = positive(parse.getLineNumber());
        Optional<Integer> column = positive(parse.getColumnNumber());
        Optional<WorkbookContentLocation> location =
            line.isPresent() || column.isPresent()
                ? Optional.of(
                    new WorkbookContentLocation(
                        Optional.empty(), Optional.empty(), Optional.empty(), line, column))
                : Optional.empty();
        return new WorkbookContentRefusalException(
            WorkbookContentIssue.XML_SOURCE_REFUSED, location, failure);
      }
    }
    return new UnexpectedNativeContentException(failure);
  }

  private static Throwable qualifiedTerminal(Exception failure) {
    Set<Throwable> seen = Collections.newSetFromMap(new IdentityHashMap<>());
    Throwable current = failure;
    while (current instanceof POIXMLException || current instanceof XmlException) {
      if (!seen.add(current) || current.getCause() == null) return current;
      current = current.getCause();
    }
    return current;
  }

  private static Optional<Integer> positive(int value) {
    return value > 0 ? Optional.of(value) : Optional.empty();
  }
}
