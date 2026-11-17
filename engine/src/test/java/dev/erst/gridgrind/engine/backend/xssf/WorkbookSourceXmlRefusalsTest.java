package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import com.sun.net.httpserver.HttpServer;
import dev.erst.gridgrind.core.plan.ResolvedInspectionStep;
import dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.problem.WorkbookContentIssue;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.engine.backend.eventread.EventWorkbookReader;
import java.io.IOException;
import java.net.InetSocketAddress;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import javax.xml.parsers.ParserConfigurationException;
import org.apache.poi.ooxml.POIXMLException;
import org.apache.xmlbeans.XmlException;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.xml.sax.SAXParseException;

/** Source XML policy is exercised with independent package bytes and an owned access canary. */
class WorkbookSourceXmlRefusalsTest {
  @TempDir Path root;
  private static final String SHEET =
      "<worksheet"
          + " xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row"
          + " r=\"1\"><c r=\"A1\""
          + " t=\"inlineStr\"><is><t>safe</t></is></c></row></sheetData></worksheet>";

  @Test
  void qualifiedNativeWrappersRefuseSafelyAndKeepInitializationAndIoDistinct() throws IOException {
    var syntax = new SAXParseException("synthetic-private-parser-text", null, null, 1, 31);
    var wrapper = new POIXMLException(new XmlException(syntax));
    var refusal =
        assertInstanceOf(
            WorkbookContentRefusalException.class, WorkbookSourceXmlRefusals.translate(wrapper));
    assertEquals(WorkbookContentIssue.XML_SOURCE_REFUSED, refusal.issue());
    assertEquals(31, refusal.location().orElseThrow().column().orElseThrow());
    assertFalse(refusal.getMessage().contains("synthetic-private"));
    assertSame(wrapper, refusal.getCause());
    var io = new IOException("synthetic disk fault");
    assertSame(
        io,
        assertThrows(
            IOException.class, () -> WorkbookSourceXmlRefusals.translate(new XmlException(io))));
    for (Exception unknown :
        List.of(
            new IllegalArgumentException("unknown definition"),
            new IllegalStateException("unknown provider"),
            new XmlException(new ParserConfigurationException("factory unavailable")))) {
      var internal =
          assertInstanceOf(
              UnexpectedNativeContentException.class, WorkbookSourceXmlRefusals.translate(unknown));
      assertSame(unknown, internal.getCause());
    }
  }

  @Test
  void recognizedWrappersPreserveFatalIdentityWithoutTraversingUnknownWrappers() {
    var fatal = new AssertionError("synthetic fatal");
    assertSame(
        fatal,
        assertThrows(
            AssertionError.class,
            () ->
                WorkbookSourceXmlRefusals.translate(new POIXMLException(new XmlException(fatal)))));
    var parse = new SAXParseException("synthetic parser wrapper", null, null, 1, 1);
    parse.initCause(fatal);
    assertSame(
        fatal,
        assertThrows(
            AssertionError.class,
            () -> WorkbookSourceXmlRefusals.translate(new XmlException(parse))));
    var unknown = new IllegalStateException("unknown wrapper", fatal);
    assertInstanceOf(
        UnexpectedNativeContentException.class,
        assertDoesNotThrow(() -> WorkbookSourceXmlRefusals.translate(unknown)));
  }

  @Test
  void nativeFullAndEventParsersNeverAccessTheDeclaredEntityCanary() throws IOException {
    AtomicInteger accesses = new AtomicInteger();
    HttpServer server =
        HttpServer.create(new InetSocketAddress(java.net.InetAddress.getLoopbackAddress(), 0), 0);
    server.createContext(
        "/canary",
        exchange -> {
          accesses.incrementAndGet();
          byte[] bytes = "synthetic canary".getBytes(StandardCharsets.UTF_8);
          exchange.sendResponseHeaders(200, bytes.length);
          try (var response = exchange.getResponseBody()) {
            response.write(bytes);
          }
        });
    server.start();
    try {
      var canary =
          assertDoesNotThrow(() ->
              new java.net.URI(
                  "http",
                  null,
                  server.getAddress().getAddress().getHostAddress(),
                  server.getAddress().getPort(),
                  "/canary",
                  null,
                  null));
      try (var control = canary.toURL().openStream()) {
        assertEquals(
            "synthetic canary", new String(control.readAllBytes(), StandardCharsets.UTF_8));
      }
      assertEquals(1, accesses.get(), "The owned resolver canary is reachable");
      accesses.set(0);
      String doctype =
          "<!DOCTYPE worksheet [<!ENTITY forbidden SYSTEM "
              + '"'
              + canary.toASCIIString()
              + '"'
              + ">]>";
      Path source = workbook("forbidden.xlsx", doctype + SHEET.replace("safe", "&forbidden;"));
      assertThrows(
          WorkbookContentRefusalException.class,
          () -> {
            try (var ignored =
                XssfWorkbooks.open(
                    source, WorkbookTempFileFactory.rooted(root.resolve("scratch")))) {
              fail("DOCTYPE opened");
            }
          });
      var command =
          new ResolvedInspectionStep(
              new StepId("summary"),
              new WorkbookSelector.Current(),
              new WorkbookIntrospectionQuery.GetWorkbookSummary());
      assertThrows(
          WorkbookContentRefusalException.class,
          () -> new EventWorkbookReader().apply(source, List.of(command)));
      assertEquals(0, accesses.get(), "No resolver access, not merely absent text in a response");
    } finally {
      server.stop(0);
    }
  }

  @Test
  void malformedSheetIsKnownInputWhileIndependentValidSourceStillOpens() throws IOException {
    Path malformed = workbook("malformed.xlsx", SHEET.replace("</sheetData>", "</broken>"));
    var failure =
        assertThrows(
            WorkbookContentRefusalException.class,
            () -> {
              try (var ignored =
                  XssfWorkbooks.open(
                      malformed, WorkbookTempFileFactory.rooted(root.resolve("scratch")))) {
                fail("Malformed source opened");
              }
            });
    assertEquals(WorkbookContentIssue.XML_SOURCE_REFUSED, failure.issue());
    Path valid = workbook("valid.xlsx", SHEET);
    try (var workbook =
        XssfWorkbooks.open(valid, WorkbookTempFileFactory.rooted(root.resolve("scratch")))) {
      assertEquals(List.of("Data"), workbook.sheets().sheetNames());
    }
  }

  // Each named ZIP entry needs its own immutable header.
  @SuppressWarnings("PMD.AvoidInstantiatingObjectsInLoops")
  private Path workbook(String name, String sheet) throws IOException {
    Map<String, String> entries =
        Map.of(
            "[Content_Types].xml",
                "<Types"
                    + " xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default"
                    + " Extension=\"rels\""
                    + " ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default"
                    + " Extension=\"xml\" ContentType=\"application/xml\"/><Override"
                    + " PartName=\"/xl/workbook.xml\""
                    + " ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override"
                    + " PartName=\"/xl/worksheets/sheet1.xml\""
                    + " ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>",
            "_rels/.rels",
                "<Relationships"
                    + " xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship"
                    + " Id=\"rId1\""
                    + " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\""
                    + " Target=\"xl/workbook.xml\"/></Relationships>",
            "xl/workbook.xml",
                "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
                    + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet"
                    + " name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>",
            "xl/_rels/workbook.xml.rels",
                "<Relationships"
                    + " xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship"
                    + " Id=\"rId1\""
                    + " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\""
                    + " Target=\"worksheets/sheet1.xml\"/></Relationships>",
            "xl/worksheets/sheet1.xml", sheet);
    Path file = root.resolve(name);
    try (ZipOutputStream zip = new ZipOutputStream(Files.newOutputStream(file))) {
      for (var entry : entries.entrySet()) {
        zip.putNextEntry(new ZipEntry(entry.getKey()));
        zip.write(entry.getValue().getBytes(StandardCharsets.UTF_8));
        zip.closeEntry();
      }
    }
    return file;
  }
}
