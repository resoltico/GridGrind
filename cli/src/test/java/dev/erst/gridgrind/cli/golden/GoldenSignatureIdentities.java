package dev.erst.gridgrind.cli.golden;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.net.URI;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.UUID;
import java.util.zip.ZipFile;
import javax.xml.XMLConstants;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import org.w3c.dom.Element;
import org.xml.sax.SAXException;
import tools.jackson.databind.JsonNode;

/** Certifies generated fixture signature identities against raw source and persisted VML. */
final class GoldenSignatureIdentities {
  private static final String SHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  private static final String REL =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  private static final String PACKAGE_REL =
      "http://schemas.openxmlformats.org/package/2006/relationships";
  private static final String VML = "urn:schemas-microsoft-com:vml";
  private static final String OFFICE = "urn:schemas-microsoft-com:office:office";

  private GoldenSignatureIdentities() {}

  static Capture capture(Path request, Path workspace) throws IOException {
    JsonNode plan = GoldenSnapshots.JSON.readTree(Files.readAllBytes(request));
    boolean relevant =
        plan.path("steps")
            .valueStream()
            .anyMatch(step ->
                Set.of("SET_SIGNATURE_LINE", "COPY_SHEET")
                    .contains(step.path("action").path("type").asString("")));
    if (!relevant) return new Capture(plan, Optional.empty(), "", Map.of());
    if ("EXISTING".equals(plan.path("source").path("type").asString())) {
      Path source = workspace.resolve(plan.path("source").path("path").asString());
      return new Capture(
          plan,
          Optional.of(source),
          GoldenSignaturePreviews.sha256(Files.readAllBytes(source)),
          read(source));
    }
    return new Capture(plan, Optional.empty(), "", Map.of());
  }

  static Set<String> certify(Capture before, Path workspace, JsonNode execution)
      throws IOException {
    JsonNode plan = before.plan();
    boolean relevant =
        plan.path("steps")
            .valueStream()
            .anyMatch(step ->
                Set.of("SET_SIGNATURE_LINE", "COPY_SHEET")
                    .contains(step.path("action").path("type").asString("")));
    if (!relevant) return Set.of();
    String persistence = plan.path("persistence").path("type").asString();
    assertNotEquals(
        "NONE", persistence, "Generated signature identities require a persisted golden artifact");
    Path output =
        workspace.resolve(
            "SAVE_AS".equals(persistence)
                ? plan.path("persistence").path("path").asString()
                : plan.path("source").path("path").asString());
    Map<String, SheetIdentity> actual = read(output);
    if (before.source().isPresent() && "SAVE_AS".equals(persistence)) {
      assertEquals(
          before.sourceHash(),
          GoldenSignaturePreviews.sha256(Files.readAllBytes(before.source().orElseThrow())),
          "Save-as must leave the original signature source package unchanged");
    }
    Set<String> generated = new LinkedHashSet<>();
    Set<String> originalIds = new LinkedHashSet<>();
    before.sheets().values().forEach(sheet -> originalIds.addAll(sheet.ids().values()));
    for (JsonNode step : plan.path("steps")) {
      JsonNode action = step.path("action");
      String kind = action.path("type").asString("");
      if ("SET_SIGNATURE_LINE".equals(kind)) {
        assertEquals(
            "SHEET_BY_NAME",
            step.path("target").path("type").asString(),
            "Golden signature target must identify one sheet");
        String sheet = step.path("target").path("name").asString();
        String name = action.path("signatureLine").path("name").asString();
        String id = identity(actual, sheet).ids().get(name);
        assertNotNull(id, "Authored signature must exist in persisted VML: " + name);
        requireGuid(id);
        assertFalse(
            originalIds.contains(id),
            "Recreated signatures must not reuse an original source identity");
        assertTrue(generated.add(id), "Generated identities must not collide");
      } else if ("COPY_SHEET".equals(kind)) {
        String sourceName = step.path("target").path("name").asString();
        String targetName = action.path("newSheetName").asString();
        SheetIdentity source = identity(actual, sourceName);
        SheetIdentity copy = identity(actual, targetName);
        SheetIdentity original = before.sheets().get(sourceName);
        if (original != null)
          assertEquals(
              original.ids(), source.ids(), "Copy must preserve source signature identities");
        certifyCopy(source, copy, originalIds, generated);
      }
    }
    verifyReports(execution, "", actual);
    return Set.copyOf(generated);
  }

  static void certifyCopy(
      SheetIdentity source, SheetIdentity copy, Set<String> originalIds, Set<String> generated) {
    assertEquals(
        source.ids().keySet(), copy.ids().keySet(), "Copy must preserve every signature name");
    if (!source.ids().isEmpty())
      assertNotEquals(
          source.vmlPart(), copy.vmlPart(), "Copied signatures require independent persisted VML");
    for (String name : source.ids().keySet()) {
      String id = copy.ids().get(name);
      requireGuid(source.ids().get(name));
      requireGuid(id);
      assertNotEquals(
          source.ids().get(name), id, "Copied signature must have its own native identity");
      assertFalse(
          originalIds.contains(id), "Copied identity must not collide with any original signature");
      assertTrue(generated.add(id), "Generated identities must not collide");
    }
  }

  private static SheetIdentity identity(Map<String, SheetIdentity> sheets, String name) {
    return sheets.getOrDefault(name, new SheetIdentity("", Map.of()));
  }

  private static void verifyReports(
      JsonNode node, String inheritedSheet, Map<String, SheetIdentity> nativeIds) {
    if (node.isObject()) {
      String sheet = node.has("sheetName") ? node.path("sheetName").asString() : inheritedSheet;
      if ("SIGNATURE_LINE".equals(node.path("type").asString(""))) {
        String id = node.path("setup").path("setupId").asString();
        assertEquals(
            identity(nativeIds, sheet).ids().get(node.path("name").asString()),
            id,
            "Reported signature identity must match persisted VML on sheet " + sheet);
      }
      node.properties().forEach(entry -> verifyReports(entry.getValue(), sheet, nativeIds));
    } else if (node.isArray()) {
      node.forEach(value -> verifyReports(value, inheritedSheet, nativeIds));
    }
  }

  static void requireGuid(String id) {
    assertTrue(
        id.startsWith("{") && id.endsWith("}"), "Native signature identity must use GUID braces");
    String body = id.substring(1, id.length() - 1);
    assertEquals(
        UUID.fromString(body).toString(),
        body.toLowerCase(Locale.ROOT),
        "Native signature GUID must be canonical");
  }

  @SuppressWarnings({
    "PMD.UseConcurrentHashMap", // One fixture inspection owns these ordered maps.
    "PMD.AvoidInstantiatingObjectsInLoops" // Each sheet requires its own immutable native identity
    // fact.
  })
  private static Map<String, SheetIdentity> read(Path path) throws IOException {
    Map<String, SheetIdentity> result = new LinkedHashMap<>();
    try (var zip = new ZipFile(path.toFile())) {
      Element workbook = xml(zip, "xl/workbook.xml");
      Map<String, String> workbookRels = relationships(zip, "xl/workbook.xml");
      var sheets = workbook.getElementsByTagNameNS(SHEET, "sheet");
      for (int index = 0; index < sheets.getLength(); index++) {
        Element sheet = (Element) sheets.item(index);
        String part = target("xl/workbook.xml", workbookRels.get(sheet.getAttributeNS(REL, "id")));
        var legacy = xml(zip, part).getElementsByTagNameNS(SHEET, "legacyDrawing");
        if (legacy.getLength() == 0) continue;
        assertEquals(1, legacy.getLength(), "Golden worksheet has one legacy drawing relationship");
        String vmlPart =
            target(
                part,
                relationships(zip, part).get(((Element) legacy.item(0)).getAttributeNS(REL, "id")));
        Map<String, String> ids = signatureIds(xml(zip, vmlPart));
        result.put(sheet.getAttribute("name"), new SheetIdentity(vmlPart, ids));
      }
    }
    return Map.copyOf(result);
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap") // One VML fixture owns its identity map.
  private static Map<String, String> signatureIds(Element vml) {
    Map<String, String> result = new LinkedHashMap<>();
    var shapes = vml.getElementsByTagNameNS(VML, "shape");
    for (int index = 0; index < shapes.getLength(); index++) {
      Element shape = (Element) shapes.item(index);
      var lines = shape.getElementsByTagNameNS(OFFICE, "signatureline");
      if (lines.getLength() == 0) continue;
      assertEquals(1, lines.getLength(), "Golden signature shape has one native setup declaration");
      String name = shape.getAttribute("alt");
      assertFalse(name.isBlank(), "Golden signature has a stored object name");
      String id = ((Element) lines.item(0)).getAttribute("id");
      requireGuid(id);
      assertNull(
          result.putIfAbsent(name, id), "Golden signature names are unique within each sheet");
    }
    return Map.copyOf(result);
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap") // Package relationships are local to one fixture.
  private static Map<String, String> relationships(ZipFile zip, String part) throws IOException {
    int slash = part.lastIndexOf('/');
    String rels = part.substring(0, slash + 1) + "_rels/" + part.substring(slash + 1) + ".rels";
    Map<String, String> result = new LinkedHashMap<>();
    if (zip.getEntry(rels) == null) return result;
    var values = xml(zip, rels).getElementsByTagNameNS(PACKAGE_REL, "Relationship");
    for (int index = 0; index < values.getLength(); index++) {
      Element relation = (Element) values.item(index);
      if (!"External".equals(relation.getAttribute("TargetMode")))
        result.put(relation.getAttribute("Id"), relation.getAttribute("Target"));
    }
    return result;
  }

  private static String target(String part, String target) {
    assertNotNull(target, "Golden package relationship must resolve internally");
    String resolved = URI.create("/" + part).resolve(target).normalize().getPath();
    assertTrue(resolved.startsWith("/xl/"), "Golden worksheet relationship remains inside xl/");
    return resolved.substring(1);
  }

  private static Element xml(ZipFile zip, String part) throws IOException {
    var entry = zip.getEntry(part);
    assertNotNull(entry, "Missing golden package part: " + part);
    var factory = DocumentBuilderFactory.newDefaultInstance();
    factory.setNamespaceAware(true);
    factory.setXIncludeAware(false);
    factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_DTD, "");
    factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_SCHEMA, "");
    try (var input = zip.getInputStream(entry)) {
      factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
      return factory.newDocumentBuilder().parse(input).getDocumentElement();
    } catch (ParserConfigurationException | SAXException failure) {
      throw new IOException("Cannot inspect golden package XML: " + part, failure);
    }
  }

  /** Input identities captured before the actual product workflow runs. */
  record Capture(
      JsonNode plan, Optional<Path> source, String sourceHash, Map<String, SheetIdentity> sheets) {}

  /** Persisted native VML ownership and exact setup IDs, independent of product reports. */
  record SheetIdentity(String vmlPart, Map<String, String> ids) {}
}
