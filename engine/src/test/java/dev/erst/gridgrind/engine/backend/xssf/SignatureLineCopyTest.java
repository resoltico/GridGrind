package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.microsoft.schemas.vml.CTGroup;
import com.microsoft.schemas.vml.CTShape;
import com.microsoft.schemas.vml.CTShapetype;
import dev.erst.gridgrind.core.problem.WorkbookContentIssue;
import dev.erst.gridgrind.core.vocabulary.SheetCopyPosition;
import dev.erst.gridgrind.engine.drawing.DrawingAdapter;
import dev.erst.gridgrind.engine.drawing.SignatureAdapter;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.UUID;
import javax.xml.XMLConstants;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlCursor;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STTrueFalse;
import org.w3c.dom.Node;

/**
 * Sheet copies preserve native signature state rather than reconstructing its incomplete report.
 */
class SignatureLineCopyTest {
  @TempDir Path root;

  @Test
  void copyReopenDeleteAndLaterCommentKeepIndependentNativeSignatureOwners() throws IOException {
    Path saved = root.resolve("copy.xlsx");
    try (var workbook = source()) {
      XSSFSheet original = workbook.xssfWorkbook().getSheetAt(0);
      augmentNativeSource(original);
      String before = original.getVMLDrawing(false).getDocument().xmlText();
      var sourceShapes = byName(original);
      workbook
          .sheets()
          .copySheet(original.getSheetName(), "Copied", new SheetCopyPosition.AppendAtEnd());
      XSSFSheet copied = workbook.xssfWorkbook().getSheet("Copied");
      assertEquals(before, original.getVMLDrawing(false).getDocument().xmlText());
      assertNotEquals(
          original.getVMLDrawing(false).getPackagePart().getPartName(),
          copied.getVMLDrawing(false).getPackagePart().getPartName());
      assertNativeStateAndPreviewOwners(original, copied, sourceShapes);
      assertTrue(copied.getCTWorksheet().isSetLegacyDrawing());
      assertEquals(1, copied.getCellComments().size());
      assertEquals(
          List.of(true, false),
          new SignatureAdapter()
              .signatureLines(copied).stream()
                  .map(line -> line.setup().orElseThrow().allowComments().orElseThrow())
                  .toList());
      write(workbook, saved);
    }
    try (var workbook = open(saved)) {
      XSSFSheet original = workbook.xssfWorkbook().getSheetAt(0);
      XSSFSheet copied = workbook.xssfWorkbook().getSheet("Copied");
      assertNativeStateAndPreviewOwners(original, copied, byName(original));
      var later =
          copied
              .createDrawingPatriarch()
              .createCellComment(new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 4, 5));
      later.setAuthor("Synthetic later comment");
      later.setString(
          workbook
              .xssfWorkbook()
              .getCreationHelper()
              .createRichTextString("Native comment type remains intact"));
      copied.createRow(2).createCell(2).setCellComment(later);
      assertCommentTypes(copied);
      assertTrue(new SignatureAdapter().deleteIfPresent(copied, "DeputyLine"));
      assertEquals(2, byName(original).size());
      assertEquals(1, byName(copied).size());
      assertPreviewPresent(copied, byName(copied).get("ReviewerLine"));
      write(workbook, root.resolve("after-delete.xlsx"));
    }
    try (var workbook = open(root.resolve("after-delete.xlsx"))) {
      assertEquals(2, byName(workbook.xssfWorkbook().getSheetAt(0)).size());
      var copied = workbook.xssfWorkbook().getSheet("Copied");
      assertEquals(1, byName(copied).size());
      assertEquals(2, copied.getCellComments().size());
      assertCommentTypes(copied);
      assertPreviewPresent(copied, byName(copied).get("ReviewerLine"));
    }
  }

  @Test
  void unsupportedNativeDependencyRefusesBeforeCloneWithoutSourceMutation()
      throws IOException, InvalidFormatException {
    try (var workbook = source()) {
      XSSFSheet original = workbook.xssfWorkbook().getSheetAt(0);
      byName(original).get("ReviewerLine").setType("#missing-native-type");
      String before = original.getVMLDrawing(false).getDocument().xmlText();
      int parts = workbook.xssfWorkbook().getPackage().getParts().size();
      var refusal =
          assertThrows(
              WorkbookContentRefusalException.class,
              () ->
                  workbook
                      .sheets()
                      .copySheet(
                          original.getSheetName(),
                          "NeverCreated",
                          new SheetCopyPosition.AppendAtEnd()));
      assertEquals(WorkbookContentIssue.UNSUPPORTED_DEPENDENCY, refusal.issue());
      assertEquals(1, workbook.xssfWorkbook().getNumberOfSheets());
      assertEquals(parts, workbook.xssfWorkbook().getPackage().getParts().size());
      assertEquals(before, original.getVMLDrawing(false).getDocument().xmlText());
    }
  }

  @Test
  void generatedSetupIdentitySkipsAnOccupiedGuid() {
    UUID occupied = UUID.fromString("11111111-2222-3333-4444-555555555555");
    UUID available = UUID.fromString("66666666-7777-8888-9999-aaaaaaaaaaaa");
    Set<String> ids = new HashSet<>(Set.of(occupied.toString().toUpperCase(Locale.ROOT)));
    var candidates = List.of(occupied, available).iterator();
    assertEquals(
        "{66666666-7777-8888-9999-AAAAAAAAAAAA}",
        SignatureLineIdentities.nextSetupId(ids, candidates::next));
    assertEquals(2, ids.size());
  }

  private XssfWorkbook source() throws IOException {
    try (var input =
        Objects.requireNonNull(
            getClass().getResourceAsStream("/workbooks/signature-line-source.xlsx"))) {
      return XssfWorkbooks.wrap(new XSSFWorkbook(input));
    }
  }

  private static XssfWorkbook open(Path path) throws IOException {
    try (var input = Files.newInputStream(path)) {
      return XssfWorkbooks.wrap(new XSSFWorkbook(input));
    }
  }

  private static void write(XssfWorkbook workbook, Path path) throws IOException {
    try (var output = Files.newOutputStream(path)) {
      workbook.xssfWorkbook().write(output);
    }
  }

  private static void augmentNativeSource(XSSFSheet source) {
    var vml = source.getVMLDrawing(false);
    var first = SignatureAdapter.signatureShapes(vml).getFirst();
    first.setId("_gridgrind_signature_1");
    first.setSpid("_gridgrind_signature_1");
    first.setType("#_x0000_t202");
    var firstSetup = first.getSignaturelineArray(0);
    firstSetup.setShowsigndate(STTrueFalse.TRUE);
    firstSetup.setAddlxml("<synthetic-native-metadata/>");
    firstSetup.setSigprovurl("https://example.com/signature-provider");
    var second = (CTShape) first.copy();
    second.setId("_gridgrind_signature_2");
    second.setSpid("_gridgrind_signature_2");
    second.setAlt("DeputyLine");
    second.getImagedataArray(0).setTitle("DeputyLine");
    second.getSignaturelineArray(0).setId("{11111111-2222-3333-4444-555555555555}");
    second.getSignaturelineArray(0).setSuggestedsigner2("Synthetic Deputy");
    second.getSignaturelineArray(0).setAllowcomments(STTrueFalse.FALSE);
    CTGroup group = CTGroup.Factory.newInstance();
    group.addNewShape().set(second);
    try (XmlCursor from = group.newCursor();
        XmlCursor to = vml.getDocument().getXml().newCursor()) {
      from.toFirstChild();
      to.toEndToken();
      from.copyXml(to);
    }
    try (XmlCursor cursor = vml.getDocument().getXml().newCursor()) {
      for (boolean child = cursor.toFirstChild(); child; child = cursor.toNextSibling()) {
        if (cursor.getObject() instanceof CTShapetype type) type.setCoordorigin("17,31");
      }
    }
    var comment =
        source
            .createDrawingPatriarch()
            .createCellComment(new XSSFClientAnchor(10, 20, 30, 40, 1, 1, 3, 4));
    comment.setAuthor("Synthetic source comment");
    comment.setString(
        source
            .getWorkbook()
            .getCreationHelper()
            .createRichTextString("Comment coexists with signature VML"));
    source.createRow(1).createCell(1).setCellComment(comment);
  }

  // This single-threaded fixture index retains native shape order for exact expected-owner
  // comparisons.
  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static Map<String, CTShape> byName(XSSFSheet sheet) {
    Map<String, CTShape> result = new LinkedHashMap<>();
    for (var shape : SignatureAdapter.signatureShapes(sheet.getVMLDrawing(false)))
      result.put(shape.getAlt(), shape);
    return result;
  }

  private static void assertNativeStateAndPreviewOwners(
      XSSFSheet original, XSSFSheet copied, Map<String, CTShape> expected) {
    var actual = byName(copied);
    assertEquals(expected.keySet(), actual.keySet());
    for (String name : expected.keySet()) {
      CTShape source = expected.get(name), target = actual.get(name);
      assertNotEquals(
          source.getSignaturelineArray(0).getId(), target.getSignaturelineArray(0).getId());
      assertNotEquals(source.getId(), target.getId());
      assertEquals(target.getId(), target.getSpid());
      assertEquals(semanticShape(source), semanticShape(target));
      var sourcePart = assertPreviewPresent(original, source);
      var targetPart = assertPreviewPresent(copied, target);
      assertNotEquals(sourcePart.getPartName(), targetPart.getPartName());
      assertArrayEquals(DrawingAdapter.partBytes(sourcePart), DrawingAdapter.partBytes(targetPart));
      CTShapetype sourceType = requiredType(original, source.getType().substring(1));
      CTShapetype targetType = requiredType(copied, target.getType().substring(1));
      var normalizedSource = (CTShapetype) sourceType.copy();
      normalizedSource.setId("NORMALIZED_TYPE");
      var normalizedTarget = (CTShapetype) targetType.copy();
      normalizedTarget.setId("NORMALIZED_TYPE");
      assertEquals(
          semanticXml(normalizedSource.getDomNode()), semanticXml(normalizedTarget.getDomNode()));
    }
    var previewNames =
        actual.values().stream()
            .map(shape -> assertPreviewPresent(copied, shape).getPartName())
            .toList();
    assertEquals(previewNames.size(), Set.copyOf(previewNames).size());
  }

  private static PackagePart assertPreviewPresent(XSSFSheet sheet, CTShape shape) {
    var vml = sheet.getVMLDrawing(false);
    String relation = shape.getImagedataArray(0).getRelid();
    assertNotNull(vml.getPackagePart().getRelationship(relation));
    return DrawingAdapter.relatedInternalPart(vml.getPackagePart(), relation).orElseThrow();
  }

  private static CTShapetype requiredType(XSSFSheet sheet, String id) {
    try (XmlCursor cursor = sheet.getVMLDrawing(false).getDocument().getXml().newCursor()) {
      for (boolean child = cursor.toFirstChild(); child; child = cursor.toNextSibling()) {
        if (cursor.getObject() instanceof CTShapetype type && id.equals(type.getId())) return type;
      }
    }
    throw new AssertionError("Missing declared native type");
  }

  private static void assertCommentTypes(XSSFSheet sheet) {
    try (XmlCursor cursor = sheet.getVMLDrawing(false).getDocument().getXml().newCursor()) {
      for (boolean child = cursor.toFirstChild(); child; child = cursor.toNextSibling()) {
        if (cursor.getObject() instanceof CTShape shape && shape.sizeOfSignaturelineArray() == 0)
          assertEquals("#_x0000_t202", shape.getType());
      }
    }
  }

  private static String semanticShape(CTShape input) {
    CTShape normalized = (CTShape) input.copy();
    normalized.setId("NORMALIZED_SHAPE");
    normalized.setSpid("NORMALIZED_SHAPE");
    normalized.setType("#NORMALIZED_TYPE");
    normalized.getSignaturelineArray(0).setId("{00000000-0000-0000-0000-000000000000}");
    for (var image : normalized.getImagedataList()) image.setRelid("rIdNormalized");
    return semanticXml(normalized.getDomNode());
  }

  private static String semanticXml(Node node) {
    StringBuilder value = new StringBuilder();
    if (node.getNodeType() == Node.ELEMENT_NODE) {
      value.append('{').append(node.getNamespaceURI()).append('}').append(node.getLocalName());
      List<String> attributes = new ArrayList<>();
      var nativeAttributes = node.getAttributes();
      for (int index = 0; index < nativeAttributes.getLength(); index++) {
        Node attribute = nativeAttributes.item(index);
        if (!XMLConstants.XMLNS_ATTRIBUTE_NS_URI.equals(attribute.getNamespaceURI()))
          attributes.add(
              "{"
                  + attribute.getNamespaceURI()
                  + "}"
                  + attribute.getLocalName()
                  + "="
                  + attribute.getNodeValue());
      }
      attributes.sort(Comparator.naturalOrder());
      value.append(attributes);
    } else if (node.getNodeType() == Node.TEXT_NODE) value.append(node.getNodeValue());
    var children = node.getChildNodes();
    for (int index = 0; index < children.getLength(); index++)
      value.append(semanticXml(children.item(index)));
    return value.toString();
  }
}
