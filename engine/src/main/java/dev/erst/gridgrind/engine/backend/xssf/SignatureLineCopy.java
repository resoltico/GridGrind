package dev.erst.gridgrind.engine.backend.xssf;

import com.microsoft.schemas.vml.CTGroup;
import com.microsoft.schemas.vml.CTShape;
import com.microsoft.schemas.vml.CTShapetype;
import dev.erst.gridgrind.core.problem.WorkbookContentIssue;
import dev.erst.gridgrind.core.problem.WorkbookContentLocation;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.engine.drawing.DrawingAdapter;
import dev.erst.gridgrind.engine.drawing.SignatureAdapter;
import dev.erst.gridgrind.engine.poi.ImageCatalog;
import java.net.URI;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import javax.xml.namespace.QName;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFVMLDrawing;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.impl.values.XmlValueOutOfRangeException;
import org.jspecify.annotations.Nullable;

/** Captures and restores native signature VML without rebuilding visual content from reports. */
final class SignatureLineCopy {
  private static final String OFFICE = "urn:schemas-microsoft-com:office:office";
  private static final String RELATIONSHIPS =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  private static final String COMMENT_TYPE = "_x0000_t202";

  private SignatureLineCopy() {}

  static Capture capture(XSSFSheet source) {
    try {
      return captureNative(source);
    } catch (WorkbookContentRefusalException | UnexpectedNativeContentException exception) {
      throw exception;
    } catch (XmlValueOutOfRangeException exception) {
      throw refused(source, WorkbookContentIssue.INVALID_NATIVE_VALUE, "signatureLine", exception);
    } catch (RuntimeException exception) {
      throw new UnexpectedNativeContentException(exception);
    }
  }

  // Capture maps belong to this invocation and preserve native dependency order.
  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static Capture captureNative(XSSFSheet source) {
    XSSFVMLDrawing vml = source.getVMLDrawing(false);
    if (vml == null) {
      if (source.getCTWorksheet().isSetLegacyDrawing()) {
        throw refused(
            source, WorkbookContentIssue.UNSUPPORTED_DEPENDENCY, "worksheet.legacyDrawing", null);
      }
      return new Capture(List.of(), List.of(), Optional.empty());
    }
    List<CTShape> shapes = SignatureAdapter.signatureShapes(vml);
    if (signatureCount(vml) != shapes.size()
        || shapes.stream().anyMatch(shape -> shape.sizeOfSignaturelineArray() != 1)) {
      throw refused(
          source, WorkbookContentIssue.UNSUPPORTED_PRESERVATION, "signatureLine.carrier", null);
    }
    if (shapes.isEmpty()) return new Capture(List.of(), List.of(), Optional.empty());
    Map<String, CTShapetype> sourceTypes = shapeTypes(vml, source);
    Map<String, TypeCapture> types = new LinkedHashMap<>();
    List<LineCapture> lines = new ArrayList<>();
    for (CTShape shape : shapes) {
      captureType(source, vml, shape, sourceTypes, types);
      lines.add(captureLine(source, vml, shape));
    }
    Set<String> copiedIds = SignatureLineIdentities.sourceIds(source, types.keySet(), shapes);
    for (LineCapture line : lines) requireLocalReferences(source, line.shape(), copiedIds);
    for (TypeCapture type : types.values())
      requireLocalReferences(source, type.definition(), copiedIds);
    return new Capture(
        lines,
        List.copyOf(types.values()),
        Optional.of(vml.getPackagePart().getPartName().getName()));
  }

  private static void captureType(
      XSSFSheet source,
      XSSFVMLDrawing vml,
      CTShape shape,
      Map<String, CTShapetype> sourceTypes,
      Map<String, TypeCapture> types) {
    String type = shape.getType();
    if (type != null && type.startsWith("#")) {
      String id = type.substring(1);
      CTShapetype definition = sourceTypes.get(id);
      if (definition == null)
        throw refused(
            source, WorkbookContentIssue.UNSUPPORTED_DEPENDENCY, "signatureLine.type", null);
      if (!types.containsKey(id)) {
        types.put(
            id,
            new TypeCapture(
                (CTShapetype) definition.copy(), relationships(source, vml, definition)));
      }
    }
  }

  private static LineCapture captureLine(XSSFSheet source, XSSFVMLDrawing vml, CTShape shape) {
    return new LineCapture((CTShape) shape.copy(), relationships(source, vml, shape));
  }

  static void restore(Capture capture, XSSFSheet target) {
    if (capture.lines().isEmpty()) return;
    try {
      restoreNative(capture, target);
    } catch (WorkbookContentRefusalException | UnexpectedNativeContentException exception) {
      throw exception;
    } catch (XmlValueOutOfRangeException exception) {
      throw refused(target, WorkbookContentIssue.INVALID_NATIVE_VALUE, "signatureLine", exception);
    } catch (RuntimeException exception) {
      throw new UnexpectedNativeContentException(exception);
    }
  }

  private static void restoreNative(Capture capture, XSSFSheet target) {
    XSSFVMLDrawing vml = target.getVMLDrawing(true);
    if (capture
        .sourcePartName()
        .filter(vml.getPackagePart().getPartName().getName()::equals)
        .isPresent()) {
      throw new IllegalStateException("Copied signatures require independently owned VML");
    }
    Set<String> setupIds = SignatureLineIdentities.setupIds(target);
    Map<String, String> identityMap =
        SignatureLineIdentities.mapping(
            vml,
            capture.types().stream().map(TypeCapture::definition).toList(),
            capture.lines().stream().map(LineCapture::shape).toList());
    for (TypeCapture type : capture.types()) restoreType(vml, type, identityMap);
    for (LineCapture line : capture.lines()) restoreLine(vml, line, identityMap, setupIds);
  }

  private static void restoreType(
      XSSFVMLDrawing vml, TypeCapture type, Map<String, String> identities) {
    CTShapetype definition = (CTShapetype) type.definition().copy();
    definition.setId(identities.get(type.definition().getId()));
    rewriteLocalReferences(definition, identities);
    retargetRelationships(vml, definition, type.relationships());
    insertTypeBeforeComments(vml, definition);
  }

  private static void restoreLine(
      XSSFVMLDrawing vml, LineCapture line, Map<String, String> identities, Set<String> setupIds) {
    CTShape shape = (CTShape) line.shape().copy();
    SignatureLineIdentities.retargetShape(shape, identities);
    rewriteLocalReferences(shape, identities);
    SignatureLineIdentities.assignSetupId(shape, setupIds);
    retargetRelationships(vml, shape, line.relationships());
    appendShape(vml, shape);
  }

  private static List<RelationshipCapture> relationships(
      XSSFSheet sheet, XSSFVMLDrawing source, XmlObject object) {
    Set<String> ids = relationshipIds(object);
    List<RelationshipCapture> result = new ArrayList<>();
    for (String id : ids) {
      var relationship = source.getPackagePart().getRelationship(id);
      if (relationship == null
          || !XSSFRelation.IMAGES.getRelation().equals(relationship.getRelationshipType())) {
        throw refused(
            sheet, WorkbookContentIssue.UNSUPPORTED_DEPENDENCY, "signatureLine.relationship", null);
      }
      if (relationship.getTargetMode() == TargetMode.EXTERNAL) {
        result.add(new ExternalImage(id, relationship.getTargetURI()));
      } else {
        PackagePart part =
            DrawingAdapter.relatedInternalPart(source.getPackagePart(), id)
                .orElseThrow(() ->
                    refused(
                        sheet,
                        WorkbookContentIssue.UNSUPPORTED_DEPENDENCY,
                        "signatureLine.preview",
                        null));
        // Observed bytes may be malformed or empty; this is preservation, not image authoring.
        result.add(
            new InternalImage(
                id,
                new SheetCopyAdapter.InternalRelationSnapshot(
                    relationship.getRelationshipType(),
                    part.getContentType(),
                    part.getPartName().getName(),
                    BinaryContent.readback(DrawingAdapter.partBytes(part)))));
      }
    }
    return List.copyOf(result);
  }

  // Relationship IDs are local to this restored native owner, in capture order.
  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static void retargetRelationships(
      XSSFVMLDrawing target, XmlObject object, List<RelationshipCapture> relationships) {
    Map<String, String> ids = new LinkedHashMap<>();
    for (RelationshipCapture relationship : relationships) {
      String id =
          switch (relationship) {
            case InternalImage internal -> {
              PackagePart copy =
                  SheetCopyAdapter.createCopiedPart(
                      target.getPackagePart().getPackage(),
                      internal.part(),
                      "signature preview",
                      "signature line");
              yield target
                  .addRelation(null, XSSFRelation.IMAGES, ImageCatalog.pictureDataForPart(copy))
                  .getRelationship()
                  .getId();
            }
            case ExternalImage external ->
                target
                    .getPackagePart()
                    .addExternalRelationship(
                        external.target().toString(), XSSFRelation.IMAGES.getRelation())
                    .getId();
          };
      ids.put(relationship.id(), id);
    }
    try (XmlCursor cursor = object.newCursor()) {
      while (cursor.hasNextToken()) {
        cursor.toNextToken();
        QName name = cursor.getName();
        if (cursor.isAttr() && name != null && relationshipAttribute(name)) {
          cursor.setTextValue(
              Objects.requireNonNull(
                  ids.get(cursor.getTextValue()), "Captured relationship must be mapped"));
        }
      }
    }
  }

  private static Set<String> relationshipIds(XmlObject object) {
    Set<String> ids = new LinkedHashSet<>();
    try (XmlCursor cursor = object.newCursor()) {
      while (cursor.hasNextToken()) {
        cursor.toNextToken();
        QName name = cursor.getName();
        if (cursor.isAttr() && name != null && relationshipAttribute(name))
          ids.add(cursor.getTextValue());
      }
    }
    return ids;
  }

  private static boolean relationshipAttribute(QName name) {
    return RELATIONSHIPS.equals(name.getNamespaceURI())
        || (OFFICE.equals(name.getNamespaceURI()) && "relid".equals(name.getLocalPart()));
  }

  private static boolean localReferenceAttribute(QName name) {
    return "idref".equals(name.getLocalPart())
        && (name.getNamespaceURI().isEmpty()
            || OFFICE.equals(name.getNamespaceURI())
            || "urn:schemas-microsoft-com:vml".equals(name.getNamespaceURI()));
  }

  private static void requireLocalReferences(
      XSSFSheet sheet, XmlObject object, Set<String> copiedIds) {
    try (XmlCursor cursor = object.newCursor()) {
      while (cursor.hasNextToken()) {
        cursor.toNextToken();
        QName name = cursor.getName();
        if (cursor.isAttr()
            && name != null
            && localReferenceAttribute(name)
            && !copiedIds.contains(cursor.getTextValue())) {
          throw refused(
              sheet,
              WorkbookContentIssue.UNSUPPORTED_DEPENDENCY,
              "signatureLine.localReference",
              null);
        }
      }
    }
  }

  private static void rewriteLocalReferences(XmlObject object, Map<String, String> identities) {
    try (XmlCursor cursor = object.newCursor()) {
      while (cursor.hasNextToken()) {
        cursor.toNextToken();
        QName name = cursor.getName();
        if (cursor.isAttr() && name != null && localReferenceAttribute(name)) {
          cursor.setTextValue(
              Objects.requireNonNull(
                  identities.get(cursor.getTextValue()),
                  "Captured local reference must be mapped"));
        }
      }
    }
  }

  // This invocation-local index preserves native type order and rejects duplicate owners.
  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static Map<String, CTShapetype> shapeTypes(XSSFVMLDrawing vml, XSSFSheet sheet) {
    Map<String, CTShapetype> result = new LinkedHashMap<>();
    try (XmlCursor cursor = vml.getDocument().getXml().newCursor()) {
      for (boolean child = cursor.toFirstChild(); child; child = cursor.toNextSibling()) {
        if (cursor.getObject() instanceof CTShapetype type
            && type.isSetId()
            && result.putIfAbsent(type.getId(), type) != null) {
          throw refused(
              sheet, WorkbookContentIssue.UNSUPPORTED_DEPENDENCY, "signatureLine.type", null);
        }
      }
    }
    return result;
  }

  private static int signatureCount(XSSFVMLDrawing vml) {
    int count = 0;
    try (XmlCursor cursor = vml.getDocument().getXml().newCursor()) {
      while (cursor.hasNextToken()) {
        cursor.toNextToken();
        QName name = cursor.getName();
        if (cursor.isStart()
            && name != null
            && OFFICE.equals(name.getNamespaceURI())
            && "signatureline".equals(name.getLocalPart())) count++;
      }
    }
    return count;
  }

  private static void appendShape(XSSFVMLDrawing target, CTShape shape) {
    CTGroup group = CTGroup.Factory.newInstance();
    group.addNewShape().set(shape);
    try (XmlCursor from = group.newCursor();
        XmlCursor to = target.getDocument().getXml().newCursor()) {
      from.toFirstChild();
      to.toEndToken();
      from.copyXml(to);
    }
  }

  private static void insertTypeBeforeComments(XSSFVMLDrawing target, CTShapetype type) {
    CTGroup group = CTGroup.Factory.newInstance();
    group.addNewShapetype().set(type);
    try (XmlCursor from = group.newCursor();
        XmlCursor to = target.getDocument().getXml().newCursor()) {
      from.toFirstChild();
      for (boolean child = to.toFirstChild(); child; child = to.toNextSibling()) {
        if (to.getObject() instanceof CTShapetype existing
            && COMMENT_TYPE.equals(existing.getId())) {
          from.copyXml(to);
          return;
        }
      }
      throw new IllegalStateException("Owned target VML must retain its native comment type");
    }
  }

  static WorkbookContentRefusalException refused(
      XSSFSheet sheet, WorkbookContentIssue issue, String field, @Nullable Throwable cause) {
    return new WorkbookContentRefusalException(
        issue,
        Optional.of(
            new WorkbookContentLocation(
                Optional.of(new SheetLabel(sheet.getSheetName())),
                field.startsWith("worksheet.")
                    ? Optional.of(sheet.getPackagePart().getPartName().getName())
                    : Optional.empty(),
                Optional.of(field),
                Optional.empty(),
                Optional.empty())),
        cause);
  }

  /** Detached native state captured before the clone mutates package relationships. */
  record Capture(
      List<LineCapture> lines, List<TypeCapture> types, Optional<String> sourcePartName) {
    Capture {
      lines = List.copyOf(lines);
      types = List.copyOf(types);
      Objects.requireNonNull(sourcePartName);
    }
  }

  /** One independently owned line and its referenced image payloads. */
  private record LineCapture(CTShape shape, List<RelationshipCapture> relationships) {}

  /** One native type dependency, retaining its full XML and image relationships. */
  private record TypeCapture(CTShapetype definition, List<RelationshipCapture> relationships) {}

  /** One captured image dependency; copying never fetches an external resource. */
  private sealed interface RelationshipCapture permits InternalImage, ExternalImage {
    /** Native relationship identifier before retargeting to the copied owner. */
    String id();
  }

  /** Internal preview bytes may be opaque; each restored owner gets a private package part. */
  private record InternalImage(String id, SheetCopyAdapter.InternalRelationSnapshot part)
      implements RelationshipCapture {}

  /** External preview metadata is preserved without network acquisition. */
  private record ExternalImage(String id, URI target) implements RelationshipCapture {}
}
