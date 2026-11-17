package dev.erst.gridgrind.engine.backend.xssf;

import com.microsoft.schemas.vml.CTShape;
import com.microsoft.schemas.vml.CTShapetype;
import dev.erst.gridgrind.core.problem.WorkbookContentIssue;
import dev.erst.gridgrind.engine.drawing.SignatureAdapter;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.UUID;
import java.util.function.Supplier;
import javax.xml.namespace.QName;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFVMLDrawing;
import org.apache.xmlbeans.XmlCursor;

/** Reserves source-local VML owners and allocates independent local/setup identities for copies. */
final class SignatureLineIdentities {
  private SignatureLineIdentities() {}

  // Source IDs are indexed once in deterministic native order within this invocation.
  @SuppressWarnings("PMD.UseConcurrentHashMap")
  static Set<String> sourceIds(XSSFSheet source, Set<String> typeIds, List<CTShape> shapes) {
    Map<String, Integer> owners = new LinkedHashMap<>();
    int owner = 0;
    for (String id : typeIds) {
      reserveSourceId(source, owners, id, owner);
      owner++;
    }
    for (CTShape shape : shapes) {
      if (shape.isSetId()) reserveSourceId(source, owners, shape.getId(), owner);
      if (shape.isSetSpid()) reserveSourceId(source, owners, shape.getSpid(), owner);
      owner++;
    }
    return Set.copyOf(owners.keySet());
  }

  // A copy owns its ordered mapping; no mutable map is shared across invocations.
  @SuppressWarnings("PMD.UseConcurrentHashMap")
  static Map<String, String> mapping(
      XSSFVMLDrawing target, List<CTShapetype> types, List<CTShape> shapes) {
    Set<String> occupied = occupiedIds(target, types, shapes);
    Map<String, String> identities = new LinkedHashMap<>();
    for (CTShapetype type : types) {
      identities.put(type.getId(), nextLocalId(occupied, "_gridgrind_signature_type_"));
    }
    for (CTShape shape : shapes) {
      if (shape.isSetId())
        identities.put(shape.getId(), nextLocalId(occupied, "_gridgrind_signature_"));
      if (shape.isSetSpid())
        identities.computeIfAbsent(
            shape.getSpid(), ignored -> nextLocalId(occupied, "_gridgrind_signature_"));
    }
    return identities;
  }

  private static Set<String> occupiedIds(
      XSSFVMLDrawing target, List<CTShapetype> types, List<CTShape> shapes) {
    Set<String> occupied = localIds(target);
    for (CTShapetype type : types) occupied.add(type.getId());
    for (CTShape shape : shapes) {
      if (shape.isSetId()) occupied.add(shape.getId());
      if (shape.isSetSpid()) occupied.add(shape.getSpid());
    }
    return occupied;
  }

  static void retargetShape(CTShape shape, Map<String, String> identities) {
    if (shape.isSetId()) shape.setId(identities.get(shape.getId()));
    if (shape.isSetSpid()) shape.setSpid(identities.get(shape.getSpid()));
    String type = shape.getType();
    if (type != null && type.startsWith("#"))
      shape.setType("#" + identities.get(type.substring(1)));
  }

  static void assignSetupId(CTShape shape, Set<String> occupied) {
    shape.getSignaturelineArray(0).setId(nextSetupId(occupied, UUID::randomUUID));
  }

  private static void reserveSourceId(
      XSSFSheet sheet, Map<String, Integer> owners, String id, int owner) {
    Integer previous = owners.putIfAbsent(id, owner);
    if (previous != null && previous.intValue() != owner) {
      throw SignatureLineCopy.refused(
          sheet, WorkbookContentIssue.UNSUPPORTED_DEPENDENCY, "signatureLine.localIdentity", null);
    }
  }

  private static Set<String> localIds(XSSFVMLDrawing vml) {
    Set<String> ids = new LinkedHashSet<>();
    try (XmlCursor cursor = vml.getDocument().getXml().newCursor()) {
      while (cursor.hasNextToken()) {
        cursor.toNextToken();
        QName name = cursor.getName();
        if (cursor.isAttr()
            && name != null
            && ("id".equals(name.getLocalPart()) || "spid".equals(name.getLocalPart())))
          ids.add(cursor.getTextValue());
      }
    }
    return ids;
  }

  static Set<String> setupIds(XSSFSheet target) {
    Set<String> ids = new LinkedHashSet<>();
    for (int index = 0; index < target.getWorkbook().getNumberOfSheets(); index++) {
      XSSFSheet sheet = target.getWorkbook().getSheetAt(index);
      XSSFVMLDrawing vml = sheet.getVMLDrawing(false);
      if (vml == null) continue;
      for (CTShape shape : SignatureAdapter.signatureShapes(vml)) {
        for (var line : shape.getSignaturelineList()) {
          if (line.isSetId()) ids.add(canonicalGuid(line.getId()));
        }
      }
    }
    return ids;
  }

  static String nextSetupId(Set<String> occupied, Supplier<UUID> identities) {
    String id;
    do {
      id = identities.get().toString().toUpperCase(Locale.ROOT);
    } while (!occupied.add(id));
    return "{" + id + "}";
  }

  private static String canonicalGuid(String value) {
    return (value.length() == 38 && value.startsWith("{") && value.endsWith("}")
            ? value.substring(1, 37)
            : value)
        .toUpperCase(Locale.ROOT);
  }

  private static String nextLocalId(Set<String> occupied, String prefix) {
    for (int sequence = 1; ; sequence++) {
      String id = prefix + sequence;
      if (occupied.add(id)) return id;
    }
  }
}
