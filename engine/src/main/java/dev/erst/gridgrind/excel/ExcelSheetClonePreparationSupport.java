package dev.erst.gridgrind.excel;

import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandles;
import java.lang.invoke.MethodType;
import java.lang.invoke.VarHandle;
import java.util.List;
import java.util.Objects;
import java.util.function.BiFunction;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTHyperlink;

/** Materializes POI sheet-clone prerequisites that are otherwise deferred until save-time. */
final class ExcelSheetClonePreparationSupport {
  private static final VarHandle HYPERLINKS_FIELD = requireHyperlinksField(MethodHandles.lookup());
  private static final BiFunction<CTHyperlink, PackageRelationship, XSSFHyperlink>
      HYPERLINK_CONSTRUCTOR = requireHyperlinkConstructor(MethodHandles.lookup());

  /** Ensures a source sheet is internally consistent before handing it to POI clone logic. */
  void prepareSourceSheetForClone(XSSFSheet sourceSheet) {
    Objects.requireNonNull(sourceSheet, "sourceSheet must not be null");
    materializeExternalHyperlinkRelationships(sourceSheet);
  }

  private static void materializeExternalHyperlinkRelationships(XSSFSheet sourceSheet) {
    List<XSSFHyperlink> hyperlinks = mutableHyperlinks(sourceSheet);
    for (int index = 0; index < hyperlinks.size(); index++) {
      hyperlinks.set(
          index, materializedExternalHyperlinkRelationship(sourceSheet, hyperlinks.get(index)));
    }
  }

  private static XSSFHyperlink materializedExternalHyperlinkRelationship(
      XSSFSheet sourceSheet, XSSFHyperlink hyperlink) {
    if (hyperlink == null || !hyperlink.needsRelationToo()) {
      return hyperlink;
    }
    String relationId = hyperlink.getCTHyperlink().getId();
    PackagePart sheetPart = sourceSheet.getPackagePart();
    PackageRelationship relationship =
        relationId == null ? null : sheetPart.getRelationship(relationId);
    if (relationship == null) {
      relationship =
          sheetPart.addExternalRelationship(
              hyperlink.getAddress(), XSSFRelation.SHEET_HYPERLINKS.getRelation(), relationId);
      hyperlink.getCTHyperlink().setId(relationship.getId());
    }
    return newHyperlink(hyperlink.getCTHyperlink(), relationship);
  }

  @SuppressWarnings("unchecked")
  private static List<XSSFHyperlink> mutableHyperlinks(XSSFSheet sourceSheet) {
    return (List<XSSFHyperlink>) HYPERLINKS_FIELD.get(sourceSheet);
  }

  private static XSSFHyperlink newHyperlink(
      CTHyperlink ctHyperlink, PackageRelationship relationship) {
    return HYPERLINK_CONSTRUCTOR.apply(ctHyperlink, relationship);
  }

  static VarHandle requireHyperlinksField(MethodHandles.Lookup lookup) {
    return PoiPrivateAccessSupport.requireVarHandle(
        lookup,
        XSSFSheet.class,
        "hyperlinks",
        List.class,
        "Failed to access POI sheet hyperlink registry");
  }

  static BiFunction<CTHyperlink, PackageRelationship, XSSFHyperlink> requireHyperlinkConstructor(
      MethodHandles.Lookup lookup) {
    MethodHandle constructor =
        PoiPrivateAccessSupport.requireConstructor(
                lookup,
                XSSFHyperlink.class,
                MethodType.methodType(void.class, CTHyperlink.class, PackageRelationship.class),
                "Failed to access POI hyperlink constructor")
            .asType(
                MethodType.methodType(
                    XSSFHyperlink.class, CTHyperlink.class, PackageRelationship.class));
    return hyperlinkConstructor(constructor);
  }

  @SuppressWarnings("unchecked")
  private static BiFunction<CTHyperlink, PackageRelationship, XSSFHyperlink> hyperlinkConstructor(
      MethodHandle constructor) {
    return PoiPrivateAccessSupport.asInterfaceInstance(BiFunction.class, constructor);
  }
}
