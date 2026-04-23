package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.io.IOException;
import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandles;
import java.lang.invoke.MethodType;
import java.lang.invoke.VarHandle;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Objects;
import java.util.function.Function;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.ooxml.POIXMLDocumentPart.RelationPart;
import org.apache.poi.ooxml.POIXMLRelation;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xssf.usermodel.XSSFFactory;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Keeps POI's private workbook picture catalog aligned with the actual OOXML package media parts.
 */
final class ExcelWorkbookImageCatalogSupport {
  private static final VarHandle PICTURES_FIELD = requirePicturesField(MethodHandles.lookup());
  private static final Function<PackagePart, XSSFPictureData> PICTURE_CONSTRUCTOR =
      requirePictureConstructor(MethodHandles.lookup());
  private static final Pattern IMAGE_PART_NAME =
      Pattern.compile("^/xl/media/image(?<index>\\d+)\\.[^/]+$");
  private static final Comparator<PackagePart> IMAGE_PART_ORDER =
      Comparator.comparingInt(ExcelWorkbookImageCatalogSupport::imagePartIndex)
          .thenComparing(part -> part.getPartName().getName());

  private ExcelWorkbookImageCatalogSupport() {}

  /**
   * Adds one picture after reconciling any image parts POI did not register in its workbook-level
   * catalog, such as VML-backed signature-line previews.
   */
  static int addPicture(XSSFWorkbook workbook, byte[] bytes, int poiPictureType) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(bytes, "bytes must not be null");
    synchronizePictureCatalog(workbook);
    List<XSSFPictureData> pictures = mutablePictures(workbook);
    POIXMLRelation pictureRelation = pictureRelation(poiPictureType);
    int pictureIndex = pictures.size();
    int partNumber =
        requireAllocatedPartNumber(workbook.getNextPartNumber(pictureRelation, pictureIndex + 1));
    RelationPart relationPart =
        workbook.createRelationship(pictureRelation, XSSFFactory.getInstance(), partNumber, true);
    XSSFPictureData pictureData = relationPart.getDocumentPart();
    writePictureBytes(pictureData.getPackagePart(), bytes);
    pictures.add(pictureData);
    return pictureIndex;
  }

  /** Rebuilds POI's private picture registry from the actual OOXML package media parts. */
  static void synchronizePictureCatalog(XSSFWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    List<XSSFPictureData> pictures = mutablePictures(workbook);
    pictures.clear();
    for (PackagePart imagePart : imageParts(workbook)) {
      pictures.add(newPictureData(imagePart));
    }
  }

  static List<String> pictureCatalogPartNames(XSSFWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    return mutablePictures(workbook).stream()
        .map(picture -> picture.getPackagePart().getPartName().getName())
        .toList();
  }

  static List<String> packageImagePartNames(XSSFWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    return imageParts(workbook).stream().map(part -> part.getPartName().getName()).toList();
  }

  @SuppressWarnings("unchecked")
  private static List<XSSFPictureData> mutablePictures(XSSFWorkbook workbook) {
    List<XSSFPictureData> pictures = (List<XSSFPictureData>) PICTURES_FIELD.get(workbook);
    if (pictures == null) {
      pictures = new ArrayList<>();
      PICTURES_FIELD.set(workbook, pictures);
    }
    return pictures;
  }

  private static List<PackagePart> imageParts(XSSFWorkbook workbook) {
    return imageParts(() -> workbook.getPackage().getParts());
  }

  static List<PackagePart> imageParts(PackagePartsSupplier partsSupplier) {
    Objects.requireNonNull(partsSupplier, "partsSupplier must not be null");
    List<PackagePart> imageParts = new ArrayList<>();
    try {
      for (PackagePart part : partsSupplier.get()) {
        if (isWorkbookImagePart(part)) {
          imageParts.add(part);
        }
      }
    } catch (InvalidFormatException exception) {
      throw new IllegalStateException("Failed to inspect workbook package media parts", exception);
    }
    imageParts.sort(IMAGE_PART_ORDER);
    return List.copyOf(imageParts);
  }

  private static boolean isWorkbookImagePart(PackagePart part) {
    String partName = part.getPartName().getName();
    if (!partName.startsWith("/xl/media/")) {
      return false;
    }
    try {
      ExcelPictureFormat.fromContentType(part.getContentType());
      return true;
    } catch (IllegalArgumentException ignored) {
      return false;
    }
  }

  private static int imagePartIndex(PackagePart part) {
    Matcher matcher = IMAGE_PART_NAME.matcher(part.getPartName().getName());
    if (!matcher.matches()) {
      return Integer.MAX_VALUE;
    }
    return Integer.parseInt(matcher.group("index"));
  }

  private static XSSFPictureData newPictureData(PackagePart part) {
    return PICTURE_CONSTRUCTOR.apply(part);
  }

  static POIXMLRelation pictureRelation(int poiPictureType) {
    return switch (ExcelPicturePoiBridge.fromPoiPictureType(poiPictureType)) {
      case EMF -> XSSFRelation.IMAGE_EMF;
      case WMF -> XSSFRelation.IMAGE_WMF;
      case PICT -> XSSFRelation.IMAGE_PICT;
      case JPEG -> XSSFRelation.IMAGE_JPEG;
      case PNG -> XSSFRelation.IMAGE_PNG;
      case DIB -> XSSFRelation.IMAGE_DIB;
      case GIF -> XSSFRelation.IMAGE_GIF;
      case TIFF -> XSSFRelation.IMAGE_TIFF;
      case EPS -> XSSFRelation.IMAGE_EPS;
      case BMP -> XSSFRelation.IMAGE_BMP;
      case WPG -> XSSFRelation.IMAGE_WPG;
    };
  }

  static int requireAllocatedPartNumber(int partNumber) {
    if (partNumber < 0) {
      throw new IllegalStateException("Failed to allocate a workbook image part");
    }
    return partNumber;
  }

  static void writePictureBytes(PackagePart part, byte[] bytes) {
    Objects.requireNonNull(part, "part must not be null");
    Objects.requireNonNull(bytes, "bytes must not be null");
    try (var outputStream = part.getOutputStream()) {
      outputStream.write(bytes);
    } catch (IOException exception) {
      throw new IllegalStateException("Failed to write workbook image bytes", exception);
    }
  }

  static VarHandle requirePicturesField(MethodHandles.Lookup lookup) {
    return PoiPrivateAccessSupport.requireVarHandle(
        lookup,
        XSSFWorkbook.class,
        "pictures",
        List.class,
        "Failed to access POI workbook picture catalog");
  }

  static Function<PackagePart, XSSFPictureData> requirePictureConstructor(
      MethodHandles.Lookup lookup) {
    MethodHandle constructor =
        PoiPrivateAccessSupport.requireConstructor(
                lookup,
                XSSFPictureData.class,
                MethodType.methodType(void.class, PackagePart.class),
                "Failed to access POI picture-data constructor")
            .asType(MethodType.methodType(XSSFPictureData.class, PackagePart.class));
    return pictureConstructor(constructor);
  }

  @SuppressWarnings("unchecked")
  private static Function<PackagePart, XSSFPictureData> pictureConstructor(
      MethodHandle constructor) {
    return PoiPrivateAccessSupport.asInterfaceInstance(Function.class, constructor);
  }

  /** Supplies workbook package parts while preserving the checked POI inspection failure type. */
  @FunctionalInterface
  interface PackagePartsSupplier {
    /** Returns the package parts to inspect. */
    Iterable<PackagePart> get() throws InvalidFormatException;
  }
}
