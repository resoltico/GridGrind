package dev.erst.gridgrind.excel.drawing;

import dev.erst.gridgrind.excel.ExcelEnumMappingSupport;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.util.Map;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Maps GridGrind picture formats to and from Apache POI picture-type constants. */
@SuppressWarnings("PMD.CommentRequired")
public final class ExcelPicturePoiBridge {
  private static final Map<ExcelPictureFormat, Integer> TO_POI_PICTURE_TYPE =
      ExcelEnumMappingSupport.exactEnumMap(
          ExcelPictureFormat.class,
          "Apache POI picture-type mapping",
          Map.ofEntries(
              Map.entry(ExcelPictureFormat.EMF, Workbook.PICTURE_TYPE_EMF),
              Map.entry(ExcelPictureFormat.WMF, Workbook.PICTURE_TYPE_WMF),
              Map.entry(ExcelPictureFormat.PICT, Workbook.PICTURE_TYPE_PICT),
              Map.entry(ExcelPictureFormat.JPEG, Workbook.PICTURE_TYPE_JPEG),
              Map.entry(ExcelPictureFormat.PNG, Workbook.PICTURE_TYPE_PNG),
              Map.entry(ExcelPictureFormat.DIB, Workbook.PICTURE_TYPE_DIB),
              Map.entry(ExcelPictureFormat.GIF, XSSFWorkbook.PICTURE_TYPE_GIF),
              Map.entry(ExcelPictureFormat.TIFF, XSSFWorkbook.PICTURE_TYPE_TIFF),
              Map.entry(ExcelPictureFormat.EPS, XSSFWorkbook.PICTURE_TYPE_EPS),
              Map.entry(ExcelPictureFormat.BMP, XSSFWorkbook.PICTURE_TYPE_BMP),
              Map.entry(ExcelPictureFormat.WPG, XSSFWorkbook.PICTURE_TYPE_WPG)));

  private static final Map<Integer, ExcelPictureFormat> FROM_POI_PICTURE_TYPE =
      Map.ofEntries(
          Map.entry(Workbook.PICTURE_TYPE_EMF, ExcelPictureFormat.EMF),
          Map.entry(Workbook.PICTURE_TYPE_WMF, ExcelPictureFormat.WMF),
          Map.entry(Workbook.PICTURE_TYPE_PICT, ExcelPictureFormat.PICT),
          Map.entry(Workbook.PICTURE_TYPE_JPEG, ExcelPictureFormat.JPEG),
          Map.entry(Workbook.PICTURE_TYPE_PNG, ExcelPictureFormat.PNG),
          Map.entry(Workbook.PICTURE_TYPE_DIB, ExcelPictureFormat.DIB),
          Map.entry(XSSFWorkbook.PICTURE_TYPE_GIF, ExcelPictureFormat.GIF),
          Map.entry(XSSFWorkbook.PICTURE_TYPE_TIFF, ExcelPictureFormat.TIFF),
          Map.entry(XSSFWorkbook.PICTURE_TYPE_EPS, ExcelPictureFormat.EPS),
          Map.entry(XSSFWorkbook.PICTURE_TYPE_BMP, ExcelPictureFormat.BMP),
          Map.entry(XSSFWorkbook.PICTURE_TYPE_WPG, ExcelPictureFormat.WPG));

  private ExcelPicturePoiBridge() {}

  public static int toPoiPictureType(ExcelPictureFormat format) {
    return ExcelEnumMappingSupport.requireMappedValue(
        TO_POI_PICTURE_TYPE, format, "GridGrind picture format");
  }

  public static ExcelPictureFormat fromPoiPictureType(int poiPictureType) {
    ExcelPictureFormat resolved = FROM_POI_PICTURE_TYPE.get(poiPictureType);
    if (resolved == null) {
      throw new IllegalArgumentException("Unsupported POI picture type: " + poiPictureType);
    }
    return resolved;
  }
}
