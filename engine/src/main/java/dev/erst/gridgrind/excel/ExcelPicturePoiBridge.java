package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Maps GridGrind picture formats to and from Apache POI picture-type constants. */
final class ExcelPicturePoiBridge {
  private ExcelPicturePoiBridge() {}

  static int toPoiPictureType(ExcelPictureFormat format) {
    return switch (format) {
      case EMF -> Workbook.PICTURE_TYPE_EMF;
      case WMF -> Workbook.PICTURE_TYPE_WMF;
      case PICT -> Workbook.PICTURE_TYPE_PICT;
      case JPEG -> Workbook.PICTURE_TYPE_JPEG;
      case PNG -> Workbook.PICTURE_TYPE_PNG;
      case DIB -> Workbook.PICTURE_TYPE_DIB;
      case GIF -> XSSFWorkbook.PICTURE_TYPE_GIF;
      case TIFF -> XSSFWorkbook.PICTURE_TYPE_TIFF;
      case EPS -> XSSFWorkbook.PICTURE_TYPE_EPS;
      case BMP -> XSSFWorkbook.PICTURE_TYPE_BMP;
      case WPG -> XSSFWorkbook.PICTURE_TYPE_WPG;
    };
  }

  static ExcelPictureFormat fromPoiPictureType(int poiPictureType) {
    return switch (poiPictureType) {
      case Workbook.PICTURE_TYPE_EMF -> ExcelPictureFormat.EMF;
      case Workbook.PICTURE_TYPE_WMF -> ExcelPictureFormat.WMF;
      case Workbook.PICTURE_TYPE_PICT -> ExcelPictureFormat.PICT;
      case Workbook.PICTURE_TYPE_JPEG -> ExcelPictureFormat.JPEG;
      case Workbook.PICTURE_TYPE_PNG -> ExcelPictureFormat.PNG;
      case Workbook.PICTURE_TYPE_DIB -> ExcelPictureFormat.DIB;
      case XSSFWorkbook.PICTURE_TYPE_GIF -> ExcelPictureFormat.GIF;
      case XSSFWorkbook.PICTURE_TYPE_TIFF -> ExcelPictureFormat.TIFF;
      case XSSFWorkbook.PICTURE_TYPE_EPS -> ExcelPictureFormat.EPS;
      case XSSFWorkbook.PICTURE_TYPE_BMP -> ExcelPictureFormat.BMP;
      case XSSFWorkbook.PICTURE_TYPE_WPG -> ExcelPictureFormat.WPG;
      default ->
          throw new IllegalArgumentException("Unsupported POI picture type: " + poiPictureType);
    };
  }
}
