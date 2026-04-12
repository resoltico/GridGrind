package dev.erst.gridgrind.excel;

import org.apache.poi.ss.usermodel.Workbook;

/** Supported picture formats for authored and extracted drawing images. */
public enum ExcelPictureFormat {
  EMF(Workbook.PICTURE_TYPE_EMF, "image/x-emf", ".emf"),
  WMF(Workbook.PICTURE_TYPE_WMF, "image/wmf", ".wmf"),
  PICT(Workbook.PICTURE_TYPE_PICT, "image/pict", ".pict"),
  JPEG(Workbook.PICTURE_TYPE_JPEG, "image/jpeg", ".jpg"),
  PNG(Workbook.PICTURE_TYPE_PNG, "image/png", ".png"),
  DIB(Workbook.PICTURE_TYPE_DIB, "image/bmp", ".dib");

  private final int poiPictureType;
  private final String defaultContentType;
  private final String defaultExtension;

  ExcelPictureFormat(int poiPictureType, String defaultContentType, String defaultExtension) {
    this.poiPictureType = poiPictureType;
    this.defaultContentType = defaultContentType;
    this.defaultExtension = defaultExtension;
  }

  /** Returns the Apache POI workbook picture type constant. */
  public int poiPictureType() {
    return poiPictureType;
  }

  /** Returns the default MIME type used for extracted payloads. */
  public String defaultContentType() {
    return defaultContentType;
  }

  /** Returns the default file extension used for extracted payloads. */
  public String defaultExtension() {
    return defaultExtension;
  }

  /** Returns the format matching the Apache POI picture type constant. */
  public static ExcelPictureFormat fromPoiPictureType(int poiPictureType) {
    for (ExcelPictureFormat format : values()) {
      if (format.poiPictureType == poiPictureType) {
        return format;
      }
    }
    throw new IllegalArgumentException("Unsupported POI picture type: " + poiPictureType);
  }

  /** Returns the format matching one OOXML content type. */
  public static ExcelPictureFormat fromContentType(String contentType) {
    for (ExcelPictureFormat format : values()) {
      if (format.defaultContentType.equals(contentType)) {
        return format;
      }
    }
    throw new IllegalArgumentException("Unsupported picture content type: " + contentType);
  }
}
