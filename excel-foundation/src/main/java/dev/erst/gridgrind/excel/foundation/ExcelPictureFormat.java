package dev.erst.gridgrind.excel.foundation;

/** Supported picture formats for authored and extracted drawing images. */
public enum ExcelPictureFormat {
  EMF("image/x-emf", ".emf"),
  WMF("image/x-wmf", ".wmf"),
  PICT("image/x-pict", ".pict"),
  JPEG("image/jpeg", ".jpg"),
  PNG("image/png", ".png"),
  DIB("image/dib", ".dib"),
  GIF("image/gif", ".gif"),
  TIFF("image/tiff", ".tif"),
  EPS("image/x-eps", ".eps"),
  BMP("image/x-ms-bmp", ".bmp"),
  WPG("image/x-wpg", ".wpg");

  private final String defaultContentType;
  private final String defaultExtension;

  ExcelPictureFormat(String defaultContentType, String defaultExtension) {
    this.defaultContentType = defaultContentType;
    this.defaultExtension = defaultExtension;
  }

  /** Returns the default MIME type used for extracted payloads. */
  public String defaultContentType() {
    return defaultContentType;
  }

  /** Returns the default file extension used for extracted payloads. */
  public String defaultExtension() {
    return defaultExtension;
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
