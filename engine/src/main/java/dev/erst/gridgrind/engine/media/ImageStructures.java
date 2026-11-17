package dev.erst.gridgrind.engine.media;

/** Concrete declared-format dispatch; never a provider registry or content-type guess. */
final class ImageStructures {
  private ImageStructures() {}

  static ImageGeometry check(MediaInput input) {
    input.record();
    return switch (input.format) {
      case PNG -> PngStructure.check(input);
      case JPEG -> JpegStructure.check(input);
      case GIF -> GifStructure.check(input);
      case TIFF -> TiffStructure.check(input);
      case BMP -> BitmapStructure.check(input, true);
      case DIB -> BitmapStructure.check(input, false);
      case EMF -> MetafileStructure.emf(input);
      case WMF -> MetafileStructure.wmf(input);
      case PICT -> PictStructure.check(input);
      case EPS -> EpsStructure.check(input);
      case WPG -> WpgStructure.check(input);
    };
  }
}
