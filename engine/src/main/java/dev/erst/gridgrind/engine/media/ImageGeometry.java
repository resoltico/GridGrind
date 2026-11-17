package dev.erst.gridgrind.engine.media;

/** Raster header facts; zero dimensions denote a vector without one raster canvas. */
record ImageGeometry(long width, long height, int sampleBits, int components) {
  ImageGeometry(long width, long height) {
    this(width, height, 0, 0);
  }

  static ImageGeometry vector() {
    return new ImageGeometry(0, 0);
  }
}
