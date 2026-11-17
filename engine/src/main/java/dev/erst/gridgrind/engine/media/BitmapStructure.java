package dev.erst.gridgrind.engine.media;

import dev.erst.gridgrind.core.problem.MediaStructureReason;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import java.util.Optional;

/** Windows/OS2 bitmap layout and bounded raw/RLE or embedded-image coding. */
final class BitmapStructure {
  private final MediaInput input;
  private final int dibStart;
  private final long selectedRows;
  private final long firstScan;
  private final int colorUsage;
  private final Optional<MediaInput> separatePixels;
  private long pixelOffset = -1;
  private long header;
  private long declaredSize = -1;
  private int reserved1;
  private int reserved2;
  private long width;
  private long height;
  private long compression;
  private long sizeImage;
  private long colors;
  private long palette;
  private int bits;
  private int planes;
  private boolean os2;
  private boolean topDown;

  private BitmapStructure(
      MediaInput input,
      boolean fileHeader,
      long selectedRows,
      long firstScan,
      int colorUsage,
      Optional<MediaInput> separatePixels) {
    this.colorUsage = colorUsage;
    this.separatePixels = separatePixels;
    input.require(colorUsage >= 0 && colorUsage <= 2, MediaStructureReason.CODE);
    this.selectedRows = selectedRows;
    this.firstScan = firstScan;
    this.input = input;
    input.little = true;
    if (fileHeader) {
      input.magic(66, 77);
      declaredSize = input.u32();
      reserved1 = input.u16();
      reserved2 = input.u16();
      pixelOffset = input.u32();
    }
    dibStart = input.offset();
  }

  static ImageGeometry check(MediaInput input, boolean fileHeader) {
    return new BitmapStructure(input, fileHeader, -1, 0, 0, Optional.empty()).read();
  }

  static ImageGeometry checkDibMember(
      MediaInput input, long selectedRows, long firstScan, int colorUsage) {
    return new BitmapStructure(input, false, selectedRows, firstScan, colorUsage, Optional.empty())
        .read();
  }

  static ImageGeometry checkParts(
      MediaInput header, MediaInput pixels, long selectedRows, long firstScan, int colorUsage) {
    var info = header.nested(0, header.size(), ExcelPictureFormat.DIB);
    var data = pixels.nested(0, pixels.size(), ExcelPictureFormat.DIB);
    return new BitmapStructure(info, false, selectedRows, firstScan, colorUsage, Optional.of(data))
        .read();
  }

  private ImageGeometry read() {
    layout();
    geometry();
    input.seek(dibStart + header);
    if (!os2 && (compression == 3 || compression == 6)) masks();
    palette();
    profile();
    pixels();
    return new ImageGeometry(width, height);
  }

  private void layout() {
    header = input.u32();
    input.require(header <= input.size() - dibStart, MediaStructureReason.LENGTH);
    os2 =
        header == 12
            || header >= 16
                && header <= 64
                && (header != 40 && header != 52 && header != 56 || declaredSize == 14 + header);
    if (os2) input.require(os2Header(header), MediaStructureReason.HEADER);
    header();
    os2 |= header <= 64 && (compression == 3 && bits == 1 || compression == 4 && bits == 24);
    fileFacts();
  }

  private void geometry() {
    input.require(width > 0 && height != 0 && planes == 1, MediaStructureReason.DIMENSIONS);
    topDown = height < 0;
    height = Math.abs(height);
    input.require(
        selectedRows < 0
            || selectedRows <= height && firstScan >= 0 && firstScan <= height - selectedRows,
        MediaStructureReason.COUNT);
    boolean valid =
        switch (bits) {
          case 1, 4, 8, 16, 24, 32 -> true;
          case 2 -> os2;
          case 0 -> compression == 4 || compression == 5;
          default -> false;
        };
    input.require(valid, MediaStructureReason.CODE);
  }

  private void palette() {
    palette = bits > 0 && bits <= 8 ? (colors == 0 ? 1L << bits : colors) : colors;
    input.require(bits > 8 || palette <= 1L << bits, MediaStructureReason.COUNT);
    if (colorUsage != 2)
      input.skip(input.product(palette, colorUsage == 1 ? 2 : header == 12 ? 3 : 4));
    if (pixelOffset < 0) pixelOffset = input.offset();
    input.require(pixelOffset >= input.offset(), MediaStructureReason.OFFSET);
  }

  private void header() {
    if (header == 12) {
      width = input.u16();
      height = input.u16();
      planes = input.u16();
      bits = input.u16();
      return;
    }
    if (header == 16) {
      width = input.u32();
      height = input.u32();
      planes = input.u16();
      bits = input.u16();
      return;
    }
    if (!os2 && header != 40 && header != 52 && header != 56 && header != 108 && header != 124)
      input.profile("bitmap/header-" + header);
    width = os2 ? input.u32() : input.s32();
    height = os2 ? input.u32() : input.s32();
    planes = input.u16();
    bits = input.u16();
    compression = headerField(16);
    sizeImage = headerField(20);
    colors = headerField(32);
  }

  private static boolean os2Header(long size) {
    return switch ((int) size) {
      case 12, 16, 20, 24, 28, 32, 36, 40, 42, 44, 46, 48, 52, 56, 60, 64 -> true;
      default -> false;
    };
  }

  private MediaInput body(long length) {
    return separatePixels
        .map(value -> value.region(0, length))
        .orElseGet(() -> input.region(pixelOffset, length));
  }

  private long codedBytes() {
    return sizeImage > 0
        ? sizeImage
        : separatePixels
            .map(value -> (long) value.size())
            .orElseGet(() -> input.size() - pixelOffset);
  }

  private long headerField(int offset) {
    return header >= offset + 4 ? input.region(dibStart + offset, 4).u32() : 0;
  }

  private void fileFacts() {
    if (declaredSize < 0) return;
    if (os2) {
      input.require(
          declaredSize == 14 + header || declaredSize == input.size(), MediaStructureReason.LENGTH);
    } else {
      input.require(declaredSize == input.size(), MediaStructureReason.LENGTH);
      input.require(reserved1 == 0 && reserved2 == 0, MediaStructureReason.HEADER);
    }
  }

  private void profile() {
    if (header != 124) return;
    var profile = input.region(dibStart + 112, 12);
    long offset = profile.u32();
    long length = profile.u32();
    profile.u32();
    if (length > 0) input.region(dibStart + offset, length);
  }

  private void masks() {
    input.require(bits == 16 || bits == 32, MediaStructureReason.CODE);
    int count = maskCount();
    var masks =
        header >= 52
            ? input.region(dibStart + 40, Math.min(header - 40, 16))
            : input.take(count * 4L);
    long occupied = 0;
    for (int i = 0; i < count; i++) {
      long mask = masks.u32();
      input.require(
          mask != 0 && (mask >>> bits) == 0 && (occupied & mask) == 0, MediaStructureReason.CODE);
      long normalized = mask / Long.lowestOneBit(mask);
      input.require((normalized & (normalized + 1)) == 0, MediaStructureReason.CODE);
      occupied |= mask;
    }
  }

  private int maskCount() {
    if (compression == 6) return 4;
    if (header >= 56 && input.region(dibStart + 52, 4).u32() != 0) return 4;
    return 3;
  }

  private void pixels() {
    if (os2 && compression == 3) {
      huffman();
      return;
    }
    if (compression == 4 || compression == 5) {
      embedded();
      return;
    }
    long row = ((input.product(width, bits) + 31) / 32) * 4;
    long rows = selectedRows >= 0 && rawCompression() ? selectedRows : height;
    input.budget.raster(width, height);
    input.budget.expand(input.product(row, rows));
    if (rawCompression()) {
      long required = input.product(row, rows);
      if (selectedRows >= 0) {
        body(required);
        return;
      }
      input.require(sizeImage == 0 || sizeImage >= required, MediaStructureReason.LENGTH);
      body(sizeImage == 0 ? required : sizeImage);
      return;
    }
    if (compression == 1 || compression == 2) {
      input.require(
          !topDown && bits == (compression == 1 ? 8 : 4) && (sizeImage > 0 || os2),
          MediaStructureReason.METADATA);
      rle(body(codedBytes()), width, height, bits, palette);
      return;
    }
    input.profile("bitmap/compression-" + compression);
  }

  private boolean rawCompression() {
    return compression == 0 || !os2 && (compression == 3 || compression == 6);
  }

  private void huffman() {
    input.require(bits == 1 && !topDown, MediaStructureReason.METADATA);
    input.raster(width, height, ((width + 31) / 32) * 4);
    FaxStructure.check(body(codedBytes()), width, height, 3, 0, 1);
  }

  private void embedded() {
    if (os2 && compression == 4) {
      input.raster(width, height, ((input.product(width, 24) + 31) / 32) * 4);
      rle(body(codedBytes()), width, height, 24, 0);
      return;
    }
    input.require(!topDown && sizeImage > 0, MediaStructureReason.METADATA);
    var encoded = body(sizeImage);
    var nested =
        encoded.nested(
            0, encoded.size(), compression == 4 ? ExcelPictureFormat.JPEG : ExcelPictureFormat.PNG);
    var image = ImageStructures.check(nested);
    input.require(
        image.width() == width && image.height() == height, MediaStructureReason.DIMENSIONS);
  }

  private static void rle(MediaInput in, long width, long height, int bits, long palette) {
    long x = 0;
    long y = 0;
    boolean end = false;
    while (in.remaining() > 0 && !end) {
      int count = in.u8();
      int value = in.u8();
      if (count > 0) {
        in.require(y < height && count <= width - x, MediaStructureReason.COUNT);
        if (bits == 8) in.require(value < palette, MediaStructureReason.CODE);
        else if (bits == 4)
          in.require(
              (value >> 4) < palette && (count == 1 || (value & 15) < palette),
              MediaStructureReason.CODE);
        else {
          in.u8();
          in.u8();
        }
        x += count;
        continue;
      }
      if (value == 0) {
        x = 0;
        y++;
        in.require(y <= height, MediaStructureReason.COUNT);
      } else if (value == 1) end = true;
      else if (value == 2) {
        x += in.u8();
        y += in.u8();
        in.require(x <= width && y < height, MediaStructureReason.COUNT);
      } else {
        in.require(y < height && value <= width - x, MediaStructureReason.COUNT);
        int bytes = bits == 4 ? (value + 1) / 2 : bits == 24 ? value * 3 : value;
        var literals = in.take(bytes);
        if (bits == 8)
          while (literals.remaining() > 0)
            in.require(literals.u8() < palette, MediaStructureReason.CODE);
        else if (bits == 4)
          for (int i = 0; i < value; i++) {
            int packed = in.bytes[literals.start + i / 2] & 255;
            in.budget.inspect(1);
            in.require(
                (i % 2 == 0 ? packed >> 4 : packed & 15) < palette, MediaStructureReason.CODE);
          }
        if ((bytes & 1) != 0) in.u8();
        x += value;
      }
    }
    in.require(end, MediaStructureReason.TERMINATOR);
  }
}
