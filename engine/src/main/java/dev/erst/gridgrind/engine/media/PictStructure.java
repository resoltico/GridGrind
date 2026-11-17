package dev.erst.gridgrind.engine.media;

import dev.erst.gridgrind.core.problem.MediaStructureReason;

/** Apple QuickDraw appendixA opcode/bitmap grammar for raw and file-prefixed PICT1/2. */
final class PictStructure {
  private PictStructure() {}

  static ImageGeometry check(MediaInput original) {
    int prefix = hasVersion(original, 0) ? 0 : 512;
    original.require(hasVersion(original, prefix), MediaStructureReason.HEADER);
    var in = original.region(prefix, original.size() - prefix);
    boolean words = header(in);
    boolean ended = false;
    while (in.remaining() > 0) {
      if (words && (in.offset() & 1) != 0) in.u8();
      in.record();
      int opcode = words ? in.u16() : in.u8();
      if (opcode == 255) {
        ended = true;
        break;
      }
      opcode(in, opcode);
    }
    in.require(ended && in.remaining() == 0, MediaStructureReason.TERMINATOR);
    return ImageGeometry.vector();
  }

  private static boolean header(MediaInput in) {
    int declared = in.u16();
    rectangle(in);
    int first = in.u8();
    boolean words = first == 0;
    if (words) {
      in.magic(17, 2, 255);
      in.require(in.u16() == 0xc00, MediaStructureReason.HEADER);
      var header = in.take(24);
      int version = header.s16();
      header.require(version == -1 || version == -2, MediaStructureReason.HEADER);
    } else {
      in.require(first == 17 && in.u8() == 1, MediaStructureReason.HEADER);
      in.require(declared == in.size(), MediaStructureReason.LENGTH);
    }
    return words;
  }

  private static boolean hasVersion(MediaInput in, int prefix) {
    if (prefix < 0 || in.size() - prefix < 12) return false;
    int at = in.start + prefix + 10;
    in.budget.inspect(2);
    return (in.bytes[at] & 255) == 17 && (in.bytes[at + 1] & 255) == 1
        || in.size() - prefix >= 14
            && in.bytes[at] == 0
            && (in.bytes[at + 1] & 255) == 17
            && (in.bytes[at + 2] & 255) == 2
            && (in.bytes[at + 3] & 255) == 255;
  }

  private static void opcode(MediaInput in, int op) {
    if (sized(in, op) || shape(in, op)) return;
    switch (op) {
      case 0, 0x1c, 0x1e -> {}
      case 1 -> region(in);
      case 2, 9, 10, 16, 32 -> in.skip(8);
      case 3, 5, 8, 13, 21, 22, 35, 0xa0 -> in.skip(2);
      case 4 -> in.skip(1);
      case 6, 7, 11, 12, 14, 15, 33 -> in.skip(4);
      case 0x12, 0x13, 0x14 -> pattern(in);
      case 0x1a, 0x1b, 0x1d, 0x1f, 34 -> in.skip(6);
      case 0x28 -> {
        in.skip(4);
        in.skip(in.u8());
      }
      case 0x29, 0x2a -> {
        in.u8();
        in.skip(in.u8());
      }
      case 0x2b -> {
        in.skip(2);
        in.skip(in.u8());
      }
      case 0x2c -> {
        int size = in.u16();
        var name = in.take(size);
        name.u16();
        name.skip(name.u8());
      }
      case 0x2d, 0x2e, 0x2f -> in.skip(in.u16());
      case 0x90, 0x91, 0x98, 0x99, 0x9a, 0x9b -> bitmap(in, op);
      case 0xa1 -> {
        in.u16();
        in.skip(in.u16());
      }
      default -> in.profile("PICT/unsized-opcode-" + op);
    }
  }

  private static boolean sized(MediaInput in, int op) {
    if (op >= 0x100 && op <= 0x7fff) {
      in.skip((op >>> 8) * 2L);
      return true;
    }
    if (op >= 0x8000) {
      if (op >= 0x8100) in.skip(in.u32());
      return true;
    }
    if (op >= 0xb0 && op <= 0xcf) return true;
    if (op >= 0xd0 && op <= 0xfe) {
      in.skip(in.u32());
      return true;
    }
    if (wordSized(op)) {
      in.skip(in.u16());
      return true;
    }
    return false;
  }

  private static boolean wordSized(int op) {
    return op >= 0xa2 && op <= 0xaf
        || op >= 0x24 && op <= 0x27
        || op >= 0x92 && op <= 0x97
        || op >= 0x9c && op <= 0x9f;
  }

  private static boolean shape(MediaInput in, int op) {
    if (op >= 0x30 && op <= 0x5f) {
      if ((op & 15) < 8) rectangle(in);
      return true;
    }
    if (op >= 0x60 && op <= 0x6f) {
      if ((op & 15) < 8) rectangle(in);
      in.skip(4);
      return true;
    }
    if (op >= 0x70 && op <= 0x77) {
      polygon(in);
      return true;
    }
    if (op >= 0x78 && op <= 0x7f || op >= 0x88 && op <= 0x8f) return true;
    if (op >= 0x80 && op <= 0x87) {
      region(in);
      return true;
    }
    return false;
  }

  private static void polygon(MediaInput in) {
    int bytes = in.u16();
    in.require(bytes >= 10 && (bytes - 10) % 4 == 0, MediaStructureReason.LENGTH);
    var p = in.take(bytes - 2);
    rectangle(p);
    p.skip(p.remaining());
  }

  private static void region(MediaInput in) {
    int bytes = in.u16();
    in.require(bytes >= 10 && (bytes & 1) == 0, MediaStructureReason.LENGTH);
    var region = in.take(bytes - 2);
    rectangle(region);
    if (bytes == 10) return;
    region.skip(region.remaining());
  }

  private static Rect rectangle(MediaInput in) {
    int top = in.s16();
    int left = in.s16();
    int bottom = in.s16();
    int right = in.s16();
    return new Rect(right - left, bottom - top);
  }

  private record Rect(long width, long height) {}

  private record Pixels(
      long width, long height, long row, int bits, int components, int pack, boolean pixmap) {}

  private static Pixels pixels(MediaInput in, boolean direct) {
    if (direct) in.u32();
    int encoded = in.u16();
    long row = encoded & 0x3fff;
    boolean pixmap = (encoded & 0x8000) != 0;
    Rect bounds = rectangle(in);
    in.require(row > 0 && (row & 1) == 0, MediaStructureReason.LENGTH);
    in.require(!direct || pixmap, MediaStructureReason.HEADER);
    Pixels result =
        pixmap
            ? map(in, bounds, row)
            : new Pixels(bounds.width, bounds.height, row, 1, 1, 0, false);
    in.require(row * 8 >= in.product(bounds.width, result.bits), MediaStructureReason.LENGTH);
    return result;
  }

  private static Pixels map(MediaInput in, Rect bounds, long row) {
    int version = in.u16();
    if (version != 0 && version != 4) in.profile("PICT/pixmap-version-" + version);
    int pack = in.u16();
    in.u32();
    in.skip(8);
    int type = in.u16();
    int bits = in.u16();
    int components = in.u16();
    int componentBits = in.u16();
    in.skip(12);
    pixelType(in, type, bits, components, componentBits);
    return new Pixels(bounds.width, bounds.height, row, bits, components, pack, true);
  }

  private static void pixelType(
      MediaInput in, int type, int bits, int components, int componentBits) {
    in.require(type == 0 || type == 16, MediaStructureReason.CODE);
    if (type == 0) {
      boolean indexed =
          switch (bits) {
            case 1, 2, 4, 8 -> true;
            default -> false;
          };
      in.require(indexed && components == 1 && componentBits == bits, MediaStructureReason.CODE);
    } else
      in.require(
          (bits == 16 || bits == 32)
              && components >= 3
              && components <= 4
              && componentBits > 0
              && components * componentBits <= bits,
          MediaStructureReason.CODE);
  }

  private static void colors(MediaInput in, int bits) {
    in.u32();
    in.u16();
    long count = 1L + in.u16();
    in.require(count <= 1L << bits, MediaStructureReason.COUNT);
    in.skip(count * 8);
  }

  private static void pattern(MediaInput in) {
    int type = in.u16();
    in.skip(8);
    if (type == 2) {
      in.skip(6);
      return;
    }
    if (type != 1) in.profile("PICT/pixel-pattern-type-" + type);
    var pixel = pixels(in, false);
    colors(in, pixel.bits);
    data(in, pixel, true);
  }

  private static void bitmap(MediaInput in, int opcode) {
    boolean direct = opcode >= 0x9a;
    var pixel = pixels(in, direct);
    if (pixel.pixmap && !direct) colors(in, pixel.bits);
    rectangle(in);
    rectangle(in);
    in.u16();
    if ((opcode & 1) != 0) region(in);
    data(in, pixel, opcode >= 0x98);
  }

  private static void data(MediaInput in, Pixels pixel, boolean packedOpcode) {
    int packing = pixel.pack;
    if (packing == 0) packing = pixel.bits == 16 ? 3 : pixel.bits == 32 ? 4 : 0;
    long row = pixel.row;
    if (packing == 2) {
      in.require(pixel.bits == 32, MediaStructureReason.CODE);
      row = row / 4 * 3;
    } else if (packing == 3) in.require(pixel.bits == 16, MediaStructureReason.CODE);
    else if (packing == 4) {
      in.require(pixel.bits == 32, MediaStructureReason.CODE);
      row = row / 4 * pixel.components;
    } else if (packing > 4) in.profile("PICT/pixel-packing-" + packing);
    in.raster(pixel.width, pixel.height, row);
    if (!packedOpcode || packing == 1 || packing == 2 || pixel.row < 8) {
      in.skip(in.product(row, pixel.height));
      return;
    }
    for (long y = 0; y < pixel.height; y++) {
      long length = pixel.row > 250 ? in.u16() : in.u8();
      var coded = in.take(length);
      pack(coded, row, packing == 3 ? 2 : 1);
    }
  }

  private static void pack(MediaInput in, long expected, int unit) {
    long output = 0;
    while (in.remaining() > 0) {
      int code = in.s8();
      if (code == -128) continue;
      long count = (long) (code >= 0 ? code + 1 : 1 - code) * unit;
      in.require(count <= expected - output, MediaStructureReason.COUNT);
      in.skip(code >= 0 ? count : unit);
      output += count;
    }
    in.require(output == expected, MediaStructureReason.COUNT);
  }
}
