package dev.erst.gridgrind.engine.media;

import dev.erst.gridgrind.core.problem.MediaStructureReason;
import java.util.BitSet;

/** MS-EMF/MS-WMF sized record structure, scalar/span checks and declared bitmap geometry. */
final class MetafileStructure {
  private MetafileStructure() {}

  /** File/control declarations shared by the EMF header and final record. */
  private record Header(long records, long palettes) {}

  static ImageGeometry emf(MediaInput in) {
    Header header = emfHeader(in);
    long count = 1;
    boolean ended = false;
    while (in.remaining() > 0) {
      var record = emfRecord(in);
      long type = record.u32();
      record.u32();
      count++;
      in.require(type != 1, MediaStructureReason.SEQUENCE);
      if (type == 14) {
        eof(record, header.palettes);
        in.require(in.remaining() == 0, MediaStructureReason.TERMINATOR);
        ended = true;
        break;
      }
      emfContent(record, type);
    }
    in.require(ended && count == header.records, MediaStructureReason.COUNT);
    return ImageGeometry.vector();
  }

  private static Header emfHeader(MediaInput in) {
    in.little = true;
    var header = emfRecord(in);
    header.require(header.u32() == 1, MediaStructureReason.HEADER);
    header.u32();
    header.need(80);
    header.seek(40);
    header.require(
        header.u32() == 0x464d4520 && header.u32() == 0x10000, MediaStructureReason.HEADER);
    long bytes = header.u32();
    long records = header.u32();
    int handles = header.u16();
    header.u16();
    long description = header.u32();
    long descriptionAt = header.u32();
    long palettes = header.u32();
    header.require(bytes == in.size() && records >= 3 && handles > 0, MediaStructureReason.COUNT);
    long fixed = header.size();
    if (description > 0 && descriptionAt > 0) {
      header.require(descriptionAt >= 88, MediaStructureReason.OFFSET);
      // The declared extent is bounded within this header; description text stays opaque.
      header.region(descriptionAt, header.product(description, 2));
      fixed = Math.min(fixed, descriptionAt);
    }
    if (fixed >= 100) {
      var extension = header.region(88, 12);
      long size = extension.u32();
      long offset = extension.u32();
      extension.require(extension.u32() <= 1, MediaStructureReason.CODE);
      if (size > 0) {
        header.require(offset >= 100, MediaStructureReason.OFFSET);
        header.region(offset, size);
        fixed = Math.min(fixed, offset);
      }
    }
    header.require(fixed >= 88, MediaStructureReason.LENGTH);
    return new Header(records, palettes);
  }

  private static void eof(MediaInput record, long palettes) {
    long entries = record.u32();
    long offset = record.u32();
    record.require(entries == palettes && record.size() >= 20, MediaStructureReason.COUNT);
    if (entries > 0) {
      record.require(
          offset >= 16 && record.product(entries, 4) <= record.size() - 4L - offset,
          MediaStructureReason.OFFSET);
      record.region(offset, entries * 4);
    }
    record.seek(record.size() - 4L);
    record.require(record.u32() == record.size(), MediaStructureReason.TERMINATOR);
  }

  private static MediaInput emfRecord(MediaInput in) {
    in.record();
    var prefix = in.region(in.offset(), 8);
    long type = prefix.u32();
    long length = prefix.u32();
    in.require(type > 0 && length >= 8 && length % 4 == 0, MediaStructureReason.LENGTH);
    in.require(type <= 122 && type != 69 && type != 107 && type != 117, MediaStructureReason.CODE);
    in.require(length >= emfMinimum(type), MediaStructureReason.TRUNCATED);
    return in.take(length);
  }

  private static int emfMinimum(long type) {
    return switch ((int) type) {
      case 1 -> 88;
      case 14 -> 20;
      case 2, 3, 4, 5, 6, 56, 85, 86, 87, 88, 89, 92 -> 28;
      case 7, 8, 90, 91 -> 32;
      case 9, 10, 11, 12, 13, 26, 27, 51, 54, 75, 105, 106, 120 -> 16;
      case 15, 50, 109, 110, 112, 113 -> 20;
      case 23 -> 32;
      case 29, 30, 31, 32, 39, 42, 43, 53, 62, 63, 64, 111, 121 -> 24;
      case 35, 44, 71, 93, 94 -> 32;
      case 36, 108, 118 -> 36;
      case 38, 41, 73, 74, 103 -> 28;
      case 45, 46, 47, 55, 72, 96, 97 -> 40;
      case 49 -> 16;
      case 76 -> 100;
      case 77, 114, 116 -> 108;
      case 78 -> 128;
      case 79 -> 140;
      case 80, 83, 84 -> 76;
      case 81 -> 80;
      case 82 -> 332;
      case 95 -> 52;
      case 99 -> 340;
      case 104 -> 48;
      case 122 -> 608;
      case 28, 33, 52, 59, 60, 61, 65, 66, 68 -> 8;
      default -> 12;
    };
  }

  private static void emfContent(MediaInput r, long type) {
    if ((type >= 2 && type <= 8) || type == 56 || (type >= 85 && type <= 92)) poly(r, type);
    else
      switch ((int) type) {
        case 49 -> {
          r.seek(12);
          r.require(r.u16() == 0x300, MediaStructureReason.HEADER);
          r.skip(r.u16() * 4L);
        }
        case 50 -> {
          r.seek(16);
          r.skip(r.product(r.u32(), 4));
        }
        case 70, 102, 103, 105, 106, 110, 112, 113, 119, 121, 122 -> sizedData(r, (int) type);
        case 71, 72, 73, 74, 75 -> region(r, (int) type);
        case 83, 84, 96, 97, 108 -> textRecord(r, (int) type);
        case 118 -> gradient(r);
        default -> {}
      }
    bitmaps(r, (int) type);
  }

  private static void poly(MediaInput r, long type) {
    r.seek(24);
    if (type == 7 || type == 8 || type == 90 || type == 91) {
      long polygons = r.u32();
      long points = r.u32();
      long total = 0;
      for (long i = 0; i < polygons; i++) {
        r.record();
        long n = r.u32();
        r.require(n <= points - total, MediaStructureReason.COUNT);
        total += n;
      }
      r.require(total == points, MediaStructureReason.COUNT);
      r.skip(r.product(points, type >= 90 ? 4 : 8));
      return;
    }
    long points = r.u32();
    r.skip(r.product(points, type >= 85 ? 4 : 8));
    if (type == 56 || type == 92) r.skip(points);
  }

  private static void region(MediaInput r, int type) {
    r.seek(type == 75 ? 8 : 24);
    long bytes = r.u32();
    r.seek(
        switch (type) {
          case 71 -> 32;
          case 72 -> 40;
          case 75 -> 16;
          default -> 28;
        });
    if (bytes > 0) regionData(r.take(bytes));
  }

  private static void textRecord(MediaInput r, int type) {
    switch (type) {
      case 83, 84 -> text(r, 36, type == 83 ? 1 : 2);
      case 96, 97 -> {
        r.seek(36);
        long count = r.u32();
        r.need(r.product(count, 40));
        for (long i = 0; i < count; i++) {
          r.record();
          text(r, 40 + i * 40, type == 96 ? 1 : 2);
        }
      }
      case 108 -> {
        r.seek(16);
        long count = r.u32();
        long flags = r.u32();
        r.seek((flags & 256) != 0 ? 36 : 52);
        r.skip(r.product(count, (flags & 512) != 0 ? 1 : 2));
      }
      default -> throw new IllegalStateException("Unknown EMF text record");
    }
  }

  private static void sizedData(MediaInput r, int type) {
    switch (type) {
      case 70, 102, 103, 105, 106 -> {
        r.seek(
            switch (type) {
              case 103 -> 24;
              case 105, 106 -> 12;
              default -> 8;
            });
        r.skip(r.u32());
      }
      case 110 -> {
        r.seek(8);
        long driver = r.u32();
        r.u32();
        long bytes = r.u32();
        r.skip(driver + bytes);
      }
      case 112, 113, 121 -> {
        r.seek(type == 121 ? 16 : 12);
        long name = r.u32();
        long data = r.u32();
        r.skip(name + data);
      }
      case 119 -> {
        r.seek(8);
        r.skip(r.product(r.u32(), 8));
      }
      case 122 -> {
        r.seek(604);
        r.skip(r.u32());
      }
      default -> throw new IllegalStateException("Unknown EMF sized record");
    }
  }

  private static void gradient(MediaInput r) {
    r.seek(24);
    long vertices = r.u32();
    long meshes = r.u32();
    long mode = r.u32();
    r.require(mode <= 2, MediaStructureReason.CODE);
    r.skip(r.product(vertices, 16));
    long indices = r.product(meshes, mode == 2 ? 3 : 2);
    for (long i = 0; i < indices; i++) r.require(r.u32() < vertices, MediaStructureReason.OFFSET);
  }

  private static void bitmaps(MediaInput r, int type) {
    switch (type) {
      case 76, 77, 114, 116 -> bitmapOffsets(r, 84, -1, 0, r.region(80, 4).u32());
      case 78 -> {
        bitmapOffsets(r, 84, -1, 0, r.region(80, 4).u32());
        bitmapOffsets(r, 112, -1, 0, r.region(108, 4).u32());
      }
      case 79 -> {
        bitmapOffsets(r, 96, -1, 0, r.region(92, 4).u32());
        bitmapOffsets(r, 124, -1, 0, r.region(120, 4).u32());
      }
      case 80 ->
          bitmapOffsets(r, 48, r.region(72, 4).u32(), r.region(68, 4).u32(), r.region(64, 4).u32());
      case 81 -> bitmapOffsets(r, 48, -1, 0, r.region(64, 4).u32());
      case 93, 94 -> bitmapOffsets(r, 16, -1, 0, r.region(12, 4).u32());
      case 95 -> {
        bitmapOffsets(r, 12, -1, 0, r.region(44, 4).u32());
        r.seek(48);
        r.skip(r.product(r.u32(), 4));
      }
      default -> {}
    }
  }

  private static void text(MediaInput r, long at, int bytesPerCharacter) {
    var text = r.region(at, 40);
    text.skip(8);
    long count = text.u32();
    long offset = text.u32();
    long flags = text.u32();
    text.skip(16);
    long advances = text.u32();
    if (count > 0) {
      r.region(offset, r.product(count, bytesPerCharacter));
      if (advances > 0) r.region(advances, r.product(count, (flags & 8192) != 0 ? 8 : 4));
    }
  }

  private static void regionData(MediaInput r) {
    r.require(r.u32() == 32 && r.u32() == 1, MediaStructureReason.HEADER);
    long count = r.u32();
    long bytes = r.u32();
    r.skip(16);
    r.require(bytes == r.product(count, 16), MediaStructureReason.COUNT);
    r.skip(bytes);
  }

  private static void bitmapOffsets(
      MediaInput r, int at, long selectedRows, long firstScan, long colorUsage) {
    var fields = r.region(at, 16);
    long infoAt = fields.u32();
    long infoLength = fields.u32();
    long bitsAt = fields.u32();
    long bitsLength = fields.u32();
    if (infoLength == 0 && bitsLength == 0) return;
    r.require(infoLength > 0 && (bitsLength > 0 || selectedRows == 0), MediaStructureReason.LENGTH);
    r.require(colorUsage <= 2, MediaStructureReason.CODE);
    BitmapStructure.checkParts(
        r.region(infoAt, infoLength),
        r.region(bitsAt, bitsLength),
        selectedRows,
        firstScan,
        (int) colorUsage);
  }

  /** Standard WMF file declarations, excluding the optional placeable prefix. */
  private record WmfHeader(int objects, long maximum) {}

  static ImageGeometry wmf(MediaInput in) {
    in.little = true;
    placeable(in);
    WmfHeader header = wmfHeader(in);
    var objects = new ObjectTable(header.objects);
    long largest = 0;
    boolean ended = false;
    while (in.remaining() > 0) {
      in.record();
      var prefix = in.region(in.offset(), 6);
      long size = prefix.u32();
      int encoded = prefix.u16();
      int function = wmfFunction(encoded);
      in.require(size >= 3 && size <= Integer.MAX_VALUE / 2, MediaStructureReason.LENGTH);
      largest = Math.max(largest, size);
      var record = in.take(size * 2);
      record.seek(6);
      if (function == 0) {
        record.require(size == 3 && in.remaining() == 0, MediaStructureReason.TERMINATOR);
        ended = true;
        break;
      }
      wmfContent(record, function, encoded);
      objects.record(record, function);
    }
    in.require(ended && largest == header.maximum, MediaStructureReason.COUNT);
    return ImageGeometry.vector();
  }

  private static void placeable(MediaInput in) {
    if (in.region(0, 4).u32() != 0x9ac6cdd7L) return;
    var header = in.take(22);
    int checksum = 0;
    for (int i = 0; i < 10; i++) checksum ^= header.u16();
    header.require(header.u16() == checksum, MediaStructureReason.CHECKSUM);
    header.seek(14);
    header.require(header.u16() > 0, MediaStructureReason.DIMENSIONS);
  }

  private static WmfHeader wmfHeader(MediaInput in) {
    int start = in.offset();
    int kind = in.u16();
    in.require((kind == 1 || kind == 2) && in.u16() == 9, MediaStructureReason.HEADER);
    int version = in.u16();
    in.require(version == 0x100 || version == 0x300, MediaStructureReason.HEADER);
    long words = in.u32();
    int objects = in.u16();
    long maximum = in.u32();
    in.u16();
    in.require(words * 2 == in.size() - start && maximum >= 3, MediaStructureReason.LENGTH);
    return new WmfHeader(objects, maximum);
  }

  /** Bounded declared WMF handle slots and their creation/deletion history within one stream. */
  private static final class ObjectTable {
    private final int count;
    private final BitSet present;

    ObjectTable(int count) {
      this.count = count;
      present = new BitSet(count);
    }

    void record(MediaInput record, int function) {
      if (function == 0x01f0) {
        int index = record.region(6, 2).u16();
        record.require(index < count && present.get(index), MediaStructureReason.OFFSET);
        present.clear(index);
      } else if (function == 0x012d || function == 0x012c || function == 0x0234) {
        int index = record.region(6, 2).u16();
        if ((index & 0x8000) == 0)
          record.require(index < count && present.get(index), MediaStructureReason.OFFSET);
      } else if (created(function)) {
        int index = present.nextClearBit(0);
        record.require(index < count, MediaStructureReason.COUNT);
        present.set(index);
      }
    }

    private static boolean created(int function) {
      return switch (function) {
        case 0x02fa, 0x02fb, 0x02fc, 0x00f7, 0x01f9, 0x0142, 0x06ff -> true;
        default -> false;
      };
    }
  }

  private static int wmfFunction(int encoded) {
    return switch (encoded & 255) {
      case 0x0 -> 0x0;
      case 0x36 -> 0x436;
      case 0x17 -> 0x817;
      case 0x22 -> 0x922;
      case 0x30 -> 0x830;
      case 0xfc -> 0x2fc;
      case 0xfb -> 0x2fb;
      case 0xf7 -> 0xf7;
      case 0xf9 -> 0x1f9;
      case 0xfa -> 0x2fa;
      case 0xff -> 0x6ff;
      case 0xf0 -> 0x1f0;
      case 0x40 -> 0x940;
      case 0x42 -> 0x142;
      case 0x41 -> 0xb41;
      case 0x18 -> 0x418;
      case 0x26 -> 0x626;
      case 0x15 -> 0x415;
      case 0x48 -> 0x548;
      case 0x32 -> 0xa32;
      case 0x28 -> 0x228;
      case 0x19 -> 0x419;
      case 0x29 -> 0x429;
      case 0x16 -> 0x416;
      case 0x2a -> 0x12a;
      case 0x13 -> 0x213;
      case 0x14 -> 0x214;
      case 0x20 -> 0x220;
      case 0x11 -> 0x211;
      case 0xf -> 0x20f;
      case 0x2b -> 0x12b;
      case 0x1d -> 0x61d;
      case 0x1a -> 0x81a;
      case 0x24 -> 0x324;
      case 0x25 -> 0x325;
      case 0x38 -> 0x538;
      case 0x35 -> 0x35;
      case 0x1b -> 0x41b;
      case 0x39 -> 0x139;
      case 0x27 -> 0x127;
      case 0x1c -> 0x61c;
      case 0x1e -> 0x1e;
      case 0x12 -> 0x412;
      case 0x10 -> 0x410;
      case 0x2c -> 0x12c;
      case 0x2d -> 0x12d;
      case 0x34 -> 0x234;
      case 0x1 -> 0x201;
      case 0x2 -> 0x102;
      case 0x33 -> 0xd33;
      case 0x49 -> 0x149;
      case 0x3 -> 0x103;
      case 0x31 -> 0x231;
      case 0x37 -> 0x37;
      case 0x1f -> 0x41f;
      case 0x6 -> 0x106;
      case 0x5 -> 0x105;
      case 0x4 -> 0x104;
      case 0x7 -> 0x107;
      case 0x2e -> 0x12e;
      case 0x8 -> 0x108;
      case 0x9 -> 0x209;
      case 0xa -> 0x20a;
      case 0xe -> 0x20e;
      case 0xd -> 0x20d;
      case 0xc -> 0x20c;
      case 0xb -> 0x20b;
      case 0x23 -> 0xb23;
      case 0x43 -> 0xf43;
      case 0x21 -> 0x521;
      default -> encoded;
    };
  }

  private static void bitmap16(MediaInput r, boolean brush) {
    r.s16();
    long width = r.s16();
    long height = r.s16();
    r.s16();
    int planes = r.u8();
    int bits = r.u8();
    r.require(width > 0 && height > 0 && planes == 1 && bits > 0, MediaStructureReason.DIMENSIONS);
    if (brush) r.skip(22);
    long row = ((r.product(width, bits) + 15) / 16) * 2;
    r.skip(r.raster(width, height, row));
  }

  private static void wmfRegion(MediaInput r) {

    r.u16();
    r.require(r.u16() == 6, MediaStructureReason.HEADER);
    r.u32();
    r.u16();
    int scans = r.s16();
    int maximum = r.s16();
    r.require(scans >= 0 && maximum >= 0, MediaStructureReason.COUNT);
    r.skip(8);
    for (int i = 0; i < scans; i++) {
      r.record();
      int count = r.u16();
      r.require((count & 1) == 0 && count <= maximum, MediaStructureReason.COUNT);
      r.skip(4L + count * 2L);
      r.require(r.u16() == count, MediaStructureReason.COUNT);
    }
  }

  private static void wmfDib(MediaInput r, int function, int encodedFunction) {
    if ((function == 0x0940 || function == 0x0b41) && r.size() / 2 == (encodedFunction >>> 8) + 3)
      return;
    int prefix =
        switch (function) {
          case 0x0142 -> 4;
          case 0x0940 -> 16;
          case 0x0b41 -> 20;
          case 0x0d33 -> 18;
          default -> 22;
        };
    int usage = wmfColorUsage(r, function);
    long rows = -1;
    long first = 0;
    if (function == 0x0d33) {
      rows = r.region(8, 2).u16();
      first = r.region(10, 2).u16();
    }
    r.seek(6L + prefix);
    if (r.remaining() > 0)
      BitmapStructure.checkDibMember(
          r.nested(
              r.offset(), r.remaining(), dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat.DIB),
          rows,
          first,
          usage);
  }

  private static int wmfColorUsage(MediaInput r, int function) {
    return switch (function) {
      case 0x0142 -> r.region(8, 2).u16();
      case 0x0d33 -> r.region(6, 2).u16();
      case 0x0f43 -> r.region(10, 2).u16();
      default -> 0;
    };
  }

  private static int wmfMinimum(int function) {
    return switch (function) {
      case 0x02fa -> 10;
      case 0x02fb -> 50;
      case 0x02fc -> 8;
      case 0x00f7, 0x0037 -> 4;
      case 0x0324, 0x0325, 0x0538, 0x0521 -> 2;
      case 0x0626 -> 4;
      case 0x0a32 -> 8;
      case 0x0142 -> 4;
      case 0x0234 -> 2;
      case 0x06ff -> 22;
      default -> (function >>> 8) * 2;
    };
  }

  private static void wmfContent(MediaInput r, int function, int encodedFunction) {
    r.need(wmfMinimum(function));
    switch (function) {
      case 0x0324, 0x0325 -> r.skip(r.product(r.u16(), 4));
      case 0x0538 -> {
        long polygons = r.u16();
        long points = 0;
        for (long i = 0; i < polygons; i++) {
          r.record();
          points += r.u16();
        }
        r.skip(r.product(points, 4));
      }
      case 0x0521 -> {
        int chars = r.u16();
        r.skip(chars + (chars & 1) + 4L);
      }
      case 0x0a32 -> {
        r.skip(4);
        int count = r.u16();
        int options = r.u16();
        if ((options & 6) != 0) r.skip(8);
        r.skip(count + (count & 1));
        if (r.remaining() > 0) r.need(r.product(count, (options & 8192) != 0 ? 4 : 2));
      }
      case 0x0626 -> {
        r.u16();
        r.skip(r.u16());
      }
      case 0x00f7, 0x0037 -> {
        r.u16();
        r.skip(r.u16() * 4L);
      }
      case 0x0436 -> {
        r.skip(4);
        r.skip(r.u16() * 4L);
      }
      case 0x06ff -> wmfRegion(r);
      case 0x01f9 -> bitmap16(r, true);
      case 0x0922, 0x0b23 -> {
        if (r.size() / 2 != (encodedFunction >>> 8) + 3) {
          r.seek(function == 0x0922 ? 22 : 26);
          bitmap16(r, false);
        }
      }
      case 0x0142, 0x0940, 0x0b41, 0x0d33, 0x0f43 -> wmfDib(r, function, encodedFunction);
      default -> r.skip(r.remaining());
    }
  }
}
