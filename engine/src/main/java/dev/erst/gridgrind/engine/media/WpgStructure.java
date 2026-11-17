package dev.erst.gridgrind.engine.media;

import dev.erst.gridgrind.core.problem.MediaStructureReason;
import java.util.ArrayDeque;
import java.util.Deque;

/** Unencrypted WPG1/2 framed records and counted raster coding, without a painter. */
final class WpgStructure {
  private WpgStructure() {}

  static ImageGeometry check(MediaInput in) {
    in.little = true;
    in.magic(255, 87, 80, 67);
    long start = in.u32();
    in.require(in.u8() == 1 && in.u8() == 22, MediaStructureReason.HEADER);
    int version = in.u8();
    int minor = in.u8();
    int key = in.u16();
    in.u16();
    if ((version != 1 && version != 2) || minor != 0 || key != 0)
      in.profile("WPG/header-version-or-encryption");
    in.require(start >= 16, MediaStructureReason.OFFSET);
    in.seek(start);
    var reader = new Records(in, version);
    reader.read();
    return ImageGeometry.vector();
  }

  private static long variable(MediaInput in) {
    int first = in.u8();
    if (first < 255) return first;
    int next = in.u16();
    return (next & 32768) == 0 ? next : ((long) (next & 32767) << 16) | in.u16();
  }

  /** One WPG document's framing, precision and bounded subgroup state. */
  private static final class Records {
    final MediaInput input;
    final int version;
    final Deque<Long> groups = new ArrayDeque<>();
    boolean begun;
    boolean ended;
    int coordinateBytes = 2;
    long groupEnd = -1;

    Records(MediaInput input, int version) {
      this.input = input;
      this.version = version;
    }

    void read() {
      while (input.remaining() > 0 && !ended) {
        input.record();
        beginRecord();
        int type = input.u8();
        input.require(type > 0 && (version == 1 || type <= 63), MediaStructureReason.CODE);
        long extension = version == 2 ? variable(input) : 0;
        long length = variable(input);
        var record = input.take(length);
        countParent();
        if (version == 1) one(type, record);
        else two(type, record);
        children(extension);
      }
      input.require(
          begun
              && ended
              && input.remaining() == 0
              && groups.isEmpty()
              && (groupEnd < 0 || groupEnd == input.position),
          MediaStructureReason.TERMINATOR);
    }

    private void beginRecord() {
      if (groupEnd >= 0) {
        input.require(input.position <= groupEnd, MediaStructureReason.OFFSET);
        if (input.position == groupEnd) groupEnd = -1;
      }
      if (version == 2) input.require(input.u8() == 0, MediaStructureReason.HEADER);
    }

    private void countParent() {
      if (version == 2 && !groups.isEmpty()) {
        long remaining = groups.pop();
        input.require(remaining > 0, MediaStructureReason.COUNT);
        if (remaining > 1) groups.push(remaining - 1);
      }
    }

    private void children(long extension) {
      if (extension > 0) {
        input.require(!ended, MediaStructureReason.SEQUENCE);
        groups.push(extension);
        input.budget.depth(input.depth + groups.size());
      }
    }

    private void one(int type, MediaInput r) {
      if (type == 15) {
        input.require(!begun, MediaStructureReason.SEQUENCE);
        r.skip(6);
        begun = true;
        return;
      }
      input.require(begun, MediaStructureReason.SEQUENCE);
      switch (type) {
        case 16 -> {
          input.require(groupEnd < 0 || groupEnd == r.start, MediaStructureReason.SEQUENCE);
          ended = true;
        }
        case 1 -> r.skip(2);
        case 2 -> r.skip(4);
        case 4, 6, 8 -> {
          long points = r.u16();
          r.skip(r.product(points, 4));
        }
        case 5, 7 -> r.skip(8);
        case 9 -> r.skip(10);
        case 11, 20 -> {
          if (type == 20) r.skip(10);
          bitmapOne(r);
        }
        case 12 -> {
          int count = r.u16();
          r.skip(4);
          r.skip(count);
        }
        case 13 -> r.skip(20);
        case 14 -> {
          long start = r.u16();
          long count = r.u16();
          r.require(start + count <= 256, MediaStructureReason.COUNT);
          r.skip(count * 3);
        }
        case 17 -> r.skip(8);
        case 19 -> {
          r.skip(4);
          long points = r.u16();
          r.skip(r.product(points, 4));
        }
        case 21 -> {
          input.require(groupEnd < 0, MediaStructureReason.SEQUENCE);
          long length = r.u32();
          groupEnd = r.position + length;
          input.region(groupEnd - input.start, 0);
        }
        // Sized application, font and reserved records remain opaque; their framing is still
        // checked.
        default -> r.skip(r.remaining());
      }
    }

    private void two(int type, MediaInput r) {
      if (type == 1) {
        input.require(!begun, MediaStructureReason.SEQUENCE);
        r.u16();
        r.u16();
        int precision = r.u8();
        r.require(precision <= 1, MediaStructureReason.CODE);
        coordinateBytes = precision == 1 ? 4 : 2;
        r.skip(8L * coordinateBytes);
        begun = true;
        return;
      }
      input.require(begun, MediaStructureReason.SEQUENCE);
      switch (type) {
        case 2 -> {
          input.require(groups.isEmpty(), MediaStructureReason.SEQUENCE);
          ended = true;
        }
        case 12, 13 -> {
          long start = r.u16();
          long count = r.u16();
          r.require(start + count <= 65536, MediaStructureReason.COUNT);
          r.skip(count * (type == 12 ? 4L : 8L));
        }
        case 14 -> bitmapTwo(r);
        case 21, 22, 23 -> {
          characterization(r);
          long count = r.u16();
          r.skip(r.product(count, (type == 23 ? 6L : 2L) * coordinateBytes));
        }
        case 24 -> {
          characterization(r);
          r.skip(6L * coordinateBytes);
        }
        case 25 -> {
          characterization(r);
          r.skip(8L * coordinateBytes);
        }
        case 26 -> characterization(r);
        case 27 -> {
          characterization(r);
          r.skip(6L * coordinateBytes);
        }
        case 37, 39, 49, 51 -> r.skip(4);
        case 38, 40, 50, 52 -> r.skip(coordinateBytes == 4 ? 8 : 4);
        case 41, 45, 46, 53 -> r.skip(2);
        case 43 -> r.skip(4);
        case 44 -> r.skip(coordinateBytes == 4 ? 8 : 4);
        default -> r.skip(r.remaining());
      }
    }

    private static void characterization(MediaInput r) {
      int flags = r.u16();
      if ((flags & 128) != 0) r.skip(4);
      if ((flags & 32) != 0) {
        int id = r.u16();
        if ((id & 32768) != 0) r.u16();
      }
      if ((flags & 16) != 0) r.skip(4);
      if ((flags & 24) != 0) r.skip(8);
      if ((flags & 20) != 0) r.skip(8);
      if ((flags & 2) != 0) r.skip(12);
      if ((flags & 1) != 0) r.skip(8);
    }
  }

  private static void bitmapOne(MediaInput r) {
    long width = r.u16();
    long height = r.u16();
    int bits = r.u16();
    r.skip(4);
    r.require(bits == 1 || bits == 2 || bits == 4 || bits == 8, MediaStructureReason.CODE);
    long row = (r.product(width, bits) + 7) / 8;
    long expected = r.raster(width, height, row);
    long output = 0;
    while (r.remaining() > 0) {
      int code = r.u8();
      long count;
      if ((code & 128) != 0) {
        count = code & 127;
        if (count > 0) r.u8();
        else count = r.u8();
      } else if (code > 0) {
        count = code;
        r.skip(count);
      } else {
        r.require(output >= row && output % row == 0, MediaStructureReason.SEQUENCE);
        count = r.product(r.u8(), row);
      }
      r.require(count <= expected - output, MediaStructureReason.COUNT);
      output += count;
    }
    r.require(output == expected, MediaStructureReason.TRUNCATED);
  }

  private static int sampleBits(int format) {
    return switch (format) {
      case 1 -> 1;
      case 2 -> 2;
      case 3 -> 4;
      case 4, 12 -> 8;
      case 5, 6, 7 -> 16;
      case 8 -> 24;
      case 9 -> 32;
      case 10 -> 48;
      case 11 -> 64;
      default -> 0;
    };
  }

  private static void bitmapTwo(MediaInput r) {
    long width = r.u16();
    long height = r.u16();
    int color = r.u8();
    int compression = r.u8();
    int bits = sampleBits(color);
    r.require(bits > 0, MediaStructureReason.CODE);
    long row = (r.product(width, bits) + 7) / 8;
    long expected = r.raster(width, height, row);
    if (compression == 0) {
      r.require(r.remaining() == expected, MediaStructureReason.LENGTH);
      r.skip(expected);
      return;
    }
    if (compression != 1) r.profile("WPG/bitmap-compression-" + compression);
    long output = 0;
    long xorEnd = -1;
    int unit = 1;
    while (r.remaining() > 0) {
      int code = r.u8();
      long count = 0;
      switch (code) {
        case 125 -> {
          unit = r.u8();
          r.require(unit > 0, MediaStructureReason.CODE);
        }
        case 126 -> {
          r.require(xorEnd < 0 && output >= row, MediaStructureReason.SEQUENCE);
          xorEnd = output + row;
          r.require(xorEnd <= expected, MediaStructureReason.COUNT);
        }
        case 127, 255, 253 -> count = r.product(1L + r.u8(), unit);
        case 254 -> {
          r.require(output >= row && output % row == 0, MediaStructureReason.SEQUENCE);
          count = r.product(1L + r.u8(), row);
        }
        default -> {
          count = r.product(1L + (code & 127), unit);
          r.skip(code >= 128 ? unit : count);
        }
      }
      r.require(count <= expected - output, MediaStructureReason.COUNT);
      output += count;
      if (xorEnd >= 0 && output >= xorEnd) xorEnd = -1;
    }
    r.require(output == expected && xorEnd < 0, MediaStructureReason.TRUNCATED);
  }
}
