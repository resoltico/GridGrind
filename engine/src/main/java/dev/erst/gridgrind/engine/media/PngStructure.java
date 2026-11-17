package dev.erst.gridgrind.engine.media;

import dev.erst.gridgrind.core.problem.MediaStructureReason;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.CRC32;

/** PNG/APNG chunk, integrity and bounded filtered-coding structure. */
final class PngStructure {
  private final MediaInput input;
  private final FrameSequence animation = new FrameSequence();

  /**
   * Active default/APNG frame and the declared animation sequence, independent of chunk CRC state.
   */
  private static final class FrameSequence {
    long declared = -1;
    long count;
    long sequence;
    long width;
    long height;
    boolean isDefault;
  }

  private final List<MediaInput> data = new ArrayList<>();
  private boolean header;
  private boolean end;
  private boolean palette;
  private boolean idat;
  private boolean idatEnded;
  private long width;
  private long height;
  private int bits;
  private int color;
  private int interlace;

  private PngStructure(MediaInput input) {
    this.input = input;
  }

  static ImageGeometry check(MediaInput input) {
    return new PngStructure(input).read();
  }

  private ImageGeometry read() {
    input.magic(137, 80, 78, 71, 13, 10, 26, 10);
    var crc = new CRC32();
    while (input.remaining() > 0 && !end) {
      input.record();
      long count = input.u32();
      int type = (int) input.u32();
      var chunk = input.take(count);
      long expectedCrc = input.u32();
      crc.reset();
      input.budget.inspect(count + 4);
      crc.update(input.bytes, chunk.start - 4, (int) count + 4);
      input.require(crc.getValue() == expectedCrc, MediaStructureReason.CHECKSUM);
      chunkType(type);
      input.require(header || type == 0x49484452, MediaStructureReason.SEQUENCE);
      if (type != 0x49444154 && idat) idatEnded = true;
      chunk(type, chunk, count);
    }
    input.require(header && end && input.remaining() == 0, MediaStructureReason.TERMINATOR);
    input.require(
        animation.declared < 0 || animation.count == animation.declared,
        MediaStructureReason.COUNT);
    input.require(!data.isEmpty(), MediaStructureReason.TRUNCATED);
    decode(input, data, animation.width, animation.height, bits, color, interlace);
    return new ImageGeometry(width, height);
  }

  private void chunkType(int type) {
    for (int shift = 24; shift >= 0; shift -= 8) {
      int letter = (type >>> shift) & 255;
      input.require(
          (letter >= 65 && letter <= 90) || (letter >= 97 && letter <= 122),
          MediaStructureReason.HEADER);
    }
    input.require((type & 0x2000) == 0, MediaStructureReason.HEADER);
  }

  private void chunk(int type, MediaInput chunk, long count) {
    switch (type) {
      case 0x49484452 -> header(chunk, count);
      case 0x504c5445 -> palette(count);
      case 0x49444154 -> {
        input.require(!idatEnded && (color != 3 || palette), MediaStructureReason.SEQUENCE);
        idat = true;
        data.add(chunk);
      }
      case 0x6163544c -> {
        input.require(animation.declared < 0 && !idat && count == 8, MediaStructureReason.SEQUENCE);
        animation.declared = chunk.u32();
        input.require(animation.declared > 0, MediaStructureReason.COUNT);
        chunk.u32();
      }
      case 0x6663544c -> frameControl(chunk, count);
      case 0x66644154 -> {
        input.require(
            animation.declared > 0
                && idat
                && animation.count > 0
                && !animation.isDefault
                && count >= 4,
            MediaStructureReason.SEQUENCE);
        nextSequence(chunk);
        data.add(chunk.take(chunk.remaining()));
      }
      case 0x49454e44 -> {
        input.require(count == 0 && idat && (color != 3 || palette), MediaStructureReason.SEQUENCE);
        end = true;
      }
      case 0x7a545874, 0x69434350 -> compressedText(chunk, false);
      case 0x69545874 -> compressedText(chunk, true);
      default -> {
        if ((type & 0x20000000) == 0)
          input.profile("PNG/critical-chunk-" + Integer.toHexString(type));
      }
    }
  }

  private void header(MediaInput chunk, long count) {
    input.require(!header && count == 13, MediaStructureReason.HEADER);
    header = true;
    width = chunk.u32();
    height = chunk.u32();
    bits = chunk.u8();
    color = chunk.u8();
    dimensions();
    boolean valid = validLayout();
    input.require(valid && chunk.u8() == 0 && chunk.u8() == 0, MediaStructureReason.HEADER);
    interlace = chunk.u8();
    input.require(interlace <= 1, MediaStructureReason.HEADER);
    input.budget.raster(width, height);
    animation.width = width;
    animation.height = height;
  }

  private void dimensions() {
    input.require(
        width > 0 && height > 0 && width <= Integer.MAX_VALUE && height <= Integer.MAX_VALUE,
        MediaStructureReason.DIMENSIONS);
  }

  private boolean validLayout() {
    return switch (color) {
      case 0 -> bits == 1 || bits == 2 || bits == 4 || bits == 8 || bits == 16;
      case 2, 4, 6 -> bits == 8 || bits == 16;
      case 3 -> bits == 1 || bits == 2 || bits == 4 || bits == 8;
      default -> false;
    };
  }

  private void palette(long count) {
    input.require(
        !palette
            && !idat
            && color != 0
            && color != 4
            && count > 0
            && count <= 768
            && count % 3 == 0,
        MediaStructureReason.SEQUENCE);
    if (color == 3) input.require(count / 3 <= 1L << bits, MediaStructureReason.COUNT);
    palette = true;
  }

  private void nextSequence(MediaInput chunk) {
    input.require(chunk.u32() == animation.sequence, MediaStructureReason.SEQUENCE);
    animation.sequence++;
  }

  private void frameControl(MediaInput chunk, long count) {
    input.require(animation.declared > 0 && count == 26, MediaStructureReason.SEQUENCE);
    input.require(animation.count == 0 || !data.isEmpty(), MediaStructureReason.TRUNCATED);
    if (!data.isEmpty()) {
      decode(input, data, animation.width, animation.height, bits, color, interlace);
      data.clear();
    }
    nextSequence(chunk);
    frameRectangle(chunk);
    chunk.u16();
    chunk.u16();
    int dispose = chunk.u8();
    int blend = chunk.u8();
    input.require(dispose <= 2 && blend <= 1, MediaStructureReason.CODE);
    if (!animation.isDefault) input.budget.raster(animation.width, animation.height);
    animation.count++;
  }

  private void frameRectangle(MediaInput chunk) {
    animation.width = chunk.u32();
    animation.height = chunk.u32();
    long x = chunk.u32();
    long y = chunk.u32();
    input.require(
        animation.width > 0
            && animation.height > 0
            && x <= width - animation.width
            && y <= height - animation.height,
        MediaStructureReason.DIMENSIONS);
    animation.isDefault = !idat;
    input.require(
        !animation.isDefault
            || (animation.count == 0
                && x == 0
                && y == 0
                && animation.width == width
                && animation.height == height),
        MediaStructureReason.SEQUENCE);
  }

  private static void compressedText(MediaInput c, boolean international) {
    int keyword = 0;
    while (c.u8() != 0) {
      keyword++;
      c.require(keyword <= 79, MediaStructureReason.COUNT);
    }
    c.require(keyword > 0, MediaStructureReason.COUNT);
    if (international) {
      int compressed = c.u8();
      c.require(compressed <= 1 && c.u8() == 0, MediaStructureReason.CODE);
      while (c.u8() != 0) {}
      while (c.u8() != 0) {}
      if (compressed == 0) return;
    } else c.require(c.u8() == 0, MediaStructureReason.CODE);
    MediaStreams.inflate(List.of(c.take(c.remaining())), -1);
  }

  private static void decode(
      MediaInput owner,
      List<MediaInput> parts,
      long width,
      long height,
      int bits,
      int color,
      int interlace) {
    int channels =
        switch (color) {
          case 0, 3 -> 1;
          case 2 -> 3;
          case 4 -> 2;
          case 6 -> 4;
          default -> throw new IllegalStateException();
        };
    long[] rows = new long[7];
    long[] strides = new long[7];
    long expected = 0;
    List<Pass> passes = interlace == 0 ? PLAIN : ADAM7;
    for (int p = 0; p < passes.size(); p++) {
      var pass = passes.get(p);
      if (width <= pass.x || height <= pass.y) continue;
      long w = (width - pass.x + pass.dx - 1) / pass.dx;
      long h = (height - pass.y + pass.dy - 1) / pass.dy;
      rows[p] = h;
      strides[p] = (owner.product(owner.product(w, channels), bits) + 7) / 8 + 1;
      expected += owner.product(h, strides[p]);
    }
    owner.budget.expand(expected);
    final long total = expected;
    var scan = new Scanlines(owner, rows, strides);
    MediaStreams.inflatePng(parts, total, scan::accept);
    owner.require(scan.complete(), MediaStructureReason.COUNT);
  }

  /** Normative Adam7 pass coordinates, independent of the image payload. */
  private record Pass(int x, int y, int dx, int dy) {}

  private static final List<Pass> PLAIN = List.of(new Pass(0, 0, 1, 1));
  private static final List<Pass> ADAM7 =
      List.of(
          new Pass(0, 0, 8, 8),
          new Pass(4, 0, 8, 8),
          new Pass(0, 4, 4, 8),
          new Pass(2, 0, 4, 4),
          new Pass(0, 2, 2, 4),
          new Pass(1, 0, 2, 2),
          new Pass(0, 1, 1, 2));

  /** Tracks filtered scanline boundaries without retaining inflated raster bytes. */
  private static final class Scanlines {
    private final MediaInput owner;
    private final long[] rows;
    private final long[] strides;
    private int pass;
    private long next;
    private long row;

    Scanlines(MediaInput owner, long[] rows, long... strides) {
      this.owner = owner;
      this.rows = rows.clone();
      this.strides = strides.clone();
      advance();
    }

    private void advance() {
      while (pass < rows.length && rows[pass] == 0) pass++;
    }

    void accept(byte[] data, int length, long offset) {
      while (pass < rows.length && next < offset + length) {
        owner.require(
            next >= offset && (data[(int) (next - offset)] & 255) <= 4, MediaStructureReason.CODE);
        next += strides[pass];
        row++;
        if (row == rows[pass]) {
          pass++;
          row = 0;
          advance();
        }
      }
    }

    boolean complete() {
      return pass == rows.length;
    }
  }
}
