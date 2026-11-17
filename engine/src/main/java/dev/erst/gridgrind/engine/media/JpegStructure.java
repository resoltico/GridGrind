package dev.erst.gridgrind.engine.media;

import dev.erst.gridgrind.core.problem.MediaStructureReason;
import org.jspecify.annotations.Nullable;

/** T.81 marker/table/frame/scan structure without universal entropy reconstruction. */
final class JpegStructure {
  private JpegStructure() {}

  /** Bounded table-presence context for abbreviated TIFF JPEG members. */
  static final class Tables {
    final boolean[] quant = new boolean[4];
    final int[] quantPrecision = new int[4];
    final boolean[][] huffman = {new boolean[4], new boolean[4]};
  }

  static Tables tables(MediaInput in) {
    Tables tables = new Tables();
    new Parser(in, tables, true, false).read();
    return tables;
  }

  static ImageGeometry check(MediaInput in) {
    return new Parser(in, new Tables(), false, true).read();
  }

  static ImageGeometry member(MediaInput in, Tables external) {
    return new Parser(in, external, false, false).read();
  }

  /** One image stream's frame/scan state; no entropy or raster buffers are retained. */
  private static final class Parser {
    private final MediaInput in;
    private final Tables source;
    private final Tables tables = new Tables();
    private final boolean onlyTables;
    private final boolean charge;
    private final int[] selected = new int[256];
    private final int[] selectors = new int[256];
    private final int[] ids = new int[256];
    private @Nullable Frame frame;
    private long firstWidth;
    private long firstHeight;
    private int outputBits;
    private int outputComponents;
    private int restartInterval;
    private int scanEpoch;
    private int frameCount;
    private boolean hierarchy;
    private int pending = -1;
    private boolean anyScan;
    private boolean ended;

    Parser(MediaInput in, Tables source, boolean onlyTables, boolean charge) {
      this.in = in;
      this.source = source;
      this.onlyTables = onlyTables;
      this.charge = charge;
      System.arraycopy(source.quant, 0, tables.quant, 0, 4);
      System.arraycopy(source.quantPrecision, 0, tables.quantPrecision, 0, 4);
      for (int i = 0; i < 2; i++) System.arraycopy(source.huffman[i], 0, tables.huffman[i], 0, 4);
    }

    ImageGeometry read() {
      in.magic(255, 216);
      while (in.remaining() > 0 || pending >= 0) {
        int marker = nextMarker();
        in.record();
        in.require(
            marker != 0 && marker != 216 && marker != 1 && !(marker >= 208 && marker <= 215),
            MediaStructureReason.CODE);
        if (marker == 217) {
          if (frame != null) complete(in, frame);
          ended = true;
          break;
        }
        int length = in.u16();
        in.require(length >= 2, MediaStructureReason.LENGTH);
        var data = in.take(length - 2);
        marker(marker, data);
      }
      return finish();
    }

    private ImageGeometry finish() {
      in.require(ended && in.remaining() == 0, MediaStructureReason.TERMINATOR);
      in.require(
          onlyTables ? frame == null && !anyScan : anyScan && frame != null && frame.height > 0,
          MediaStructureReason.TRUNCATED);
      if (onlyTables) {
        System.arraycopy(tables.quant, 0, source.quant, 0, 4);
        System.arraycopy(tables.quantPrecision, 0, source.quantPrecision, 0, 4);
        for (int i = 0; i < 2; i++) System.arraycopy(tables.huffman[i], 0, source.huffman[i], 0, 4);
      }
      return new ImageGeometry(firstWidth, firstHeight, outputBits, outputComponents);
    }

    private int nextMarker() {
      if (pending >= 0) {
        int marker = pending;
        pending = -1;
        return marker;
      }
      in.require(in.u8() == 255, MediaStructureReason.SEQUENCE);
      int marker;
      do {
        marker = in.u8();
      } while (marker == 255);
      return marker;
    }

    private void marker(int marker, MediaInput data) {
      if (isFrame(marker)) {
        frame(marker, data);
        return;
      }
      switch (marker) {
        case 219 -> quant(data, tables);
        case 196 -> huffman(data, tables);
        case 204 -> arithmetic(data);
        case 221 -> {
          in.require(data.size() == 2, MediaStructureReason.LENGTH);
          restartInterval = data.u16();
        }
        case 222 -> hierarchical(data);
        case 223 -> {
          in.require(!onlyTables && data.size() == 1, MediaStructureReason.LENGTH);
          int expand = data.u8();
          in.require((expand >> 4) <= 1 && (expand & 15) <= 1, MediaStructureReason.CODE);
        }
        case 218 -> scan(data);
        default -> {
          if (!((marker >= 224 && marker <= 239) || marker == 254))
            in.profile("JPEG/marker-" + marker);
        }
      }
    }

    private void arithmetic(MediaInput data) {
      in.require(data.size() % 2 == 0, MediaStructureReason.LENGTH);
      while (data.remaining() > 0) {
        int code = data.u8();
        int value = data.u8();
        in.require((code >> 4) <= 1 && (code & 15) <= 3, MediaStructureReason.CODE);
        if ((code >> 4) == 0) in.require((value & 15) <= (value >> 4), MediaStructureReason.CODE);
        else in.require(value >= 1 && value <= 63, MediaStructureReason.CODE);
      }
    }

    private void frame(int marker, MediaInput data) {
      beginFrame(marker);
      var next = new Frame(marker);
      next.precision = data.u8();
      next.height = data.u16();
      next.width = data.u16();
      int count = data.u8();
      in.require(
          next.width > 0 && count > 0 && data.remaining() == count * 3,
          MediaStructureReason.DIMENSIONS);
      precision(next);
      in.require(!progressive(next) || count <= 4, MediaStructureReason.COUNT);
      components(data, next, count);
      next.count = count;
      acceptFrame(next);
    }

    private void beginFrame(int marker) {
      in.require(
          !onlyTables && (frame == null || (hierarchy && frame.scanned)),
          MediaStructureReason.SEQUENCE);
      in.require((marker & 4) == 0 || (hierarchy && frameCount > 0), MediaStructureReason.SEQUENCE);
      if (frame != null) complete(in, frame);
    }

    private void precision(Frame next) {
      in.require(
          lossless(next)
              ? next.precision >= 2 && next.precision <= 16
              : next.precision == 8 || next.precision == 12,
          MediaStructureReason.CODE);
      if (next.kind == 192) in.require(next.precision == 8, MediaStructureReason.CODE);
      if (hierarchy) in.require(next.precision == outputBits, MediaStructureReason.CODE);
    }

    private void components(MediaInput data, Frame next, int count) {
      for (int c = 0; c < count; c++) {
        int id = data.u8();
        int sampling = data.u8();
        int quantization = data.u8();
        in.require(
            !next.components[id]
                && (sampling >> 4) >= 1
                && (sampling >> 4) <= 4
                && (sampling & 15) >= 1
                && (sampling & 15) <= 4
                && quantization <= 3,
            MediaStructureReason.CODE);
        next.components[id] = true;
        next.componentOrder[id] = c;
        next.samplingProducts[id] = (sampling >> 4) * (sampling & 15);
        next.quant[id] = quantization;
      }
    }

    private void acceptFrame(Frame next) {
      if (firstWidth == 0) {
        firstWidth = next.width;
        firstHeight = next.height;
        outputBits = next.precision;
        outputComponents = next.count;
      }
      frameCount++;
      if (hierarchy)
        in.require(
            next.width <= firstWidth && (next.height == 0 || next.height <= firstHeight),
            MediaStructureReason.DIMENSIONS);
      if (next.height > 0 && (charge || frameCount > 1)) reserve(in, next);
      frame = next;
    }

    private void hierarchical(MediaInput data) {
      in.require(
          !onlyTables && !hierarchy && frame == null && data.size() >= 6,
          MediaStructureReason.SEQUENCE);
      hierarchy = true;
      int precision = data.u8();
      long height = data.u16();
      long width = data.u16();
      int count = data.u8();
      in.require(
          width > 0
              && height > 0
              && count > 0
              && data.remaining() == count * 3
              && precision >= 2
              && precision <= 16,
          MediaStructureReason.HEADER);
      firstWidth = width;
      firstHeight = height;
      outputBits = precision;
      outputComponents = count;
      hierarchicalComponents(data, count);
      if (charge) {
        in.budget.raster(width, height);
        in.budget.expand(
            in.product((in.product(in.product(width, count), precision) + 7) / 8, height));
      }
    }

    private void hierarchicalComponents(MediaInput data, int count) {
      boolean[] ids = new boolean[256];
      for (int i = 0; i < count; i++) {
        int id = data.u8();
        int sampling = data.u8();
        int q = data.u8();
        in.require(
            !ids[id]
                && (sampling >> 4) >= 1
                && (sampling >> 4) <= 4
                && (sampling & 15) >= 1
                && (sampling & 15) <= 4
                && q == 0,
            MediaStructureReason.CODE);
        ids[id] = true;
      }
    }

    private Frame requiredFrame() {
      var current = frame;
      if (current == null || onlyTables)
        throw MediaIssues.structure(in.format, MediaStructureReason.SEQUENCE, in.position);
      return current;
    }

    private void scan(MediaInput data) {
      Frame current = requiredFrame();
      int count = data.u8();
      in.require(
          count > 0 && count <= 4 && count <= current.count && data.remaining() == count * 2 + 3,
          MediaStructureReason.COUNT);
      scanEpoch++;
      for (int i = 0; i < count; i++) scanComponent(data, current, i);
      interleavedSampling(current, count);
      int start = data.u8();
      int finish = data.u8();
      int approx = data.u8();
      scanShape(current, count, start, finish, approx);
      scanTables(current, count, start, approx);
      current.scanned = true;
      anyScan = true;
      pending = entropy(current);
    }

    private void scanComponent(MediaInput data, Frame current, int index) {

      int id = data.u8();
      int table = data.u8();
      in.require(
          current.components[id]
              && selected[id] != scanEpoch
              && (table >> 4) <= 3
              && (table & 15) <= 3,
          MediaStructureReason.CODE);
      selected[id] = scanEpoch;
      selectors[id] = table;
      in.require(
          index == 0 || current.componentOrder[id] > current.componentOrder[ids[index - 1]],
          MediaStructureReason.SEQUENCE);
      if (current.kind == 192)
        in.require((table >> 4) <= 1 && (table & 15) <= 1, MediaStructureReason.CODE);
      if (lossless(current)) in.require((table & 15) == 0, MediaStructureReason.CODE);
      ids[index] = id;
      if (!lossless(current)) {
        in.require(tables.quant[current.quant[id]], MediaStructureReason.METADATA);
        if (current.precision == 8)
          in.require(tables.quantPrecision[current.quant[id]] == 0, MediaStructureReason.METADATA);
      }
    }

    private void interleavedSampling(Frame current, int count) {
      if (count == 1) return;
      int blocks = 0;
      for (int i = 0; i < count; i++) blocks += current.samplingProducts[ids[i]];
      in.require(blocks <= 10, MediaStructureReason.COUNT);
    }

    private void scanShape(Frame current, int count, int start, int finish, int approx) {
      if (lossless(current)) {
        in.require(
            ((current.kind & 4) != 0 ? start == 0 : start >= 1 && start <= 7)
                && finish == 0
                && (approx >> 4) == 0
                && (approx & 15) < current.precision,
            MediaStructureReason.CODE);
        return;
      }
      if (progressive(current)) {
        in.require(
            start <= finish
                && finish <= 63
                && (start != 0 || finish == 0)
                && (start == 0 || count == 1)
                && (approx & 15) <= 13
                && (approx >> 4) <= 13
                && ((approx >> 4) == 0 || (approx >> 4) == (approx & 15) + 1),
            MediaStructureReason.CODE);
        return;
      }
      in.require(start == 0 && finish == 63 && approx == 0, MediaStructureReason.CODE);
    }

    private void scanTables(Frame current, int count, int start, int approx) {
      for (int i = 0; i < count; i++) {
        int id = ids[i];
        int table = selectors[id];
        if (current.kind < 201) {
          if (lossless(current) || !progressive(current) || (start == 0 && (approx >> 4) == 0))
            in.require(tables.huffman[0][table >> 4], MediaStructureReason.METADATA);
          if (!lossless(current) && (!progressive(current) || start > 0))
            in.require(tables.huffman[1][table & 15], MediaStructureReason.METADATA);
        }
        if (!progressive(current) || (start == 0 && (approx >> 4) == 0)) current.covered[id] = true;
      }
    }

    private int entropy(Frame current) {
      long bytes = 0;
      int restart = 0;
      while (true) {
        int value = in.u8();
        if (value != 255) {
          bytes++;
          continue;
        }
        value = entropyMarker();
        if (value == 0) {
          bytes++;
          continue;
        }
        if (value >= 208 && value <= 215) {
          in.record();
          in.require(restartInterval > 0 && value == 208 + restart, MediaStructureReason.SEQUENCE);
          restart = (restart + 1) & 7;
          continue;
        }
        if (value == 220) height(current);
        else {
          // T.81 D.1.8 permits discarding final arithmetic zero bytes, including an empty segment.
          in.require(
              (bytes > 0 || current.kind >= 201) && current.height > 0,
              MediaStructureReason.TRUNCATED);
          return value;
        }
      }
    }

    private int entropyMarker() {
      int value = in.u8();
      boolean fill = false;
      while (value == 255) {
        fill = true;
        value = in.u8();
      }
      in.require(value != 0 || !fill, MediaStructureReason.CODE);
      return value;
    }

    private void height(Frame current) {
      in.record();
      in.require(current.height == 0 && in.u16() == 4, MediaStructureReason.SEQUENCE);
      current.height = in.u16();
      in.require(current.height > 0, MediaStructureReason.DIMENSIONS);
      if (firstHeight == 0) firstHeight = current.height;
      if (hierarchy) in.require(current.height <= firstHeight, MediaStructureReason.DIMENSIONS);
      if (charge || frameCount > 1) reserve(in, current);
    }
  }

  private static boolean isFrame(int marker) {
    return marker >= 192 && marker <= 207 && marker != 196 && marker != 200 && marker != 204;
  }

  private static boolean lossless(Frame frame) {
    return (frame.kind & 3) == 3;
  }

  private static boolean progressive(Frame frame) {
    return (frame.kind & 3) == 2;
  }

  private static void complete(MediaInput in, Frame frame) {
    for (int i = 0; i < frame.components.length; i++)
      in.require(!frame.components[i] || frame.covered[i], MediaStructureReason.TRUNCATED);
  }

  private static void reserve(MediaInput in, Frame f) {
    in.budget.raster(f.width, f.height);
    in.budget.expand(
        in.product((in.product(in.product(f.width, f.count), f.precision) + 7) / 8, f.height));
  }

  private static void quant(MediaInput d, Tables tables) {
    while (d.remaining() > 0) {
      d.record();
      int header = d.u8();
      int precision = header >> 4;
      int id = header & 15;
      d.require(precision <= 1 && id <= 3, MediaStructureReason.CODE);
      for (int i = 0; i < 64; i++)
        d.require((precision == 0 ? d.u8() : d.u16()) != 0, MediaStructureReason.CODE);
      tables.quant[id] = true;
      tables.quantPrecision[id] = precision;
    }
  }

  private static void huffman(MediaInput d, Tables tables) {
    while (d.remaining() > 0) {
      d.record();
      int header = d.u8();
      int kind = header >> 4;
      int id = header & 15;
      d.require(kind <= 1 && id <= 3, MediaStructureReason.CODE);
      int symbols = 0;
      int slots = 1;
      for (int i = 0; i < 16; i++) {
        int count = d.u8();
        symbols += count;
        slots = (slots << 1) - count;
        d.require(slots >= 0 && symbols <= 256, MediaStructureReason.COUNT);
      }
      d.require(symbols > 0 && slots > 0, MediaStructureReason.CODE);
      d.skip(symbols);
      tables.huffman[kind][id] = true;
    }
  }

  /** One declared coding frame and component coverage, bounded by T.81 component identifiers. */
  private static final class Frame {
    final int kind;
    long width;
    long height;
    int precision;
    int count;
    boolean scanned;
    final boolean[] components = new boolean[256];
    final boolean[] covered = new boolean[256];
    final int[] quant = new int[256];
    final int[] componentOrder = new int[256];
    final int[] samplingProducts = new int[256];

    Frame(int kind) {
      this.kind = kind;
    }
  }
}
