package dev.erst.gridgrind.engine.media;

import dev.erst.gridgrind.core.problem.MediaStructureReason;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/** TIFF/BigTIFF directory graph, coded storage geometry and explicit stream/profile admission. */
final class TiffStructure {
  private final MediaInput root;
  private final boolean big;
  private final Set<Long> active = new HashSet<>();
  private final Set<Long> imagesDone = new HashSet<>();
  private final Set<Long> metadataDone = new HashSet<>();
  private ImageGeometry first = new ImageGeometry(0, 0);

  private TiffStructure(MediaInput root, boolean big) {
    this.root = root;
    this.big = big;
  }

  static ImageGeometry check(MediaInput in) {
    int a = in.u8();
    int b = in.u8();
    in.require(a == b && (a == 73 || a == 77), MediaStructureReason.HEADER);
    in.little = a == 73;
    int version = in.u16();
    in.require(version == 42 || version == 43, MediaStructureReason.HEADER);
    boolean big = version == 43;
    if (big) in.require(in.u16() == 8 && in.u16() == 0, MediaStructureReason.HEADER);
    long first = big ? in.u64() : in.u32();
    in.require(first > 0, MediaStructureReason.OFFSET);
    var reader = new TiffStructure(in, big);
    reader.walk(first, true, 0);
    in.require(reader.first.width() > 0, MediaStructureReason.METADATA);
    return reader.first;
  }

  @SuppressWarnings("PMD.AvoidInstantiatingObjectsInLoops")
  // Each distinct IFD needs its own immutable checked image layout. Aliases skip reconstruction;
  // shared record/pixel limits bound instances, and independent directories cannot share state.
  private void walk(long firstOffset, boolean image, int depth) {
    root.budget.depth(root.depth + depth);
    List<Long> chain = new ArrayList<>();
    Set<Long> completed = image ? imagesDone : metadataDone;
    long offset = firstOffset;
    try {
      while (offset != 0) {
        root.require(
            offset >= (big ? 16 : 8) && offset % 2 == 0 && !active.contains(offset),
            MediaStructureReason.OFFSET);
        if (completed.contains(offset)) return;
        active.add(offset);
        chain.add(offset);
        Directory directory = directory(offset);
        if (image) new Image(directory.fields).read();
        children(directory.fields, depth);
        completed.add(offset);
        offset = directory.next;
      }
    } finally {
      active.removeAll(chain);
    }
  }

  /** One directory's local parsed fields and next-directory edge. */
  private record Directory(Map<Integer, Field> fields, long next) {}

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  // The directory map is local to this private, single-invocation parser and never concurrently
  // read.
  private Directory directory(long offset) {
    root.record();
    var directory = root.region(offset, root.size() - offset);
    long count = big ? directory.u64() : directory.u16();
    root.require(count <= directory.remaining() / (big ? 20 : 12), MediaStructureReason.COUNT);
    Map<Integer, Field> fields = new HashMap<>();
    int previous = -1;
    for (long i = 0; i < count; i++) {
      root.record();
      int tag = directory.u16();
      int type = directory.u16();
      root.require(tag > previous, MediaStructureReason.SEQUENCE);
      previous = tag;
      fields.put(tag, field(directory, type));
    }
    return new Directory(fields, unsigned(directory));
  }

  private long unsigned(MediaInput input) {
    return big ? input.u64() : input.u32();
  }

  private Field field(MediaInput directory, int type) {
    long values = unsigned(directory);
    int unit = unit(type);
    if (unit == 0) throw MediaIssues.profile(root.format, "TIFF/field-type-" + type);
    root.require(values <= root.size() / unit, MediaStructureReason.LENGTH);
    long bytes = values * unit;
    var slot = directory.take(big ? 8 : 4);
    MediaInput data =
        bytes <= slot.size() ? slot.region(0, bytes) : root.region(unsigned(slot), bytes);
    root.budget.inspect(bytes);
    return new Field(type, values, data);
  }

  private void children(Map<Integer, Field> fields, int depth) {
    for (int pointer : POINTER_TAGS) {
      Field targets = fields.get(pointer);
      if (targets == null) continue;
      for (long i = 0; i < targets.count; i++) {
        long target = targets.value(i);
        if (target != 0) walk(target, pointer == 330, depth + 1);
      }
    }
  }

  private static final int[] POINTER_TAGS = {330, 34665, 34853, 40965};

  /** A directory's checked sample/storage facts and concrete coding dispatch. */
  private final class Image {
    private final Map<Integer, Field> fields;
    private final long width;
    private final long height;
    private final long samples;
    private final long planar;
    private final long compression;
    private final long photo;
    private final long predictor;
    private final long fill;
    private final @org.jspecify.annotations.Nullable Field bitField;
    private final long totalBits;
    private final Sampling sampling;
    private final Blocks blocks;

    Image(Map<Integer, Field> fields) {
      this.fields = fields;
      width = one(fields, 256, -1);
      height = one(fields, 257, -1);
      samples = one(fields, 277, 1);
      planar = one(fields, 284, 1);
      compression = one(fields, 259, 1);
      photo = one(fields, 262, -1);
      predictor = one(fields, 317, 1);
      fill = one(fields, 266, 1);
      basic();
      bitField = fields.get(258);
      root.require(bitField == null || bitField.count == samples, MediaStructureReason.COUNT);
      totalBits = sampleBits();
      palette();
      sampling = subsampling();
      blocks = new Blocks(this);
    }

    private void basic() {
      root.require(
          width > 0
              && height > 0
              && samples > 0
              && samples <= 65535
              && (planar == 1 || planar == 2)
              && (fill == 1 || fill == 2)
              && photo >= 0,
          MediaStructureReason.METADATA);
      root.require(
          compression >= 1 && compression <= 65535 && predictor >= 1 && predictor <= 3,
          MediaStructureReason.CODE);
    }

    private long sampleBits() {
      long total = 0;
      for (long i = 0; i < samples; i++) {
        long bits = bits(i);
        root.require(bits > 0 && bits <= 65535, MediaStructureReason.CODE);
        total += bits;
      }
      return total;
    }

    private long bits(long component) {
      return bitField == null ? 1 : bitField.value(component);
    }

    private void palette() {
      if (photo != 3) return;
      root.require(samples == 1 && bitField != null, MediaStructureReason.METADATA);
      long bits = bits(0);
      Field colors = fields.get(320);
      root.require(
          bits < 31 && colors != null && colors.count == 3 * (1L << bits),
          MediaStructureReason.COUNT);
    }

    private Sampling subsampling() {
      long x = 2;
      long y = 2;
      if (photo == 6) {
        root.require(samples == 3 && totalBits == 24, MediaStructureReason.METADATA);
        Field sub = fields.get(530);
        if (sub != null) {
          root.require(sub.count == 2, MediaStructureReason.COUNT);
          x = sub.value(0);
          y = sub.value(1);
        }
        root.require(
            (x == 1 || x == 2 || x == 4) && (y == 1 || y == 2 || y == 4) && y <= x,
            MediaStructureReason.CODE);
      }
      return new Sampling(x, y);
    }

    void read() {
      JpegStructure.Tables tables = tables();
      if (compression == 6)
        oldJpeg(fields, width, height, samples, totalBits, blocks.offsets, blocks.counts);
      else for (long i = 0; i < blocks.count; i++) block(i, tables);
      if (first.width() == 0) first = new ImageGeometry(width, height);
    }

    private JpegStructure.Tables tables() {
      Field field = fields.get(347);
      if (field == null) return new JpegStructure.Tables();
      var table = field.data.nested(0, field.data.size(), ExcelPictureFormat.JPEG);
      table.record();
      return JpegStructure.tables(table);
    }

    private void block(long number, JpegStructure.Tables tables) {
      root.record();
      long plane = number / blocks.perPlane;
      long index = number % blocks.perPlane;
      long blockRow = index / blocks.across;
      long w = blocks.width;
      long h =
          blocks.tiled ? blocks.height : Math.min(blocks.height, height - blockRow * blocks.height);
      if (plane == 0) root.budget.raster(w, h);
      long bits = planar == 2 ? bits(plane) : totalBits;
      long rowBytes;
      long rows = h;
      if (photo == 6 && planar == 1) {
        rowBytes = root.product(ceil(w, sampling.x), sampling.x * sampling.y + 2);
        rows = ceil(h, sampling.y);
      } else {
        if (photo == 6 && plane > 0) {
          w = ceil(w, sampling.x);
          h = ceil(h, sampling.y);
          rows = h;
        }
        rowBytes = (root.product(w, bits) + 7) / 8;
      }
      long expected = root.product(rowBytes, rows);
      root.budget.expand(expected);
      long offset = blocks.offsets.value(number);
      long length = blocks.counts.value(number);
      root.require(length > 0, MediaStructureReason.LENGTH);
      var data = root.region(offset, length);
      root.budget.inspect(length);
      coding(data, w, h, rows, rowBytes, expected, plane, tables);
    }

    private void coding(
        MediaInput data,
        long w,
        long h,
        long rows,
        long rowBytes,
        long expected,
        long plane,
        JpegStructure.Tables tables) {
      switch ((int) compression) {
        case 1 -> data.require(data.size() >= expected, MediaStructureReason.TRUNCATED);
        case 2, 3, 4 -> fax(data, w, rows);
        case 5 -> MediaStreams.tiffLzw(data, expected);
        case 8, 32946 -> MediaStreams.inflate(List.of(data), expected);
        case 32773 -> {
          // Subsampled contiguous YCbCr clumps interleave pixel rows; there is no separable byte
          // row.
          if (photo == 6 && planar == 1 && sampling.y > 1) MediaStreams.packBits(data, expected, 1);
          else MediaStreams.packBits(data, rowBytes, rows);
        }
        case 7 -> jpeg(data, w, h, plane, tables);
        default -> root.profile("TIFF/compression-" + compression);
      }
    }

    private void fax(MediaInput data, long w, long rows) {
      root.require(samples == 1 && totalBits == 1 && planar == 1, MediaStructureReason.METADATA);
      long options = compression == 2 ? 0 : one(fields, compression == 3 ? 292 : 293, 0);
      FaxStructure.check(data, w, rows, (int) compression, options, (int) fill);
    }

    private void jpeg(MediaInput data, long w, long h, long plane, JpegStructure.Tables tables) {
      var member = data.nested(0, data.size(), ExcelPictureFormat.JPEG);
      member.record();
      var geometry = JpegStructure.member(member, tables);
      root.require(
          geometry.width() == w
              && geometry.height() == h
              && geometry.components() == (planar == 2 ? 1 : samples),
          MediaStructureReason.METADATA);
      if (planar == 2)
        root.require(geometry.sampleBits() == bits(plane), MediaStructureReason.METADATA);
      else
        for (long c = 0; c < samples; c++)
          root.require(geometry.sampleBits() == bits(c), MediaStructureReason.METADATA);
    }
  }

  /** Horizontal/vertical chroma sampling is a pair in the TIFF contract. */
  private record Sampling(long x, long y) {}

  /** Checked strip/tile cardinality and encoded member extents for one image directory. */
  private final class Blocks {
    final boolean tiled;
    final long width;
    final long height;
    final long across;
    final long perPlane;
    final long count;
    final Field offsets;
    final Field counts;

    Blocks(Image image) {
      tiled = image.fields.containsKey(324);
      if (tiled) {
        width = one(image.fields, 322, -1);
        height = one(image.fields, 323, -1);
        offsets = required(image.fields, 324);
        counts = required(image.fields, 325);
      } else {
        width = image.width;
        height = one(image.fields, 278, 4294967295L);
        offsets = required(image.fields, 273);
        counts = required(image.fields, 279);
      }
      root.require(width > 0 && height > 0, MediaStructureReason.DIMENSIONS);
      if (tiled) root.require(width % 16 == 0 && height % 16 == 0, MediaStructureReason.DIMENSIONS);
      across = tiled ? ceil(image.width, width) : 1;
      perPlane = root.product(across, ceil(image.height, height));
      count = root.product(perPlane, image.planar == 2 ? image.samples : 1);
      root.require(offsets.count == count && counts.count == count, MediaStructureReason.COUNT);
    }
  }

  private void oldJpeg(
      Map<Integer, Field> fields,
      long width,
      long height,
      long samples,
      long bits,
      Field offsets,
      Field counts) {
    Field position = required(fields, 513);
    Field length = required(fields, 514);
    root.require(position.count == 1 && length.count == 1, MediaStructureReason.COUNT);
    long start = position.value(0);
    long size = length.value(0);
    var member = root.nested(start, size, ExcelPictureFormat.JPEG);
    member.record();
    var geometry = JpegStructure.member(member, new JpegStructure.Tables());
    root.require(
        geometry.width() == width
            && geometry.height() == height
            && geometry.components() == samples
            && root.product(geometry.sampleBits(), samples) == bits,
        MediaStructureReason.METADATA);
    if (one(fields, 512, 1) != 1 || offsets.count != 1)
      root.profile("TIFF/old-JPEG-reconstruction");
    var layout = JpegInterchangeLayout.read(root.nested(start, size, ExcelPictureFormat.JPEG));
    root.require(
        offsets.value(0) == start + layout.entropyOffset
            && counts.value(0) == layout.entropyEndOffset - layout.entropyOffset,
        MediaStructureReason.OFFSET);
    oldPointers(fields, layout, samples, start);
    root.budget.raster(width, height);
    root.budget.expand(root.product((root.product(width, bits) + 7) / 8, height));
  }

  private void oldPointers(
      Map<Integer, Field> fields,
      JpegInterchangeLayout layout,
      long samples,
      long interchangeOffset) {
    for (int tag : new int[] {519, 520, 521}) {
      Field pointers = required(fields, tag);
      root.require(pointers.count == samples, MediaStructureReason.METADATA);
      long[] actual =
          tag == 519 ? layout.quantOffsets : tag == 520 ? layout.dcOffsets : layout.acOffsets;
      for (int i = 0; i < samples; i++)
        root.require(
            actual[i] >= 0 && pointers.value(i) == interchangeOffset + actual[i],
            MediaStructureReason.OFFSET);
    }
  }

  private Field required(Map<Integer, Field> fields, int tag) {
    Field value = fields.get(tag);
    if (value == null)
      throw MediaIssues.structure(root.format, MediaStructureReason.METADATA, root.position);
    return value;
  }

  private long one(Map<Integer, Field> fields, int tag, long fallback) {
    Field field = fields.get(tag);
    if (field == null) {
      root.require(fallback >= 0, MediaStructureReason.METADATA);
      return fallback;
    }
    root.require(field.count == 1, MediaStructureReason.COUNT);
    return field.value(0);
  }

  private static long ceil(long numerator, long denominator) {
    return numerator / denominator + (numerator % denominator == 0 ? 0 : 1);
  }

  private static int unit(int type) {
    return switch (type) {
      case 1, 2, 6, 7 -> 1;
      case 3, 8 -> 2;
      case 4, 9, 11, 13 -> 4;
      case 5, 10, 12, 16, 17, 18 -> 8;
      default -> 0;
    };
  }

  private record Field(int type, long count, MediaInput data) {
    long value(long index) {
      data.require(index >= 0 && index < count, MediaStructureReason.COUNT);
      int unit = unit(type);
      var value = data.region(index * unit, unit);
      return switch (type) {
        case 1 -> value.u8();
        case 3 -> value.u16();
        case 4, 13 -> value.u32();
        case 16, 18 -> value.u64();
        default ->
            throw MediaIssues.structure(data.format, MediaStructureReason.METADATA, data.start);
      };
    }
  }
}
