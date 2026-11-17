package dev.erst.gridgrind.engine.media;

import dev.erst.gridgrind.core.problem.MediaStructureReason;
import java.util.ArrayList;
import java.util.List;

/** GIF87a/89a blocks, palettes and bounded LZW dictionary/output grammar. */
final class GifStructure {
  private final MediaInput input;
  private int version;
  private long width;
  private long height;
  private int global;
  private int transparent = -1;
  private int images;

  private GifStructure(MediaInput input) {
    this.input = input;
  }

  static ImageGeometry check(MediaInput input) {
    return new GifStructure(input).read();
  }

  private ImageGeometry read() {
    header();
    boolean ended = false;
    while (input.remaining() > 0 && !ended) {
      input.record();
      int marker = input.u8();
      switch (marker) {
        case 59 -> ended = true;
        case 44 -> image();
        case 33 -> extension();
        default -> input.require(false, MediaStructureReason.CODE);
      }
    }
    input.require(ended && input.remaining() == 0, MediaStructureReason.TERMINATOR);
    return new ImageGeometry(width, height);
  }

  private void header() {
    input.little = true;
    input.magic(71, 73, 70, 56);
    version = input.u8();
    input.require(
        (version == 55 || version == 57) && input.u8() == 97, MediaStructureReason.HEADER);
    width = input.u16();
    height = input.u16();
    input.require(width > 0 && height > 0, MediaStructureReason.DIMENSIONS);
    input.budget.raster(width, height);
    int packed = input.u8();
    int background = input.u8();
    input.u8();
    global = (packed & 128) != 0 ? 2 << (packed & 7) : 0;
    if (global > 0) {
      input.require(background < global, MediaStructureReason.CODE);
      input.skip(global * 3L);
    }
  }

  private void image() {
    long x = input.u16();
    long y = input.u16();
    long w = input.u16();
    long h = input.u16();
    input.require(
        w > 0 && h > 0 && x <= width - w && y <= height - h, MediaStructureReason.DIMENSIONS);
    int flags = input.u8();
    input.require((flags & 24) == 0, MediaStructureReason.CODE);
    int colors = (flags & 128) != 0 ? 2 << (flags & 7) : global;
    if ((flags & 128) != 0) input.skip(colors * 3L);
    if (colors == 0) input.profile("GIF/external-palette");
    input.require(transparent < colors, MediaStructureReason.CODE);
    if (images > 0) input.budget.raster(w, h);
    images++;
    long pixels = input.product(w, h);
    input.budget.expand(pixels);
    int minimum = input.u8();
    input.require(minimum >= 2 && minimum <= 8, MediaStructureReason.CODE);
    decode(input, blocks(input), minimum, pixels, colors);
    transparent = -1;
  }

  private void extension() {
    int label = input.u8();
    switch (label) {
      case 249 -> {
        input.require(version == 57 && input.u8() == 4, MediaStructureReason.HEADER);
        int control = input.u8();
        input.require((control & 224) == 0 && ((control >> 2) & 7) <= 3, MediaStructureReason.CODE);
        input.u16();
        int index = input.u8();
        transparent = (control & 1) != 0 ? index : -1;
        input.require(input.u8() == 0, MediaStructureReason.TERMINATOR);
      }
      case 1 -> plainText();
      case 255 -> {
        input.require(version == 57 && input.u8() == 11, MediaStructureReason.HEADER);
        input.skip(11);
        blocks(input);
      }
      default -> blocks(input);
    }
  }

  private void plainText() {
    input.require(version == 57 && input.u8() == 12, MediaStructureReason.HEADER);
    long x = input.u16();
    long y = input.u16();
    long w = input.u16();
    long h = input.u16();
    int cw = input.u8();
    int ch = input.u8();
    int fg = input.u8();
    int bg = input.u8();
    input.require(
        w > 0 && h > 0 && x <= width - w && y <= height - h && cw > 0 && ch > 0,
        MediaStructureReason.DIMENSIONS);
    if (global == 0) input.profile("GIF/external-text-palette");
    input.require(fg < global && bg < global && transparent < global, MediaStructureReason.CODE);
    blocks(input);
    transparent = -1;
  }

  private static List<MediaInput> blocks(MediaInput in) {
    var result = new ArrayList<MediaInput>();
    while (true) {
      in.record();
      int size = in.u8();
      if (size == 0) return result;
      result.add(in.take(size));
    }
  }

  private static void decode(
      MediaInput owner, List<MediaInput> blocks, int minimum, long expected, int palette) {
    var bits = new BlockBits(owner, blocks);
    int clear = 1 << minimum;
    int end = clear + 1;
    int next = end + 1;
    int width = minimum + 1;
    int previous = -1;
    int[] lengths = new int[4096];
    int[] first = new int[4096];
    int[] maximum = new int[4096];
    for (int i = 0; i < clear; i++) {
      lengths[i] = 1;
      first[i] = i;
      maximum[i] = i;
    }
    long output = 0;
    while (true) {
      int code = bits.read(width);
      if (code == clear) {
        next = end + 1;
        width = minimum + 1;
        previous = -1;
        continue;
      }
      if (code == end) {
        owner.require(output == expected, MediaStructureReason.COUNT);
        return;
      }
      int count;
      int head;
      int largest;
      if (code < next && lengths[code] > 0) {
        count = lengths[code];
        head = first[code];
        largest = maximum[code];
      } else {
        owner.require(previous >= 0 && code == next && next < 4096, MediaStructureReason.CODE);
        count = lengths[previous] + 1;
        head = first[previous];
        largest = maximum[previous];
      }
      owner.require(largest < palette && count <= expected - output, MediaStructureReason.CODE);
      output += count;
      if (previous >= 0 && next < 4096) {
        lengths[next] = lengths[previous] + 1;
        first[next] = first[previous];
        maximum[next] = Math.max(maximum[previous], head);
        next++;
        if (width < 12 && next == (1 << width)) width++;
      }
      previous = code;
    }
  }

  /** Traverses complete GIF subblocks as one little-endian coding bit stream. */
  private static final class BlockBits {
    private final MediaInput owner;
    private final List<MediaInput> blocks;
    private int block;
    private int available;
    private long buffer;

    BlockBits(MediaInput owner, List<MediaInput> blocks) {
      this.owner = owner;
      this.blocks = blocks;
    }

    int read(int width) {
      while (available < width) {
        while (block < blocks.size() && blocks.get(block).remaining() == 0) block++;
        owner.require(block < blocks.size(), MediaStructureReason.TRUNCATED);
        buffer |= (long) blocks.get(block).u8() << available;
        available += 8;
      }
      int value = (int) (buffer & ((1 << width) - 1));
      buffer >>>= width;
      available -= width;
      return value;
    }
  }
}
