package dev.erst.gridgrind.engine.media;

import dev.erst.gridgrind.core.problem.MediaStructureReason;
import java.util.Arrays;

/** JPEG-member-relative pointer facts for the qualified single-scan baseline interchange layout. */
final class JpegInterchangeLayout {
  final long[] quantOffsets = new long[256];
  final long[] dcOffsets = new long[256];
  final long[] acOffsets = new long[256];
  long entropyOffset;
  long entropyEndOffset;
  private final int memberStart;
  private int components;
  private boolean frame;
  private final long[] quantTables = new long[4];
  private final long[] dcTables = new long[4];
  private final long[] acTables = new long[4];
  private final int[] ids = new int[256];
  private final int[] quantIds = new int[256];

  private JpegInterchangeLayout(int memberStart) {
    this.memberStart = memberStart;
    Arrays.fill(quantOffsets, -1);
    Arrays.fill(dcOffsets, -1);
    Arrays.fill(acOffsets, -1);
    Arrays.fill(quantTables, -1);
    Arrays.fill(dcTables, -1);
    Arrays.fill(acTables, -1);
  }

  static JpegInterchangeLayout read(MediaInput input) {
    return new JpegInterchangeLayout(input.start).readStream(input);
  }

  private JpegInterchangeLayout readStream(MediaInput in) {
    in.magic(255, 216);
    while (in.remaining() > 0) {
      in.record();
      in.require(in.u8() == 255, MediaStructureReason.HEADER);
      int marker;
      do {
        marker = in.u8();
      } while (marker == 255);
      if (marker == 217) break;
      int length = in.u16();
      in.require(length >= 2, MediaStructureReason.LENGTH);
      var data = in.take(length - 2);
      switch (marker) {
        case 219 -> quantization(data);
        case 196 -> huffman(data);
        case 218 -> {
          return scan(in, data);
        }
        default -> {
          if (marker >= 192 && marker <= 207 && marker != 200 && marker != 204) frame(data, marker);
        }
      }
    }
    throw MediaIssues.structure(in.format, MediaStructureReason.TRUNCATED, in.position);
  }

  private void quantization(MediaInput data) {
    while (data.remaining() > 0) {
      data.record();
      int selector = data.u8();
      int precision = selector >> 4;
      int id = selector & 15;
      data.require(id < 4 && precision <= 1, MediaStructureReason.CODE);
      if (precision != 0) data.profile("TIFF/old-JPEG-nonbaseline-quantization");
      quantTables[id] = data.position - memberStart;
      data.skip(64);
    }
  }

  private void huffman(MediaInput data) {
    while (data.remaining() > 0) {
      data.record();
      int selector = data.u8();
      int kind = selector >> 4;
      int id = selector & 15;
      data.require(kind < 2 && id < 4, MediaStructureReason.CODE);
      long tableOffset = data.position - memberStart;
      int symbols = 0;
      for (int j = 0; j < 16; j++) symbols += data.u8();
      data.skip(symbols);
      if (kind == 0) dcTables[id] = tableOffset;
      else acTables[id] = tableOffset;
    }
  }

  private void frame(MediaInput data, int marker) {
    if (marker != 192 || frame) data.profile("TIFF/old-JPEG-nonbaseline-frame");
    frame = true;
    data.skip(5);
    components = data.u8();
    for (int i = 0; i < components; i++) {
      ids[i] = data.u8();
      data.u8();
      quantIds[ids[i]] = data.u8();
    }
  }

  private JpegInterchangeLayout scan(MediaInput in, MediaInput data) {
    int count = data.u8();
    if (!frame || count != components) in.profile("TIFF/old-JPEG-multiple-scans");
    for (int i = 0; i < count; i++) {
      int id = data.u8();
      int tables = data.u8();
      int index = index(id, count);
      in.require(index >= 0, MediaStructureReason.METADATA);
      quantOffsets[index] = quantTables[quantIds[id]];
      dcOffsets[index] = dcTables[tables >> 4];
      acOffsets[index] = acTables[tables & 15];
    }
    entropyOffset = in.position - memberStart;
    while (in.remaining() > 0) {
      long start = in.position;
      int value = in.u8();
      if (value == 255) {
        do {
          value = in.u8();
        } while (value == 255);
        if (value == 0 || value >= 208 && value <= 215) continue;
        if (value != 217) in.profile("TIFF/old-JPEG-multiple-scans");
        entropyEndOffset = start - memberStart;
        in.require(in.remaining() == 0, MediaStructureReason.EXCESS_DATA);
        return this;
      }
    }
    throw MediaIssues.structure(in.format, MediaStructureReason.TRUNCATED, in.position);
  }

  private int index(int id, int count) {
    for (int i = 0; i < count; i++) if (ids[i] == id) return i;
    return -1;
  }
}
