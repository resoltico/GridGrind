package dev.erst.gridgrind.engine.media;

import dev.erst.gridgrind.core.problem.MediaStructureReason;
import java.util.BitSet;

/** TIFF6 sections10/11 and ITU-T T.4 tables2/3/4/5: bounded two-row coding grammar. */
final class FaxStructure {
  private FaxStructure() {}

  private static final Codes WHITE =
      new Codes(
          "0:00110101 1:000111 2:0111 3:1000 4:1011 5:1100 6:1110 7:1111 8:10011 9:10100 10:00111"
              + " 11:01000 12:001000 13:000011 14:110100 15:110101 16:101010 17:101011 18:0100111"
              + " 19:0001100 20:0001000 21:0010111 22:0000011 23:0000100 24:0101000 25:0101011"
              + " 26:0010011 27:0100100 28:0011000 29:00000010 30:00000011 31:00011010 32:00011011"
              + " 33:00010010 34:00010011 35:00010100 36:00010101 37:00010110 38:00010111"
              + " 39:00101000 40:00101001 41:00101010 42:00101011 43:00101100 44:00101101"
              + " 45:00000100 46:00000101 47:00001010 48:00001011 49:01010010 50:01010011"
              + " 51:01010100 52:01010101 53:00100100 54:00100101 55:01011000 56:01011001"
              + " 57:01011010 58:01011011 59:01001010 60:01001011 61:00110010 62:00110011"
              + " 63:00110100 64:11011 128:10010 192:010111 256:0110111 320:00110110 384:00110111"
              + " 448:01100100 512:01100101 576:01101000 640:01100111 704:011001100 768:011001101"
              + " 832:011010010 896:011010011 960:011010100 1024:011010101 1088:011010110"
              + " 1152:011010111 1216:011011000 1280:011011001 1344:011011010 1408:011011011"
              + " 1472:010011000 1536:010011001 1600:010011010 1664:011000 1728:010011011"
              + " 1792:00000001000 1856:00000001100 1920:00000001101 1984:000000010010"
              + " 2048:000000010011 2112:000000010100 2176:000000010101 2240:000000010110"
              + " 2304:000000010111 2368:000000011100 2432:000000011101 2496:000000011110"
              + " 2560:000000011111  -1:000000001111");
  private static final Codes BLACK =
      new Codes(
          "0:0000110111 1:010 2:11 3:10 4:011 5:0011 6:0010 7:00011 8:000101 9:000100 10:0000100"
              + " 11:0000101 12:0000111 13:00000100 14:00000111 15:000011000 16:0000010111"
              + " 17:0000011000 18:0000001000 19:00001100111 20:00001101000 21:00001101100"
              + " 22:00000110111 23:00000101000 24:00000010111 25:00000011000 26:000011001010"
              + " 27:000011001011 28:000011001100 29:000011001101 30:000001101000 31:000001101001"
              + " 32:000001101010 33:000001101011 34:000011010010 35:000011010011 36:000011010100"
              + " 37:000011010101 38:000011010110 39:000011010111 40:000001101100 41:000001101101"
              + " 42:000011011010 43:000011011011 44:000001010100 45:000001010101 46:000001010110"
              + " 47:000001010111 48:000001100100 49:000001100101 50:000001010010 51:000001010011"
              + " 52:000000100100 53:000000110111 54:000000111000 55:000000100111 56:000000101000"
              + " 57:000001011000 58:000001011001 59:000000101011 60:000000101100 61:000001011010"
              + " 62:000001100110 63:000001100111 64:0000001111 128:000011001000 192:000011001001"
              + " 256:000001011011 320:000000110011 384:000000110100 448:000000110101"
              + " 512:0000001101100 576:0000001101101 640:0000001001010 704:0000001001011"
              + " 768:0000001001100 832:0000001001101 896:0000001110010 960:0000001110011"
              + " 1024:0000001110100 1088:0000001110101 1152:0000001110110 1216:0000001110111"
              + " 1280:0000001010010 1344:0000001010011 1408:0000001010100 1472:0000001010101"
              + " 1536:0000001011010 1600:0000001011011 1664:0000001100100 1728:0000001100101"
              + " 1792:00000001000 1856:00000001100 1920:00000001101 1984:000000010010"
              + " 2048:000000010011 2112:000000010100 2176:000000010101 2240:000000010110"
              + " 2304:000000010111 2368:000000011100 2432:000000011101 2496:000000011110"
              + " 2560:000000011111  -1:000000001111");

  private static final Codes MODES =
      new Codes(
          "0:1 1:011 -1:010 2:000011 -2:000010 3:0000011 -3:0000010 10:001 11:0001 12:0000001");

  static void check(
      MediaInput input, long width, long height, int compression, long options, int fill) {
    input.require(
        width > 0 && width <= Integer.MAX_VALUE && height > 0, MediaStructureReason.DIMENSIONS);
    options(input, compression, options);
    var bits = new Cursor(input, fill == 2);
    var row = new Row(bits, (int) width, (options & 2) != 0);
    for (long line = 0; line < height; line++) {
      row.decode(oneDimensional(bits, compression, options, line));
      if (compression == 2) bits.align();
    }
    if (compression == 2) {
      input.require(bits.remaining() == 0, MediaStructureReason.EXCESS_DATA);
      return;
    }
    terminator(bits, compression, options);
    bits.zeroPadding();
  }

  private static void options(MediaInput input, int compression, long options) {
    boolean valid =
        switch (compression) {
          case 2 -> options == 0;
          case 3 -> (options & ~7L) == 0;
          case 4 -> (options & ~2L) == 0;
          default -> false;
        };
    input.require(valid, MediaStructureReason.CODE);
  }

  private static boolean oneDimensional(Cursor bits, int compression, long options, long line) {
    if (compression != 3) return compression != 4;
    bits.eol();
    if ((options & 4) != 0) bits.input.require(bits.aligned(), MediaStructureReason.SEQUENCE);
    if ((options & 1) == 0) return true;
    boolean one = bits.bit() == 1;
    bits.input.require(line != 0 || one, MediaStructureReason.SEQUENCE);
    return one;
  }

  private static void terminator(Cursor bits, int compression, long options) {
    // TIFF permits stopping after all declared rows; any remaining terminator must be framed.
    if (bits.remaining() <= 7) return;
    bits.eol();
    if (compression == 4) {
      bits.eol();
      return;
    }
    if ((options & 1) != 0) bits.input.require(bits.bit() == 1, MediaStructureReason.CODE);
    if (bits.remaining() <= 7) return;
    for (int i = 1; i < 6; i++) {
      bits.eol();
      if ((options & 1) != 0) bits.input.require(bits.bit() == 1, MediaStructureReason.CODE);
    }
  }

  /** Exactly two packed bit rows and current coding position; never a complete raster. */
  private static final class Row {
    private final Cursor bits;
    private final int width;
    private final boolean uncompressed;
    private BitSet reference;
    private BitSet output;
    private int position;
    private boolean black;
    private boolean initial;

    Row(Cursor bits, int width, boolean uncompressed) {
      this.bits = bits;
      this.width = width;
      this.uncompressed = uncompressed;
      reference = new BitSet(width);
      output = new BitSet(width);
    }

    void decode(boolean one) {
      output.clear();
      position = 0;
      black = false;
      initial = true;
      while (position < width) {
        if (one) one();
        else two();
        initial = false;
      }
      bits.input.require(position == width, MediaStructureReason.COUNT);
      BitSet swap = reference;
      reference = output;
      output = swap;
    }

    private void one() {
      int count = run(bits, black, width - position);
      if (count < 0) {
        bits.input.require(uncompressed, MediaStructureReason.CODE);
        literal();
      } else {
        emit(count, black);
        black = !black;
      }
    }

    private void two() {
      int mode = MODES.read(bits);
      if (mode == 12) {
        bits.input.require(uncompressed && bits.take(3) == 7, MediaStructureReason.CODE);
        literal();
      } else if (mode == 10) horizontal();
      else relative(mode);
    }

    private void horizontal() {
      int first = run(bits, black, width - position);
      bits.input.require(first >= 0, MediaStructureReason.CODE);
      emit(first, black);
      int second = run(bits, !black, width - position);
      bits.input.require(second >= 0 && first + second > 0, MediaStructureReason.CODE);
      emit(second, !black);
    }

    private void relative(int mode) {
      int b1 = transition(reference, position, initial, !black, width);
      int end = mode == 11 ? transition(reference, b1, false, black, width) : b1 + mode;
      bits.input.require(
          end >= position && end <= width && (mode != 11 || end > position),
          MediaStructureReason.COUNT);
      emit(end - position, black);
      if (mode != 11) black = !black;
    }

    private void emit(int count, boolean color) {
      bits.input.require(count >= 0 && count <= width - position, MediaStructureReason.COUNT);
      if (color) output.set(position, position + count);
      position += count;
    }

    private void literal() {
      while (true) {
        int zeros = 0;
        while (bits.bit() == 0) {
          zeros++;
          bits.input.require(zeros <= 10, MediaStructureReason.CODE);
        }
        if (zeros <= 4) {
          emit(zeros, false);
          emit(1, true);
        } else if (zeros == 5) emit(5, false);
        else {
          emit(zeros - 6, false);
          black = bits.bit() == 1;
          return;
        }
      }
    }
  }

  private static int transition(BitSet row, int x, boolean initial, boolean target, int width) {
    int start = initial ? x : x + 1;
    if (start >= width) return width;
    int p = start;
    if (row.get(p) == target) {
      if (p == 0 || row.get(p - 1) != target) return p;
      p = target ? row.nextClearBit(p) : row.nextSetBit(p);
      if (p < 0 || p >= width) return width;
    }
    p = target ? row.nextSetBit(p) : row.nextClearBit(p);
    return p < 0 || p >= width ? width : p;
  }

  private static int run(Cursor bits, boolean black, int remaining) {
    int total = 0;
    while (true) {
      int value = (black ? BLACK : WHITE).read(bits);
      if (value < 0) {
        bits.input.require(total == 0, MediaStructureReason.SEQUENCE);
        return value;
      }
      bits.input.require(value <= remaining - total, MediaStructureReason.COUNT);
      total += value;
      if (value < 64) return total;
    }
  }

  /** MSB/LSB-selected source cursor and terminal bit-padding checks. */
  private static final class Cursor {
    final MediaInput input;
    final boolean reversed;
    int value;
    int left;

    Cursor(MediaInput input, boolean reversed) {
      this.input = input;
      this.reversed = reversed;
    }

    int bit() {
      if (left == 0) {
        value = input.u8();
        left = 8;
      }
      int result = reversed ? value & 1 : (value >>> 7) & 1;
      value = reversed ? value >>> 1 : value << 1;
      left--;
      return result;
    }

    int take(int count) {
      int v = 0;
      for (int i = 0; i < count; i++) v = (v << 1) | bit();
      return v;
    }

    long remaining() {
      return input.remaining() * 8L + left;
    }

    boolean aligned() {
      return left == 0;
    }

    void align() {
      left = 0;
      value = 0;
    }

    void eol() {
      int n = 0;
      while (bit() == 0) n++;
      input.require(n >= 11, MediaStructureReason.TERMINATOR);
    }

    void zeroPadding() {
      while (remaining() > 0) input.require(bit() == 0, MediaStructureReason.EXCESS_DATA);
    }
  }

  /** Prefix trie built solely from the bounded normative T.4 codeword tables. */
  private static final class Codes {
    // At most104normative codewords, each at most13bits; the trie capacity follows that bound.
    private final int[][] edges = new int[1 + 105 * 13][2];
    private final int[] values = new int[1 + 105 * 13];

    Codes(String specification) {
      int used = 1;
      java.util.Arrays.fill(values, Integer.MIN_VALUE);
      for (String entry : specification.trim().split(" +")) {
        String[] pair = entry.split(":");
        int node = 0;
        for (char bit : pair[1].toCharArray()) {
          int side = bit - '0';
          if (edges[node][side] == 0) {
            edges[node][side] = used;
            used++;
          }
          node = edges[node][side];
        }
        values[node] = Integer.parseInt(pair[0]);
      }
    }

    int read(Cursor bits) {
      int node = 0;
      for (int i = 0; i < 13; i++) {
        node = edges[node][bits.bit()];
        bits.input.require(node != 0, MediaStructureReason.CODE);
        if (values[node] != Integer.MIN_VALUE) return values[node];
      }
      throw MediaIssues.structure(
          bits.input.format, MediaStructureReason.CODE, bits.input.position);
    }
  }
}
