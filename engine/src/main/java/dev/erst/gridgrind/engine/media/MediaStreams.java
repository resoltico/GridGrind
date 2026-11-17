package dev.erst.gridgrind.engine.media;

import dev.erst.gridgrind.core.problem.MediaStructureReason;
import java.util.Arrays;
import java.util.List;
import java.util.zip.DataFormatException;
import java.util.zip.Inflater;

/** Bounded coding grammars; output is counted rather than retained as a complete raster. */
final class MediaStreams {
  private MediaStreams() {}

  /** Consumes a bounded expansion chunk without retaining the raster. */
  @FunctionalInterface
  interface InflatedBytes {
    /** Observes exactly count bytes at the absolute expanded-stream offset. */
    void accept(byte[] bytes, int count, long offset);
  }

  static void inflate(List<MediaInput> parts, long expected) {
    inflate(parts, expected, false, (bytes, count, offset) -> {});
  }

  static void inflatePng(List<MediaInput> parts, long expected, InflatedBytes output) {
    inflate(parts, expected, true, output);
  }

  private static void inflate(
      List<MediaInput> parts, long expected, boolean trailingPng, InflatedBytes observer) {
    MediaInput owner = parts.getFirst();
    long output = 0;
    long read = 0;
    int part = 0;
    try (Inflater inflater = new Inflater()) {
      byte[] work = new byte[8192];
      while (!inflater.finished()) {
        if (inflater.needsInput()) {
          owner.require(part < parts.size(), MediaStructureReason.TRUNCATED);
          var input = parts.get(part);
          part++;
          inflater.setInput(input.bytes, input.start, input.size());
        }
        int length = work.length;
        if (expected >= 0) length = (int) Math.min(length, expected - output + 1);
        else length = (int) Math.min(length, owner.budget.remainingExpansion() + 1);
        int count;
        try {
          count = inflater.inflate(work, 0, length);
        } catch (DataFormatException failure) {
          long written = inflater.getBytesWritten();
          if (expected < 0) owner.budget.expand(written - output);
          else if (written > expected) owner.budget.expand(written - expected);
          owner.budget.inspect(inflater.getBytesRead() - read);
          throw MediaIssues.coding(owner.format, owner.start, failure);
        }
        long nowRead = inflater.getBytesRead();
        owner.budget.inspect(nowRead - read);
        read = nowRead;
        if (expected < 0) owner.budget.expand(count);
        observer.accept(work, count, output);
        output += count;
        if (expected >= 0 && output > expected) {
          owner.budget.expand(output - expected);
          owner.require(false, MediaStructureReason.COUNT);
        }
        if (count == 0 && !inflater.finished()) {
          if (inflater.needsDictionary()) owner.profile("zlib/preset-dictionary");
          owner.require(inflater.needsInput(), MediaStructureReason.CODE);
        }
      }
      owner.require(expected < 0 || output == expected, MediaStructureReason.COUNT);
      owner.require(
          trailingPng || (inflater.getRemaining() == 0 && part == parts.size()),
          MediaStructureReason.EXCESS_DATA);
    }
  }

  static void packBits(MediaInput input, long rowBytes, long rows) {
    long row = 0;
    long within = 0;
    while (input.remaining() > 0) {
      int n = input.s8();
      if (n == -128) continue;
      int count = n >= 0 ? n + 1 : 1 - n;
      input.require(row < rows && count <= rowBytes - within, MediaStructureReason.COUNT);
      if (n >= 0) input.skip(count);
      else input.u8();
      within += count;
      if (within == rowBytes) {
        within = 0;
        row++;
      }
    }
    input.require(row == rows && within == 0, MediaStructureReason.TRUNCATED);
  }

  static void tiffLzw(MediaInput input, long expected) {
    var bits = new Bits(input, false);
    int[] lengths = new int[4096];
    Arrays.fill(lengths, 0, 256, 1);
    int width = 9;
    int next = 258;
    int previous = -1;
    long output = 0;
    boolean first = true;
    while (true) {
      int code = bits.read(width);
      input.require(!first || code == 256, MediaStructureReason.SEQUENCE);
      first = false;
      if (code == 256) {
        width = 9;
        next = 258;
        previous = -1;
        continue;
      }
      if (code == 257) {
        input.require(output == expected, MediaStructureReason.COUNT);
        bits.finish();
        return;
      }
      int count;
      if (previous < 0) {
        input.require(code < 256, MediaStructureReason.CODE);
        count = 1;
      } else if (code < next) count = lengths[code];
      else {
        input.require(code == next && next < 4096, MediaStructureReason.CODE);
        count = lengths[previous] + 1;
      }
      input.require(count > 0 && count <= expected - output, MediaStructureReason.COUNT);
      output += count;
      if (previous >= 0) {
        input.require(next <= 4094, MediaStructureReason.CODE);
        lengths[next] = lengths[previous] + 1;
        next++;
        if (next == 511 || next == 1023 || next == 2047) width++;
      }
      previous = code;
    }
  }

  /** A fixed-width bit cursor with byte-order selected by its coding standard. */
  static final class Bits {
    private final MediaInput input;
    private final boolean lowFirst;
    private long buffer;
    private int available;

    Bits(MediaInput input, boolean lowFirst) {
      this.input = input;
      this.lowFirst = lowFirst;
    }

    int read(int width) {
      while (available < width) {
        long value = input.u8();
        if (lowFirst) buffer |= value << available;
        else buffer = (buffer << 8) | value;
        available += 8;
      }
      int value;
      if (lowFirst) {
        value = (int) (buffer & ((1L << width) - 1));
        buffer >>>= width;
        available -= width;
      } else {
        available -= width;
        value = (int) ((buffer >>> available) & ((1L << width) - 1));
        buffer &= (1L << available) - 1;
      }
      return value;
    }

    void align() {
      available = 0;
      buffer = 0;
    }

    void finish() {
      input.require(input.remaining() == 0, MediaStructureReason.EXCESS_DATA);
    }
  }
}
