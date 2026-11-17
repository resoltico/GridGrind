package dev.erst.gridgrind.engine.media;

import dev.erst.gridgrind.core.problem.MediaStructureReason;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import java.util.Arrays;

/** Checked image-local byte region; offsets and nested readers retain one work owner. */
final class MediaInput {
  final byte[] bytes;
  final int start;
  final int end;
  final int depth;
  final MediaAdmissionBudget budget;
  final ExcelPictureFormat format;
  int position;
  boolean little;

  @SuppressWarnings("PMD.ArrayIsStoredDirectly")
  // Every region borrows its admission owner's private immutable byte array; copying breaks that
  // owner.
  MediaInput(
      byte[] bytes,
      int start,
      int length,
      ExcelPictureFormat format,
      MediaAdmissionBudget budget,
      int depth) {
    if (start < 0 || length < 0 || start > bytes.length - length)
      throw new IllegalStateException("Invalid internal media region");
    this.bytes = bytes;
    this.start = start;
    this.end = start + length;
    this.position = start;
    this.format = format;
    this.budget = budget;
    this.depth = depth;
    budget.depth(depth);
  }

  int size() {
    return end - start;
  }

  int offset() {
    return position - start;
  }

  int remaining() {
    return end - position;
  }

  void require(boolean valid, MediaStructureReason reason) {
    if (!valid) throw MediaIssues.structure(format, reason, position);
  }

  void need(long count) {
    require(count >= 0 && count <= remaining(), MediaStructureReason.TRUNCATED);
  }

  int u8() {
    need(1);
    budget.inspect(1);
    int value = bytes[position] & 255;
    position++;
    return value;
  }

  int s8() {
    return (byte) u8();
  }

  int u16() {
    int a = u8();
    int b = u8();
    return little ? a | (b << 8) : (a << 8) | b;
  }

  int s16() {
    return (short) u16();
  }

  long u32() {
    long a = u16();
    long b = u16();
    return little ? a | (b << 16) : (a << 16) | b;
  }

  int s32() {
    return (int) u32();
  }

  long u64() {
    long a = u32();
    long b = u32();
    long value = little ? a | (b << 32) : (a << 32) | b;
    require(value >= 0, MediaStructureReason.OFFSET);
    return value;
  }

  void skip(long count) {
    need(count);
    budget.inspect(count);
    position += (int) count;
  }

  void seek(long offset) {
    require(offset >= 0 && offset <= size(), MediaStructureReason.OFFSET);
    position = start + (int) offset;
  }

  MediaInput region(long offset, long length) {
    require(
        offset >= 0 && length >= 0 && offset <= size() && length <= size() - offset,
        MediaStructureReason.OFFSET);
    var region = new MediaInput(bytes, start + (int) offset, (int) length, format, budget, depth);
    region.little = little;
    return region;
  }

  MediaInput take(long count) {
    need(count);
    var region = region(offset(), count);
    skip(count);
    return region;
  }

  MediaInput nested(long offset, long length, ExcelPictureFormat nestedFormat) {
    var region = region(offset, length);
    return new MediaInput(bytes, region.start, region.size(), nestedFormat, budget, depth + 1);
  }

  void magic(int... expected) {
    for (int value : expected) require(u8() == value, MediaStructureReason.HEADER);
  }

  void record() {
    budget.record();
  }

  long product(long a, long b) {
    require(a >= 0 && b >= 0, MediaStructureReason.LENGTH);
    if (a != 0 && b > Long.MAX_VALUE / a)
      throw MediaIssues.structure(format, MediaStructureReason.LENGTH, offset());
    return a * b;
  }

  long raster(long width, long height, long rowBytes) {
    require(width > 0 && height > 0, MediaStructureReason.DIMENSIONS);
    budget.raster(width, height);
    long size = product(rowBytes, height);
    budget.expand(size);
    return size;
  }

  byte[] copy() {
    budget.inspect(size());
    return Arrays.copyOfRange(bytes, start, end);
  }

  void profile(String value) {
    throw MediaIssues.profile(format, value);
  }
}
