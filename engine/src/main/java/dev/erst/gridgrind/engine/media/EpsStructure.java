package dev.erst.gridgrind.engine.media;

import dev.erst.gridgrind.core.problem.MediaStructureReason;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import java.util.ArrayDeque;
import java.util.Deque;

/** EPSF/DSC framing and PostScript lexical balance; never executes a program or opens resources. */
final class EpsStructure {
  private EpsStructure() {}

  static ImageGeometry check(MediaInput input) {
    MediaInput program = input;
    if (input.size() >= 4 && (input.bytes[input.start] & 255) == 197) {
      input.little = true;
      input.magic(197, 208, 211, 198);
      long ps = input.u32();
      long psBytes = input.u32();
      long wmf = input.u32();
      long wmfBytes = input.u32();
      long tiff = input.u32();
      long tiffBytes = input.u32();
      int checksum = input.u16();
      int calculated = 0;
      var header = input.region(0, 28);
      for (int i = 0; i < 14; i++) calculated ^= header.u16();
      input.require(checksum == 65535 || checksum == calculated, MediaStructureReason.CHECKSUM);
      input.require(
          ps >= 30
              && psBytes > 0
              && (wmf == 0) == (wmfBytes == 0)
              && (tiff == 0) == (tiffBytes == 0),
          MediaStructureReason.OFFSET);
      program = input.nested(ps, psBytes, ExcelPictureFormat.EPS);
      if (wmfBytes > 0) {
        input.require(
            wmf >= 30 && !overlap(ps, psBytes, wmf, wmfBytes), MediaStructureReason.OFFSET);
        ImageStructures.check(input.nested(wmf, wmfBytes, ExcelPictureFormat.WMF));
      }
      if (tiffBytes > 0) {
        input.require(
            tiff >= 30
                && !overlap(ps, psBytes, tiff, tiffBytes)
                && !overlap(wmf, wmfBytes, tiff, tiffBytes),
            MediaStructureReason.OFFSET);
        ImageStructures.check(input.nested(tiff, tiffBytes, ExcelPictureFormat.TIFF));
      }
    }
    new Scanner(program).scan();
    return ImageGeometry.vector();
  }

  private static boolean overlap(long a, long n, long b, long m) {
    return n > 0 && m > 0 && a < b + m && b < a + n;
  }

  /** Bounded DSC section and PostScript lexical framing state. */
  private static final class Scanner {
    private final MediaInput in;
    private final Deque<Integer> containers = new ArrayDeque<>();
    private boolean boundingBox;
    private boolean declaredBox;
    private boolean headerRegion = true;
    private final Line buffer;
    private boolean atEnd;
    private boolean trailer;
    private int documents;

    Scanner(MediaInput input) {
      in = input;
      buffer = new Line();
    }

    void scan() {
      header();
      while (in.remaining() > 0) {
        int c = in.u8();
        if (space(c)) continue;
        in.budget.token();
        if (c == '%') {
          in.position--;
          comment(line());
          continue;
        }
        headerRegion = false;
        token(c);
      }
      in.require(
          containers.isEmpty() && documents == 0 && boundingBox && (!atEnd || trailer),
          MediaStructureReason.METADATA);
    }

    private void header() {
      Line value = line();
      in.require(value.starts("%!"), MediaStructureReason.HEADER);
      boolean third = value.word("%!PS-Adobe-3.0");
      in.require(third || value.word("%!PS-Adobe-2.0"), MediaStructureReason.HEADER);
      boolean older = value.word("EPSF-1.2") || value.word("EPSF-2.0");
      in.require(older || (third && value.word("EPSF-3.0")), MediaStructureReason.HEADER);
    }

    private void token(int c) {
      switch (c) {
        case '(' -> literal();
        case '<' -> {
          if (peek('<')) {
            in.u8();
            open('>');
          } else if (peek('~')) {
            in.u8();
            ascii85();
          } else hex();
        }
        case '>' -> {
          in.require(peek('>'), MediaStructureReason.SEQUENCE);
          in.u8();
          close('>');
        }
        case '[' -> open(']');
        case '{' -> open('}');
        case ']', '}' -> close(c);
        case ')' -> in.require(false, MediaStructureReason.SEQUENCE);
        default -> {
          while (in.remaining() > 0 && !delimiter(in.bytes[in.position] & 255)) in.u8();
        }
      }
    }

    private void comment(Line value) {
      in.record();
      documentComment(value);
      if (value.starts("%%BeginBinary:")) {
        value.cursor = value.start + 14;
        long bytes = value.unsigned();
        in.skip(bytes);
        endSection("%%EndBinary");
      } else if (value.starts("%%BeginData:")) {
        value.cursor = value.start + 12;
        long count = value.unsigned();
        if (value.contains(" Lines")) {
          in.require(count <= in.remaining(), MediaStructureReason.COUNT);
          for (long i = 0; i < count; i++) line();
        } else {
          in.require(
              value.contains(" Bytes") || value.cursor == value.end, MediaStructureReason.METADATA);
          in.skip(count);
        }
        endSection("%%EndData");
      } else if (value.starts("%%BeginPreview:")) preview(value);
    }

    private void documentComment(Line value) {
      if (value.starts("%%BeginDocument:")) {
        documents++;
        in.budget.depth(in.depth + documents);
        return;
      }
      if (value.keyword("%%EndDocument")) {
        in.require(documents > 0, MediaStructureReason.SEQUENCE);
        documents--;
        return;
      }
      if (documents == 0 && value.keyword("%%EndComments")) headerRegion = false;
      if (documents == 0 && value.keyword("%%Trailer")) {
        trailer = true;
        headerRegion = false;
      }
      if (documents == 0 && value.starts("%%BoundingBox:")) boundingBox(value);
    }

    private void boundingBox(Line value) {
      in.require(headerRegion || (atEnd && trailer), MediaStructureReason.METADATA);
      if (headerRegion) {
        in.require(!declaredBox, MediaStructureReason.SEQUENCE);
        declaredBox = true;
      } else in.require(!boundingBox, MediaStructureReason.SEQUENCE);
      value.cursor = value.start + 14;
      value.whitespace();
      if (value.at("(atend)")) {
        in.require(headerRegion && !trailer, MediaStructureReason.SEQUENCE);
        value.cursor += 7;
        value.whitespace();
        in.require(value.cursor == value.end, MediaStructureReason.METADATA);
        atEnd = true;
        return;
      }
      Decimal x = value.integer();
      Decimal y = value.integer();
      Decimal right = value.integer();
      Decimal top = value.integer();
      value.whitespace();
      in.require(
          value.cursor == value.end && right.compare(x) >= 0 && top.compare(y) >= 0,
          MediaStructureReason.DIMENSIONS);
      boundingBox = true;
    }

    private void endSection(String expected) {
      if (peek(13)) {
        in.u8();
        if (peek(10)) in.u8();
      } else if (peek(10)) in.u8();
      in.record();
      in.require(line().keyword(expected), MediaStructureReason.TERMINATOR);
    }

    private void preview(Line declaration) {
      declaration.cursor = declaration.start + 15;
      long width = declaration.unsigned();
      long height = declaration.unsigned();
      long bits = declaration.unsigned();
      long lines = declaration.unsigned();
      in.require(
          width > 0 && height > 0 && (bits == 1 || bits == 2 || bits == 4 || bits == 8),
          MediaStructureReason.METADATA);
      in.budget.raster(width, height);
      long bytes = in.product((in.product(width, bits) + 7) / 8, height);
      in.budget.expand(bytes);
      in.require(lines <= in.remaining(), MediaStructureReason.COUNT);
      long nibbles = 0;
      for (long row = 0; row < lines; row++) {
        Line data = line();
        in.require(data.starts("%"), MediaStructureReason.SEQUENCE);
        for (int i = data.start + 1; i < data.end; i++) {
          in.budget.inspect(1);
          int c = in.bytes[i] & 255;
          if (space(c)) continue;
          in.require(hexDigit(c), MediaStructureReason.CODE);
          nibbles++;
        }
      }
      in.record();
      in.require(
          nibbles == in.product(bytes, 2) && line().keyword("%%EndPreview"),
          MediaStructureReason.COUNT);
    }

    private void open(int closing) {
      containers.push(closing);
      in.budget.syntaxDepth(containers.size());
    }

    private void close(int closing) {
      in.require(
          !containers.isEmpty() && containers.pop() == closing, MediaStructureReason.SEQUENCE);
    }

    private boolean peek(int c) {
      return in.remaining() > 0 && (in.bytes[in.position] & 255) == c;
    }

    private void literal() {
      int depth = 1;
      in.budget.syntaxDepth(containers.size() + depth);
      while (depth > 0) {
        int c = in.u8();
        if (c == '\\') {
          escape();
        } else if (c == '(') {
          depth++;
          in.budget.syntaxDepth(containers.size() + depth);
        } else if (c == ')') depth--;
      }
    }

    private void escape() {
      int c = in.u8();
      if (c == '\r' && peek('\n')) {
        in.u8();
        return;
      }
      if (c < '0' || c > '7') return;
      for (int i = 0; i < 2 && in.remaining() > 0; i++) {
        int next = in.bytes[in.position] & 255;
        if (next < '0' || next > '7') break;
        in.u8();
      }
    }

    private void hex() {
      while (true) {
        int c = in.u8();
        if (c == '>') return;
        in.require(space(c) || hexDigit(c), MediaStructureReason.CODE);
      }
    }

    private void ascii85() {
      int count = 0;
      long tuple = 0;
      while (true) {
        int c = in.u8();
        if (space(c)) continue;
        if (c == '~') {
          in.require(in.u8() == '>' && count != 1, MediaStructureReason.TERMINATOR);
          if (count > 1) {
            for (int i = count; i < 5; i++) tuple = tuple * 85 + 84;
            in.require(tuple <= 0xffffffffL, MediaStructureReason.CODE);
          }
          return;
        }
        if (c == 'z') {
          in.require(count == 0, MediaStructureReason.CODE);
          continue;
        }
        in.require(c >= '!' && c <= 'u', MediaStructureReason.CODE);
        tuple = tuple * 85 + c - '!';
        count++;
        if (count == 5) {
          in.require(tuple <= 0xffffffffL, MediaStructureReason.CODE);
          count = 0;
          tuple = 0;
        }
      }
    }

    private Line line() {
      int start = in.position;
      while (in.remaining() > 0) {
        int c = in.u8();
        if (c == '\r' || c == '\n') {
          int end = in.position - 1;
          if (c == '\r' && peek('\n')) in.u8();
          return buffer.reset(start, end);
        }
      }
      return buffer.reset(start, in.position);
    }

    /** Zero-copy DSC line and bounded integer cursor. */
    private final class Line {
      int start;
      int end;
      int cursor;

      Line reset(int start, int end) {
        this.start = start;
        this.end = end;
        cursor = start;
        return this;
      }

      boolean word(String expected) {
        whitespace();
        if (!at(expected)) return false;
        int next = cursor + expected.length();
        if (next < end && !space(in.bytes[next] & 255)) return false;
        cursor = next;
        return true;
      }

      boolean at(String expected) {
        if (expected.length() > end - cursor) return false;
        for (int i = 0; i < expected.length(); i++) {
          in.budget.inspect(1);
          if ((in.bytes[cursor + i] & 255) != expected.charAt(i)) return false;
        }
        return true;
      }

      boolean starts(String expected) {
        int save = cursor;
        cursor = start;
        boolean match = at(expected);
        cursor = save;
        return match;
      }

      boolean keyword(String expected) {
        if (!starts(expected)) return false;
        int after = start + expected.length();
        if (after == end) return true;
        in.budget.inspect(1);
        return space(in.bytes[after] & 255);
      }

      boolean contains(String expected) {
        for (int i = start; i <= end - expected.length(); i++) {
          int save = cursor;
          cursor = i;
          boolean match = at(expected);
          cursor = save;
          if (match) return true;
        }
        return false;
      }

      void whitespace() {
        while (cursor < end && space(in.bytes[cursor] & 255)) cursor++;
      }

      long unsigned() {
        Decimal value = integer();
        in.require(!value.negative, MediaStructureReason.COUNT);
        return value.count();
      }

      Decimal integer() {
        whitespace();
        boolean negative = sign();
        int first = cursor;
        digits();
        in.require(
            cursor > first && (cursor == end || space(in.bytes[cursor] & 255)),
            MediaStructureReason.CODE);
        int significant = significant(first);
        boolean zero = cursor - significant == 1 && in.bytes[significant] == '0';
        return new Decimal(negative && !zero, significant, cursor);
      }

      private boolean sign() {
        boolean negative = cursor < end && in.bytes[cursor] == '-';
        if (negative || (cursor < end && in.bytes[cursor] == '+')) cursor++;
        return negative;
      }

      private void digits() {
        while (cursor < end) {
          in.budget.inspect(1);
          int c = in.bytes[cursor] & 255;
          if (c < '0' || c > '9') break;
          cursor++;
        }
      }

      private int significant(int first) {
        int next = first;
        while (next < cursor - 1 && in.bytes[next] == '0') {
          in.budget.inspect(1);
          next++;
        }
        return next;
      }
    }

    /** Exact signed decimal span; DSC vector coordinates need no interpreter-size assumption. */
    private final class Decimal {
      private final boolean negative;
      private final int start;
      private final int end;

      Decimal(boolean negative, int start, int end) {
        this.negative = negative;
        this.start = start;
        this.end = end;
      }

      int compare(Decimal other) {
        if (negative != other.negative) return negative ? -1 : 1;
        int length = Integer.compare(end - start, other.end - other.start);
        if (length != 0) return negative ? -length : length;
        for (int i = 0; i < end - start; i++) {
          in.budget.inspect(2);
          int comparison =
              Integer.compare(in.bytes[start + i] & 255, in.bytes[other.start + i] & 255);
          if (comparison != 0) return negative ? -comparison : comparison;
        }
        return 0;
      }

      long count() {
        long value = 0;
        for (int i = start; i < end; i++) {
          in.budget.inspect(1);
          int digit = (in.bytes[i] & 255) - '0';
          if (value > (Long.MAX_VALUE - digit) / 10) return Long.MAX_VALUE;
          value = value * 10 + digit;
        }
        return value;
      }
    }
  }

  private static boolean space(int c) {
    return c == 0 || c == 9 || c == 10 || c == 12 || c == 13 || c == 32;
  }

  private static boolean delimiter(int c) {
    return space(c) || c == '(' || c == ')' || c == '<' || c == '>' || c == '[' || c == ']'
        || c == '{' || c == '}' || c == '/' || c == '%';
  }

  private static boolean hexDigit(int c) {
    return (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F');
  }
}
