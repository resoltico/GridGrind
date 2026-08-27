package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.dto.RequestWarning;
import java.nio.ByteBuffer;
import java.nio.CharBuffer;
import java.nio.charset.CharacterCodingException;
import java.nio.charset.CodingErrorAction;
import java.nio.charset.StandardCharsets;
import java.util.List;

/** Strictly decodes the request transport bytes and preserves their original byte positions. */
final class RequestUtf8Decoder {
  private RequestUtf8Decoder() {}

  static RequestUtf8DecodeResult decode(byte[] bytes) {
    int initialOffset = hasLeadingBom(bytes) ? 3 : 0;
    ByteBuffer input = ByteBuffer.wrap(bytes, initialOffset, bytes.length - initialOffset);
    try {
      CharBuffer characters =
          StandardCharsets.UTF_8
              .newDecoder()
              .onMalformedInput(CodingErrorAction.REPORT)
              .onUnmappableCharacter(CodingErrorAction.REPORT)
              .decode(input);
      String text = characters.toString();
      return new RequestUtf8DecodeResult(
          text,
          byteOffsets(text, initialOffset),
          List.of(),
          initialOffset == 0 ? List.of() : List.of(RequestWarning.utf8BomIgnored()));
    } catch (CharacterCodingException exception) {
      return new RequestUtf8DecodeResult(
          "",
          new long[] {0},
          List.of(
              new RequestInvalidEncoding(
                  "Request bytes must be valid UTF-8", initialOffset + input.position())),
          List.of());
    }
  }

  private static boolean hasLeadingBom(byte[] bytes) {
    return bytes.length >= 3
        && Byte.toUnsignedInt(bytes[0]) == 0xEF
        && Byte.toUnsignedInt(bytes[1]) == 0xBB
        && Byte.toUnsignedInt(bytes[2]) == 0xBF;
  }

  private static long[] byteOffsets(String text, long initialOffset) {
    long[] offsets = new long[text.length() + 1];
    long byteOffset = initialOffset;
    int index = 0;
    while (index < text.length()) {
      offsets[index] = byteOffset;
      int codePoint = text.codePointAt(index);
      int characterCount = Character.charCount(codePoint);
      if (characterCount == 2) {
        offsets[index + 1] = byteOffset;
      }
      byteOffset += utf8Length(codePoint);
      index += characterCount;
    }
    offsets[text.length()] = byteOffset;
    return offsets;
  }

  private static int utf8Length(int codePoint) {
    if (codePoint <= 0x7f) {
      return 1;
    }
    if (codePoint <= 0x7ff) {
      return 2;
    }
    return codePoint <= 0xffff ? 3 : 4;
  }
}
