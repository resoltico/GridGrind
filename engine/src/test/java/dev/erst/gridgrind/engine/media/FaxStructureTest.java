package dev.erst.gridgrind.engine.media;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import org.junit.jupiter.api.Test;

/** Independent ITU-T T.4 table5literal codes and TIFF6 strip framing controls. */
class FaxStructureTest {
  private static final String EOL = "000000000001";

  @Test
  void uncompressedExtensionWorksInBothCodingModesAndFillOrders() {
    // Literal001 gives white,white,black; exit code adds two whites and restoreswhite runs.
    String literal = "0010000000010";
    for (int fill : new int[] {1, 2}) {
      assertDoesNotThrow(() -> check(EOL + "000000001111" + literal, 5, 1, 3, 2, fill));
      assertDoesNotThrow(() -> check("0000001111" + literal + EOL + EOL, 5, 1, 4, 2, fill));
      assertDoesNotThrow(
          () -> check(EOL + "1" + "1100" + EOL + "0" + "0000001111" + literal, 5, 2, 3, 3, fill));
    }
  }

  @Test
  void malformedCodesCountsOptionsAndEarlyTerminatorsRefuse() {
    assertThrows(
        InvalidRequestException.class, () -> check(EOL + "000000001111001", 5, 1, 3, 2, 1));
    assertThrows(
        InvalidRequestException.class,
        () -> check(EOL + "0000000011110010000000010", 4, 1, 3, 2, 1));
    assertThrows(
        InvalidRequestException.class,
        () -> check(EOL + "0000000011110010000000010", 5, 1, 3, 0, 1));
    assertThrows(InvalidRequestException.class, () -> check(EOL + "0" + "1", 5, 1, 3, 1, 1));
    assertThrows(InvalidRequestException.class, () -> check(EOL + EOL, 5, 1, 4, 0, 1));
    assertThrows(InvalidRequestException.class, () -> check("1", 5, 1, 4, 1, 1));
    assertThrows(
        InvalidRequestException.class, () -> check("0000001111" + "000000000001", 5, 1, 4, 2, 1));
  }

  @Test
  void modifiedHuffmanStartsEachRowOnAByteBoundary() {
    // White5terminator is1100; ignored row padding is deliberately nonzero.
    assertDoesNotThrow(() -> check("1100111111001111", 5, 2, 2, 0, 1));
    assertThrows(InvalidRequestException.class, () -> check("11000000", 6, 1, 2, 0, 1));
  }

  private static void check(
      String encoded, long width, long rows, int compression, long options, int fill) {
    byte[] data = new byte[(encoded.length() + 7) / 8];
    for (int i = 0; i < encoded.length(); i++)
      if (encoded.charAt(i) == '1') data[i / 8] |= (byte) (1 << (fill == 1 ? 7 - i % 8 : i % 8));
    var input =
        new MediaInput(
            data, 0, data.length, ExcelPictureFormat.TIFF, new MediaAdmissionBudget(), 0);
    FaxStructure.check(input, width, rows, compression, options, fill);
  }
}
