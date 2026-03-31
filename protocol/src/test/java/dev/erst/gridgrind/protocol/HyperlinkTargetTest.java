package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelHyperlink;
import org.junit.jupiter.api.Test;

/** Tests for HyperlinkTarget record construction and engine conversion. */
class HyperlinkTargetTest {
  @Test
  void buildsSupportedHyperlinkTargets() {
    HyperlinkTarget.Url url = new HyperlinkTarget.Url("https://example.com/report");
    HyperlinkTarget.Email email = new HyperlinkTarget.Email("mailto:team@example.com");
    HyperlinkTarget.Email plainEmail = new HyperlinkTarget.Email("team@example.com");
    HyperlinkTarget.File file = new HyperlinkTarget.File("/tmp/report.xlsx");
    HyperlinkTarget.File fileUri = new HyperlinkTarget.File("file:///tmp/report.xlsx");
    HyperlinkTarget.Document document = new HyperlinkTarget.Document("Budget!B4");

    assertEquals(new ExcelHyperlink.Url("https://example.com/report"), url.toExcelHyperlink());
    assertEquals(new ExcelHyperlink.Email("team@example.com"), email.toExcelHyperlink());
    assertEquals(new ExcelHyperlink.Email("team@example.com"), plainEmail.toExcelHyperlink());
    assertEquals(new ExcelHyperlink.File("/tmp/report.xlsx"), file.toExcelHyperlink());
    assertEquals(new ExcelHyperlink.File("/tmp/report.xlsx"), fileUri.toExcelHyperlink());
    assertEquals("/tmp/report.xlsx", fileUri.path());
    assertEquals(new ExcelHyperlink.Document("Budget!B4"), document.toExcelHyperlink());
  }

  @Test
  void validatesHyperlinkTargetInputs() {
    assertThrows(NullPointerException.class, () -> new HyperlinkTarget.Url(null));
    assertThrows(IllegalArgumentException.class, () -> new HyperlinkTarget.Url("relative/path"));
    assertThrows(
        IllegalArgumentException.class, () -> new HyperlinkTarget.Url("https://exa mple.com"));
    assertThrows(
        IllegalArgumentException.class, () -> new HyperlinkTarget.Url("file:///tmp/report.xlsx"));
    assertThrows(NullPointerException.class, () -> new HyperlinkTarget.Email(null));
    assertThrows(IllegalArgumentException.class, () -> new HyperlinkTarget.Email("mailto:"));
    assertThrows(NullPointerException.class, () -> new HyperlinkTarget.File(null));
    assertThrows(IllegalArgumentException.class, () -> new HyperlinkTarget.File(" "));
    assertThrows(
        IllegalArgumentException.class, () -> new HyperlinkTarget.File("https://example.com"));
    assertThrows(NullPointerException.class, () -> new HyperlinkTarget.Document(null));
    assertThrows(IllegalArgumentException.class, () -> new HyperlinkTarget.Document(" "));
  }

  @Test
  void emailRejectsAddressWithoutAtSign() {
    assertThrows(IllegalArgumentException.class, () -> new HyperlinkTarget.Email("notanemail"));
    assertThrows(
        IllegalArgumentException.class, () -> new HyperlinkTarget.Email("mailto:notanemail"));
  }

  @Test
  void emailRejectsAddressWithEmptyLocalPartOrDomain() {
    assertThrows(IllegalArgumentException.class, () -> new HyperlinkTarget.Email("@example.com"));
    assertThrows(IllegalArgumentException.class, () -> new HyperlinkTarget.Email("user@"));
  }
}
