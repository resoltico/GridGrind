package dev.erst.gridgrind.engine.drawing;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;

import com.microsoft.schemas.office.office.SignaturelineDocument;
import dev.erst.gridgrind.core.problem.WorkbookContentIssue;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookContentRefusalException;
import java.util.Optional;
import java.util.stream.Stream;
import org.apache.xmlbeans.XmlException;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;
import org.junit.jupiter.params.provider.ValueSource;

/** VML STTrueFalse is a four-token schema, not generic XML boolean or a default false. */
class VmlTrueFalseReadTest {
  static Stream<Arguments> validTokens() {
    return Stream.of(
        Arguments.of("t", true),
        Arguments.of("true", true),
        Arguments.of("f", false),
        Arguments.of("false", false));
  }

  @ParameterizedTest
  @MethodSource("validTokens")
  void readsAllActualNativeSchemaAlternatives(String token, boolean expected) throws XmlException {
    var line =
        SignaturelineDocument.Factory.parse(
                "<o:signatureline xmlns:o='urn:schemas-microsoft-com:office:office' allowcomments='"
                    + token
                    + "'/>")
            .getSignatureline();
    assertEquals(
        Optional.of(expected),
        SignatureAdapter.signatureLineSetup(line, null, Optional.empty())
            .orElseThrow()
            .allowComments());
  }

  @Test
  void absenceDoesNotFabricateFalse() throws XmlException {
    var line =
        SignaturelineDocument.Factory.parse(
                "<o:signatureline xmlns:o='urn:schemas-microsoft-com:office:office'/>")
            .getSignatureline();
    assertEquals(
        Optional.empty(), SignatureAdapter.signatureLineSetup(line, null, Optional.empty()));
  }

  @ParameterizedTest
  @ValueSource(strings = {"1", "0", "unexpected", ""})
  void invalidTokensBecomeSafeLocatedNativeContentRefusals(String token) throws XmlException {
    var line =
        SignaturelineDocument.Factory.parse(
                "<o:signatureline xmlns:o='urn:schemas-microsoft-com:office:office' allowcomments='"
                    + token
                    + "'/>")
            .getSignatureline();
    var refusal =
        assertThrows(
            WorkbookContentRefusalException.class,
            () ->
                SignatureAdapter.signatureLineSetup(
                    line, null, Optional.of("/xl/drawings/fixture.vml")));
    assertEquals(WorkbookContentIssue.INVALID_NATIVE_VALUE, refusal.issue());
    assertEquals(
        Optional.of("signatureLine.allowComments"), refusal.location().orElseThrow().fieldName());
    assertEquals(
        Optional.of("/xl/drawings/fixture.vml"), refusal.location().orElseThrow().partName());
    if (!token.isEmpty()) assertFalse(refusal.getMessage().contains(token));
  }
}
