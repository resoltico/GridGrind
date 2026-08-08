package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.fasterxml.jackson.annotation.JsonIgnore;
import dev.erst.gridgrind.contract.dto.ProtocolField;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import org.junit.jupiter.api.Test;

/** Verifies owner-path diagnostic protection without value-based output corruption. */
class RequestDiagnosticRedactorTest {
  @Test
  void hidesAValidationMessageOnlyWhenItsExactRequestOwnerIsSecret() {
    assertEquals(
        "Sensitive request value is invalid.",
        RequestDiagnosticRedactor.safeValidationFailureMessage(
            "password source-secret is invalid",
            WorkbookPlan.class,
            Optional.of("source.security.password")));
    assertEquals(
        "path source-secret remains an ordinary value",
        RequestDiagnosticRedactor.safeValidationFailureMessage(
            "path source-secret remains an ordinary value",
            WorkbookPlan.class,
            Optional.of("source.path")));
    assertEquals(
        "Request value violates the request contract.",
        RequestDiagnosticRedactor.safeValidationFailureMessage(
            null, WorkbookPlan.class, Optional.of("source.path")));
    assertEquals(
        "Request value violates the request contract.",
        RequestDiagnosticRedactor.safeValidationFailureMessage(
            "", WorkbookPlan.class, Optional.of("source.path")));
  }

  @Test
  void distinguishesSameNamedFieldsByTheirCreatorOwnershipPath() {
    assertEquals(
        "Sensitive request value is invalid.",
        RequestDiagnosticRedactor.safeValidationFailureMessage(
            "secret value", SecretCarrier.class, Optional.of("credential")));
    assertEquals(
        "public value",
        RequestDiagnosticRedactor.safeValidationFailureMessage(
            "public value", SecretCarrier.class, Optional.of("nested.credential")));
  }

  @Test
  void doesNotCorruptUnrelatedTextWhenAValidSecretHasOneCharacter() throws IOException {
    RequestDiagnosticRedactor redactor =
        GridGrindJson.analyzeRequest(
                """
                {
                  "source": {"type":"EXISTING","path":"input.xlsx","security":{"password":"a"}},
                  "persistence": {"type":"NONE"},
                  "protocolVersion":"V2",
                  "steps": []
                }
                """
                    .getBytes(StandardCharsets.UTF_8))
            .diagnosticRedactor();
    byte[] payload =
        """
        {"message":"a harmless path detail","resolution":"Use a valid path","context":{"stage":"READ_REQUEST","json":{"type":"PATH_ONLY","jsonPath":"source.path"}},"causes":[{"message":"a harmless cause"}]}
        """
            .getBytes(StandardCharsets.UTF_8);

    assertEquals(
        new String(payload, StandardCharsets.UTF_8),
        new String(redactor.redactSerializedJson(payload, false), StandardCharsets.UTF_8));
  }

  @Test
  void redactsAProblemOnlyWhenItsStructuredRequestContextNamesASecretField() throws IOException {
    RequestDiagnosticRedactor redactor =
        GridGrindJson.analyzeRequest(
                """
                {
                  "source": {"type":"EXISTING","path":"input.xlsx","security":{"password":"source-secret"}},
                  "persistence": {"type":"NONE"},
                  "protocolVersion":"V2",
                  "steps": []
                }
                """
                    .getBytes(StandardCharsets.UTF_8))
            .diagnosticRedactor();
    byte[] payload =
        """
        {"message":"source-secret escaped","resolution":"retry source-secret","context":{"stage":"READ_REQUEST","json":{"type":"PATH_ONLY","jsonPath":"source.security.password"}},"causes":[{"message":"source-secret cause"}]}
        """
            .getBytes(StandardCharsets.UTF_8);

    String rendered =
        new String(redactor.redactSerializedJson(payload, false), StandardCharsets.UTF_8);
    assertFalse(rendered.contains("source-secret"));
    assertEquals(3, rendered.split("\\[REDACTED]", -1).length - 1);

    String pretty =
        new String(redactor.redactSerializedJson(payload, true), StandardCharsets.UTF_8);
    assertTrue(pretty.startsWith("{\n"));
    assertFalse(pretty.contains("source-secret"));
  }

  @Test
  void leavesNonTextualOrAbsentDiagnosticFieldsUnchanged() throws IOException {
    RequestDiagnosticRedactor redactor =
        RequestDiagnosticRedactor.forRequestType(WorkbookPlan.class);
    byte[] nonTextualPayload =
        """
        {"message":7,"resolution":false,"context":{"json":{"jsonPath":"source.security.password"}},"causes":[{"detail":"no message"},1]}
        """
            .getBytes(StandardCharsets.UTF_8);
    byte[] objectCausesPayload =
        """
        {"message":"sensitive","context":{"json":{"jsonPath":"source.security.password"}},"causes":{}}
        """
            .getBytes(StandardCharsets.UTF_8);

    assertEquals(
        new String(nonTextualPayload, StandardCharsets.UTF_8),
        new String(
            redactor.redactSerializedJson(nonTextualPayload, false), StandardCharsets.UTF_8));
    assertFalse(
        new String(
                redactor.redactSerializedJson(objectCausesPayload, false), StandardCharsets.UTF_8)
            .contains("sensitive"));
  }

  @Test
  void returnsPayloadUnchangedWhenTheRequestTypeHasNoSecretOwners() throws IOException {
    byte[] payload = "{\"message\":\"ordinary\"}".getBytes(StandardCharsets.UTF_8);

    assertEquals(
        new String(payload, StandardCharsets.UTF_8),
        new String(
            RequestDiagnosticRedactor.empty().redactSerializedJson(payload, false),
            StandardCharsets.UTF_8));
  }

  @Test
  void discoversSecretPathsWithoutPromotingIgnoredOrGenericMembers() {
    assertEquals(
        Set.of("published"), RequestSecretFieldPaths.forRequestType(IgnoredSecretCarrier.class));
    assertEquals(Set.of(), RequestSecretFieldPaths.forRequestType(GenericCarrier.class));
    assertEquals(
        Set.of("preferred.credential", "history[].credential"),
        RequestSecretFieldPaths.forRequestType(CollectionSecretCarrier.class));
  }

  private record SecretCarrier(
      @ProtocolField(secret = true) String credential, PublicCarrier nested) {}

  private record PublicCarrier(String credential) {}

  private record IgnoredSecretCarrier(
      @JsonIgnore @ProtocolField(secret = true) String ignored,
      @ProtocolField(secret = true) String published) {}

  private record GenericCarrier<T>(T value, Map<String, String> labels) {}

  private record CollectionSecretCarrier(
      Optional<SecretCarrier> preferred, List<SecretCarrier> history) {}
}
