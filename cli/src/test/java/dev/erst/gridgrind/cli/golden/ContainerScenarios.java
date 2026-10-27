package dev.erst.gridgrind.cli.golden;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

/** Carries the original container smoke's path, progress, font and backend contracts. */
final class ContainerScenarios {
  private ContainerScenarios() {}

  static void verify(WorkbookLauncher launcher, Path workspace)
      throws java.io.IOException, InterruptedException {
    Path requests = Files.createDirectories(workspace.resolve("requests odd"));
    for (String name : List.of("create", "reopen", "signature", "streaming", "streaming-read")) {
      Path request = requests.resolve(name + " [docker #smoke].json");
      try (var input =
          ContainerScenarios.class.getResourceAsStream("/container/" + name + ".json")) {
        if (input == null) throw new IllegalStateException("Missing container fixture: " + name);
        Files.write(request, input.readAllBytes());
      }
      var plan = GoldenSnapshots.JSON.readTree(Files.readAllBytes(request));
      Path grant = GoldenWorkbook.grant(plan, requests, name + ".grant.json");
      Path response = workspace.resolve("responses odd/nested/" + name + " [#smoke].json");
      var result =
          launcher.run(
              workspace,
              "run",
              "--request",
              request.toString(),
              "--grant",
              grant.toString(),
              "--response",
              response.toString());
      assertEquals(0, result.exit(), () -> name + ": " + result.stdout() + result.stderr());
      assertEquals("", result.stdout());
      String responseText = Files.readString(response, StandardCharsets.UTF_8);
      var document = GoldenSnapshots.JSON.readTree(responseText);
      assertEquals("SUCCEEDED", document.path("status").asText());
      if ("SAVE_AS".equals(plan.path("persistence").path("type").asText())) {
        String output = plan.path("persistence").path("path").asText();
        assertTrue(Files.isRegularFile(requests.resolve(output)), output);
        assertFalse(
            Files.exists(workspace.resolve(output)), "Output must follow the request directory");
      }
      if ("create".equals(name)) {
        assertEquals("VERBOSE", document.path("journal").path("level").asText());
        assertFalse(document.path("journal").has("events"));
        assertFalse(result.stderr().isBlank(), "VERBOSE progress must reach stderr");
        for (String line : result.stderr().lines().toList()) {
          var event = GoldenSnapshots.JSON.readTree(line);
          assertTrue(
              List.of("STARTED", "SUCCEEDED", "FAILED").contains(event.path("status").asText()));
          assertFalse(event.path("timestamp").asText().isBlank());
          assertFalse(event.path("category").asText().isBlank());
          assertFalse(event.has("detail"));
        }
      } else {
        assertEquals("", result.stderr(), name);
      }
      if ("signature".equals(name)) assertTrue(responseText.contains("BudgetSignature"));
      if ("streaming-read".equals(name)) assertTrue(responseText.contains("Ledger"));
    }
  }
}
