package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Exercises exact token resolution and its safe absence cases independently from binding. */
class RequestJsonTokenLocationSupportTest {
  @Test
  void indexesCanonicalPathsWithoutRetainingParserNodes() {
    assertEquals(
        java.util.Map.ofEntries(
            java.util.Map.entry("request", 0L),
            java.util.Map.entry("source", 10L),
            java.util.Map.entry("source.path", 30L),
            java.util.Map.entry("steps", 50L),
            java.util.Map.entry("steps[0]", 70L),
            java.util.Map.entry("steps[0].stepId", 80L),
            java.util.Map.entry("array", 100L),
            java.util.Map.entry("array[0]", 120L),
            java.util.Map.entry("array[0].field", 130L),
            java.util.Map.entry("array[0].field[0]", 150L),
            java.util.Map.entry("scalar", 160L)),
        RequestJsonTokenLocationSupport.index(root()));
  }

  @Test
  void indexKeepsTheFirstDuplicateMemberAndTheRootRequestLocation() {
    RequestJsonObject root =
        new RequestJsonObject(
            0,
            List.of(
                new RequestJsonMember(
                    "source",
                    10,
                    new RequestJsonObject(
                        20,
                        List.of(
                            new RequestJsonMember(
                                "path", 30, new RequestJsonString(40, "first"))))),
                new RequestJsonMember(
                    "source",
                    50,
                    new RequestJsonObject(
                        60,
                        List.of(
                            new RequestJsonMember(
                                "extra", 70, new RequestJsonString(80, "second"))))),
                new RequestJsonMember("request", 90, new RequestJsonString(100, "field")),
                new RequestJsonMember("a.b", 110, new RequestJsonString(120, "ambiguous"))));
    assertEquals(
        java.util.Map.of("request", 0L, "source", 10L, "source.path", 30L),
        RequestJsonTokenLocationSupport.index(root));
  }

  @Test
  void locatesRootMembersNestedMembersAndArrayElements() {
    RequestJsonObject root = root();

    assertEquals(Optional.of(0L), RequestJsonTokenLocationSupport.byteOffsetAt(root, "request"));
    assertEquals(Optional.of(10L), RequestJsonTokenLocationSupport.byteOffsetAt(root, "source"));
    assertEquals(
        Optional.of(30L), RequestJsonTokenLocationSupport.byteOffsetAt(root, "source.path"));
    assertEquals(Optional.of(70L), RequestJsonTokenLocationSupport.byteOffsetAt(root, "steps[0]"));
    assertEquals(
        Optional.of(80L), RequestJsonTokenLocationSupport.byteOffsetAt(root, "steps[0].stepId"));
    assertEquals(
        Optional.of(150L), RequestJsonTokenLocationSupport.byteOffsetAt(root, "array[0].field[0]"));
  }

  @Test
  void returnsEmptyForMissingOrMalformedPathsWithoutInventingALocation() {
    RequestJsonObject root = root();

    for (String path :
        List.of(
            "",
            ".source",
            "missing",
            "source.",
            "scalar.value",
            "steps.value",
            "steps[",
            "steps[entry]",
            "steps[-1]",
            "steps[1]",
            "steps[0]suffix")) {
      assertEquals(
          Optional.empty(), RequestJsonTokenLocationSupport.byteOffsetAt(root, path), path);
    }
  }

  private static RequestJsonObject root() {
    return new RequestJsonObject(
        0,
        List.of(
            new RequestJsonMember(
                "source",
                10,
                new RequestJsonObject(
                    20,
                    List.of(
                        new RequestJsonMember("path", 30, new RequestJsonString(40, "input"))))),
            new RequestJsonMember(
                "steps",
                50,
                new RequestJsonArray(
                    60,
                    List.of(
                        new RequestJsonObject(
                            70,
                            List.of(
                                new RequestJsonMember(
                                    "stepId", 80, new RequestJsonString(90, "first"))))))),
            new RequestJsonMember(
                "array",
                100,
                new RequestJsonArray(
                    110,
                    List.of(
                        new RequestJsonObject(
                            120,
                            List.of(
                                new RequestJsonMember(
                                    "field",
                                    130,
                                    new RequestJsonArray(
                                        140, List.of(new RequestJsonString(150, "value"))))))))),
            new RequestJsonMember("scalar", 160, new RequestJsonString(170, "value"))));
  }
}
