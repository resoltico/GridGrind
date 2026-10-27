package dev.erst.gridgrind.core.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;
import org.junit.jupiter.api.Test;

/** Checks delivered reference coverage separately from the rendering implementation. */
class GridGrindReferenceDocsCoverageTest {
  @Test
  @SuppressWarnings({
    "PMD.UseConcurrentHashMap",
    "PMD.AvoidInstantiatingObjectsInLoops"
  }) // Local ordered index and an independent heading list/set per group.
  void everyDeclaredGroupAndTypeIsPublishedExactlyOnce() throws Exception {
    var catalog = GridGrindProtocolCatalog.catalog();
    Map<String, List<TypeEntry>> groups = new LinkedHashMap<>();
    groups.put("request", List.of(catalog.requestType()));
    catalog.topLevelGroups().forEach(group -> groups.put(group.group(), group.types()));
    catalog.nestedTypes().forEach(group -> groups.put(group.group(), group.types()));
    catalog.plainTypes().forEach(group -> groups.put(group.group(), List.of(group.type())));
    Path root = RepositoryRootTestSupport.repositoryRoot().resolve("docs/reference");
    Pattern heading = Pattern.compile("(?m)^## (\\S+)$");
    for (var group : groups.entrySet()) {
      String text = Files.readString(root.resolve(group.getKey() + ".md"));
      List<String> actual = new ArrayList<>();
      var matcher = heading.matcher(text);
      while (matcher.find()) actual.add(matcher.group(1));
      assertEquals(actual.size(), new LinkedHashSet<>(actual).size(), group.getKey());
      assertEquals(group.getValue().stream().map(TypeEntry::id).toList(), actual, group.getKey());
    }
    assertTrue(Files.readString(root.resolve("README.md")).contains("schema/response.schema.json"));
  }
}
