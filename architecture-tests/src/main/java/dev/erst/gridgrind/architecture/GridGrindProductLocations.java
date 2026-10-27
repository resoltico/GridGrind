package dev.erst.gridgrind.architecture;

import com.tngtech.archunit.core.importer.Location;
import com.tngtech.archunit.junit.LocationProvider;
import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.Set;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/** Imports compiled product roots supplied by the Gradle project model, excluding test fixtures. */
public final class GridGrindProductLocations implements LocationProvider {
  @Override
  public Set<Location> get(Class<?> locationProviderClass) {
    String roots = System.getProperty("gridgrind.architecture.classRoots", "");
    if (roots.isBlank()) {
      throw new IllegalStateException(
          "The architecture task must supply compiled product class roots");
    }
    return Arrays.stream(roots.split(Pattern.quote(File.pathSeparator), -1))
        .map(Path::of)
        .map(
            path -> {
              if (!Files.isDirectory(path)) {
                throw new IllegalStateException("Missing compiled product class root: " + path);
              }
              return Location.of(path);
            })
        .collect(Collectors.toUnmodifiableSet());
  }
}
