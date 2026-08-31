package dev.erst.gridgrind.architecture.fixture;

import java.lang.reflect.Field;
import java.util.function.Function;

/** Deliberately benign reflection metadata reference used to verify narrow matching. */
public final class ArchitectureBenignReflectionMethodFixture {
  private ArchitectureBenignReflectionMethodFixture() {}

  /** References public field metadata without requesting private access. */
  public static Function<Field, String> fieldNameReader() {
    return Field::getName;
  }
}
