package dev.erst.gridgrind.architecture.fixture;

import java.lang.reflect.Field;
import java.util.function.BiConsumer;
import java.util.function.Function;
import org.apache.poi.ss.usermodel.Cell;

/** Deliberately invalid method references used to prove access-site rules inspect invokedynamic. */
public final class ArchitectureMethodReferenceViolationFixture {
  private ArchitectureMethodReferenceViolationFixture() {}

  /** References direct POI formula writing solely for architecture-rule regression. */
  public static BiConsumer<Cell, String> formulaWriter() {
    return Cell::setCellFormula;
  }

  /** References private reflection solely for architecture-rule regression. */
  public static Function<Field, Boolean> privateAccessibilityWriter() {
    return Field::trySetAccessible;
  }

  /** References forced private accessibility solely for architecture-rule regression. */
  public static BiConsumer<Field, Boolean> forcedAccessibilityWriter() {
    return Field::setAccessible;
  }

  /** References Class metadata lookup solely for architecture-rule regression. */
  public static Function<Class<?>, Field[]> declaredFieldsReader() {
    return Class::getDeclaredFields;
  }
}
