package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.TableInput;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Exhaustive contract coverage for the bounded scalar-constraint vocabulary. */
class FieldConstraintTest {
  @Test
  void validatesAndOrdersEveryConstraintVariant() {
    List<FieldConstraint> constraints =
        List.of(
            new FieldConstraint.NonBlank(),
            new FieldConstraint.StringPattern("x+"),
            new FieldConstraint.LengthRange(0, 1),
            new FieldConstraint.NumberRange(0, 1),
            new FieldConstraint.Integral(),
            new FieldConstraint.PathSuffix(".xlsx"));
    assertEquals(
        List.of(
            "NON_BLANK",
            "STRING_PATTERN",
            "LENGTH_RANGE",
            "NUMBER_RANGE",
            "INTEGRAL",
            "PATH_SUFFIX"),
        constraints.stream().map(FieldConstraint::type).toList());
    assertEquals(
        List.of("", "x+", "0:1", "0.0:1.0", "", ".xlsx"),
        constraints.stream().map(FieldConstraint::sortKey).toList());
    assertThrows(IllegalArgumentException.class, () -> new FieldConstraint.LengthRange(-1, 0));
    assertThrows(IllegalArgumentException.class, () -> new FieldConstraint.LengthRange(2, 1));
    assertThrows(
        IllegalArgumentException.class, () -> new FieldConstraint.NumberRange(Double.NaN, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new FieldConstraint.NumberRange(1, Double.POSITIVE_INFINITY));
    assertThrows(IllegalArgumentException.class, () -> new FieldConstraint.NumberRange(2, 1));
    assertThrows(IllegalArgumentException.class, () -> new FieldConstraint.PathSuffix("xlsx"));
  }

  @Test
  void typeEntryConvenienceConstructorsRetainEmptyPreconditions() {
    TypeEntry entry =
        new TypeEntry(
            "TYPE", "summary", List.of(), List.of(), Optional.empty(), List.of(), Optional.empty());

    assertEquals(List.of(), entry.preconditions());
    assertEquals(Optional.empty(), entry.stepTemplate());
  }

  @Test
  void typeEntryNoteOnlyConstructorRetainsEmptyPreconditionsAndTemplate() {
    TypeEntry entry =
        new TypeEntry("TYPE", "summary", List.of(), List.of(), Optional.empty(), List.of());

    assertEquals(List.of(), entry.preconditions());
    assertEquals(Optional.empty(), entry.stepTemplate());
  }

  @Test
  void catalogDefinedNameConstraintsApplyOnlyToTheThreeDefinedNameOwners() {
    assertEquals(
        2,
        CatalogFieldConstraints.forComponent(StructuredMutationAction.SetNamedRange.class, "name")
            .size());
    assertEquals(2, CatalogFieldConstraints.forComponent(TableInput.class, "name").size());
    assertEquals(2, CatalogFieldConstraints.forComponent(NamedRangeTarget.class, "name").size());
    assertEquals(
        List.of(),
        CatalogFieldConstraints.forComponent(
            StructuredMutationAction.SetNamedRange.class, "range"));
    assertEquals(List.of(), CatalogFieldConstraints.forComponent(TableInput.class, "range"));
  }
}
