package dev.erst.gridgrind.architecture.fixture;

/** Deliberately invalid non-record variant used to prove the sealed-union rule fails. */
public sealed interface ArchitectureSealedUnionViolation
    permits ArchitectureSealedUnionViolation.NonSealedVariant,
        ArchitectureSealedUnionViolation.OrdinaryVariant {
  /** Non-sealed interface that deliberately violates GridGrind's closed data-union shape. */
  non-sealed interface NonSealedVariant extends ArchitectureSealedUnionViolation {}

  /** Final ordinary class that deliberately violates GridGrind's closed data-union shape. */
  final class OrdinaryVariant implements ArchitectureSealedUnionViolation {}
}
