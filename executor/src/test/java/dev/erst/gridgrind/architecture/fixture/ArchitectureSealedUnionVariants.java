package dev.erst.gridgrind.architecture.fixture;

/** Deliberately valid variants that cover each allowed sealed-interface implementation shape. */
public sealed interface ArchitectureSealedUnionVariants
    permits ArchitectureSealedUnionVariants.EnumVariant,
        ArchitectureSealedUnionVariants.RecordVariant,
        ArchitectureSealedUnionVariants.SealedSubinterface,
        ArchitectureSealedUnionVariants.TypedExceptionVariant {
  /** Valid finite-enum variant. */
  enum EnumVariant implements ArchitectureSealedUnionVariants {
    INSTANCE
  }

  /** Valid record variant. */
  record RecordVariant(String value) implements ArchitectureSealedUnionVariants {}

  /** Valid sealed subinterface variant. */
  sealed interface SealedSubinterface extends ArchitectureSealedUnionVariants
      permits SealedSubinterface.Leaf {
    /** Valid finite leaf. */
    record Leaf(String value) implements SealedSubinterface {}
  }

  /** Valid typed-exception variant. */
  final class TypedExceptionVariant extends RuntimeException
      implements ArchitectureSealedUnionVariants {
    private static final long serialVersionUID = 1L;
  }
}
