package dev.erst.gridgrind.architecture.fixture;

/** Deliberately invalid non-exception sealed class used to prove the shape rule fails. */
public abstract sealed class ArchitectureSealedClassViolation
    permits ArchitectureSealedClassViolation.OrdinaryVariant {
  /** Identifies the deliberately invalid closed-class variant. */
  public abstract String marker();

  /** Final ordinary subclass that deliberately violates the sealed-class policy. */
  public static final class OrdinaryVariant extends ArchitectureSealedClassViolation {
    @Override
    public String marker() {
      return "ordinary";
    }
  }
}
