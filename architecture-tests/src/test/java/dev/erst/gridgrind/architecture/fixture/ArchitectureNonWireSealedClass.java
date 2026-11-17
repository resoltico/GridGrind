package dev.erst.gridgrind.architecture.fixture;

/** A legitimate non-wire sealed class with an ordinary immutable variant. */
public abstract sealed class ArchitectureNonWireSealedClass
    permits ArchitectureNonWireSealedClass.OrdinaryVariant {
  /** Identifies the closed-class variant. */
  public abstract String marker();

  /** An ordinary final class is valid outside a reflective wire union. */
  public static final class OrdinaryVariant extends ArchitectureNonWireSealedClass {
    @Override
    public String marker() {
      return "ordinary";
    }
  }
}
