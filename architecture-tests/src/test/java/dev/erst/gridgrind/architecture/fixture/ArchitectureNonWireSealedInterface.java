package dev.erst.gridgrind.architecture.fixture;

/** A non-wire closed interface may permit ordinary classes and open extension points. */
public sealed interface ArchitectureNonWireSealedInterface
    permits ArchitectureNonWireSealedInterface.NonSealedVariant,
        ArchitectureNonWireSealedInterface.OrdinaryVariant {
  /** An open extension point outside the wire model. */
  non-sealed interface NonSealedVariant extends ArchitectureNonWireSealedInterface {}

  /** An ordinary immutable implementation outside the wire model. */
  final class OrdinaryVariant implements ArchitectureNonWireSealedInterface {}
}
