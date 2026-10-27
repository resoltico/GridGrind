package dev.erst.gridgrind.architecture.fixture;

import dev.erst.gridgrind.core.metadata.Union;

/** Deliberate noncanonical and absent union tags outside the structural step envelope. */
public enum ArchitectureDiscriminatorViolations {
  ;

  /** Uses a retired discriminator name. */
  @Union(discriminator = "kind")
  public sealed interface WrongKind permits KindCase {}

  public record KindCase() implements WrongKind {}

  /** Claims an unowned structural exception. */
  @Union(discriminator = "")
  public sealed interface WrongEmpty permits EmptyCase {}

  public record EmptyCase() implements WrongEmpty {}
}
