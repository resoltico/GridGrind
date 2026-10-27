package dev.erst.gridgrind.architecture.fixture;

import dev.erst.gridgrind.core.metadata.Union;

/** Valid wire union with record leaves behind a sealed branch. */
@Union
public sealed interface ArchitectureWireUnionVariants
    permits ArchitectureWireUnionVariants.RecordBranch {

  /** Closed intermediate branch whose concrete alternative is a record. */
  sealed interface RecordBranch extends ArchitectureWireUnionVariants permits RecordVariant {}

  /** Valid concrete wire alternative. */
  record RecordVariant() implements RecordBranch {}

  /** Deliberately invalid wire union exposing an enum and an open subtype family. */
  @Union
  sealed interface WithEnum permits EnumVariant, OpenBranch {}

  /** An open family cannot enumerate its concrete wire alternatives. */
  non-sealed interface OpenBranch extends WithEnum {}

  /** An enum is not a record wire contract. */
  enum EnumVariant implements WithEnum {
    INVALID
  }
}
