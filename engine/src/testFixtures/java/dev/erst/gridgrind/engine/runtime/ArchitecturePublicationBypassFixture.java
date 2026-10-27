package dev.erst.gridgrind.engine.runtime;

/** Deliberately references the final-publication primitive outside its approved access boundary. */
public final class ArchitecturePublicationBypassFixture {
  private static final Class<RequestPathPublication> BYPASS_TYPE = RequestPathPublication.class;

  /** Exposes the dependency only for bytecode inspection of the publication boundary. */
  public Class<RequestPathPublication> bypassType() {
    return BYPASS_TYPE;
  }
}
