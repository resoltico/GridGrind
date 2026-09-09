package dev.erst.gridgrind.engine.runtime;

/** Deliberately references the final-publication primitive outside its approved access boundary. */
public final class ArchitecturePublicationBypassFixture {
  private static final Class<RequestPathPublication> BYPASS_TYPE = RequestPathPublication.class;

  public Class<RequestPathPublication> bypassType() {
    return BYPASS_TYPE;
  }
}
