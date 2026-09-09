package dev.erst.gridgrind.architecture.fixture;

/** Provides a public nested type whose package-private owner prevents external API exposure. */
@SuppressWarnings({
  "PMD.PublicMemberInNonPublicType",
  "PMD.MissingStaticMethodInNonInstantiatableClass"
})
final class ArchitectureEnclosingVisibilityFixture {
  private ArchitectureEnclosingVisibilityFixture() {}

  /** Deliberately public nested type inside a package-private owner. */
  public static final class PublicNested {}
}
