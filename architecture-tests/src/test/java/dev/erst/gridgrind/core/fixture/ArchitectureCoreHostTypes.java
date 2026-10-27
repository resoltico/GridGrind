package dev.erst.gridgrind.core.fixture;

/** Deliberate concrete-host names rejected by the core discovery boundary. */
public interface ArchitectureCoreHostTypes {

  public record CliTransport() {}

  public record DockerInvocation() {}

  public record ContainerImage() {}

  public record TerminalSession() {}
}
