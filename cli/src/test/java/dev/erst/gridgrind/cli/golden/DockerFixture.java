package dev.erst.gridgrind.cli.golden;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

/** Owns one anonymous Docker client configuration, image and any timed-out test containers. */
final class DockerFixture implements AutoCloseable {
  private final Path workspace;
  private final List<String> client;
  private final List<String> containers = new ArrayList<>();
  private final List<String> volumes = new ArrayList<>();
  private final String image = "gridgrind-test:" + UUID.randomUUID();
  private final String user;

  DockerFixture(Path workspace) throws java.io.IOException, InterruptedException {
    this.workspace = workspace;
    Path config = Files.createDirectories(workspace.resolve("anonymous-docker"));
    Files.writeString(config.resolve("config.json"), "{}\n");
    String endpoint = System.getenv("DOCKER_HOST");
    if (endpoint == null || endpoint.isBlank()) {
      endpoint =
          assertSuccess(
                  GoldenLauncher.runProcess(
                      List.of(
                          "docker", "context", "inspect", "--format", "{{.Endpoints.docker.Host}}"),
                      workspace,
                      60))
              .strip();
    }
    var plugins =
        GoldenSnapshots.JSON.readTree(
            assertSuccess(
                GoldenLauncher.runProcess(
                    List.of("docker", "info", "--format", "{{json .ClientInfo.Plugins}}"),
                    workspace,
                    60)));
    String buildx = "";
    for (var plugin : plugins)
      if ("buildx".equals(plugin.path("Name").asText())) buildx = plugin.path("Path").asText();
    if (buildx.isBlank())
      throw new IllegalStateException("docker buildx must be installed for container verification");
    Path pluginDirectory = Files.createDirectories(config.resolve("cli-plugins"));
    Files.createSymbolicLink(
        pluginDirectory.resolve("docker-buildx"), Path.of(buildx).toRealPath());
    client = List.of("docker", "--config", config.toString(), "--host", endpoint);
    user =
        assertSuccess(GoldenLauncher.runProcess(List.of("id", "-u"), workspace, 60)).strip()
            + ":"
            + assertSuccess(GoldenLauncher.runProcess(List.of("id", "-g"), workspace, 60)).strip();
  }

  String image() {
    return image;
  }

  GoldenLauncher.Result command(long timeout, String... arguments)
      throws java.io.IOException, InterruptedException {
    List<String> invocation = new ArrayList<>(client);
    invocation.addAll(List.of(arguments));
    return GoldenLauncher.runProcess(invocation, workspace, timeout);
  }

  void build(Path context, List<String> options) throws java.io.IOException, InterruptedException {
    List<String> arguments = new ArrayList<>(List.of("buildx", "build", "--load", "--tag", image));
    arguments.addAll(options);
    arguments.add(context.toString());
    assertSuccess(command(900, arguments.toArray(String[]::new)));
  }

  String createVolume() throws java.io.IOException, InterruptedException {
    String name = "gridgrind-test-" + UUID.randomUUID();
    volumes.add(name);
    assertSuccess(command(60, "volume", "create", name));
    return name;
  }

  GoldenLauncher.Result run(List<String> options, String... arguments)
      throws java.io.IOException, InterruptedException {
    String container = "gridgrind-test-" + UUID.randomUUID();
    containers.add(container);
    List<String> invocation = new ArrayList<>(List.of("run", "--rm", "--name", container));
    invocation.addAll(options);
    invocation.add(image);
    invocation.addAll(List.of(arguments));
    return command(120, invocation.toArray(String[]::new));
  }

  GoldenLauncher.Result runWorkbook(Path root, String... arguments)
      throws java.io.IOException, InterruptedException {
    Path cwd = Files.createDirectories(root.resolve("unrelated working directory"));
    return run(
        List.of("--user", user, "--volume", root + ":" + root, "--workdir", cwd.toString()),
        arguments);
  }

  static String assertSuccess(GoldenLauncher.Result result) {
    assertEquals(0, result.exit(), () -> result.stdout() + result.stderr());
    return result.stdout();
  }

  @Override
  public void close() throws java.io.IOException, InterruptedException {
    if (!containers.isEmpty()) {
      List<String> remove = new ArrayList<>(List.of("container", "rm", "--force"));
      remove.addAll(containers);
      command(60, remove.toArray(String[]::new));
    }
    if (!volumes.isEmpty()) {
      List<String> remove = new ArrayList<>(List.of("volume", "rm", "--force"));
      remove.addAll(volumes);
      command(60, remove.toArray(String[]::new));
    }
    command(60, "image", "rm", "--force", image);
  }
}
