package dev.erst.gridgrind.cli.golden;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

/** Owns anonymous Docker configuration and test resources; borrowed images remain external. */
final class DockerFixture implements AutoCloseable {
  private final Path workspace;
  private final List<String> client;
  private final List<String> containers = new ArrayList<>();
  private final List<String> volumes = new ArrayList<>();
  private final String image;
  private final boolean ownsImage;
  private final String platform;
  private final String user;

  DockerFixture(Path workspace) throws java.io.IOException, InterruptedException {
    this(workspace, "gridgrind-test:" + UUID.randomUUID(), true, "");
  }

  DockerFixture(Path workspace, String borrowedImage, String requestedPlatform)
      throws java.io.IOException, InterruptedException {
    this(workspace, borrowedImage, false, requestedPlatform);
  }

  private DockerFixture(
      Path workspace, String requestedImage, boolean ownsImage, String requestedPlatform)
      throws java.io.IOException, InterruptedException {
    if (requestedImage.isBlank() || requestedImage.startsWith("-"))
      throw new IllegalArgumentException("Expected a local Docker image reference");
    this.workspace = workspace;
    this.ownsImage = ownsImage;
    client = anonymousClient(workspace);
    user =
        assertSuccess(GoldenLauncher.runProcess(List.of("id", "-u"), workspace, 60)).strip()
            + ":"
            + assertSuccess(GoldenLauncher.runProcess(List.of("id", "-g"), workspace, 60)).strip();
    String selectedPlatform = validateRequestedPlatform(requestedPlatform);
    // Keep an immutable retained index ID intact; resolve a borrowed mutable tag only once.
    if (ownsImage || requestedImage.matches("sha256:[0-9a-f]{64}")) {
      image = requestedImage;
    } else {
      image = assertSuccess(inspect(requestedImage, "", "{{.Id}}")).strip();
    }
    if (!ownsImage && !image.matches("sha256:[0-9a-f]{64}"))
      throw new IllegalStateException("Docker did not resolve the borrowed local image ID");
    platform =
        ownsImage
            ? ""
            : requirePlatform(
                selectedPlatform,
                assertSuccess(inspect(image, selectedPlatform, "{{.Os}}/{{.Architecture}}"))
                    .strip());
  }

  static String validateRequestedPlatform(String value) {
    return switch (value) {
      case "", "linux/amd64", "linux/arm64" -> value;
      default -> throw new IllegalArgumentException("Unsupported containerPlatform: " + value);
    };
  }

  static String requirePlatform(String requested, String observed) {
    validateRequestedPlatform(requested);
    if (!observed.matches("linux/(amd64|arm64)"))
      throw new IllegalStateException("Unsupported inspected image platform: " + observed);
    if (!requested.isEmpty() && !requested.equals(observed))
      throw new IllegalStateException(
          "Requested image platform " + requested + " but Docker inspected " + observed);
    return observed;
  }

  static void requireJvmPlatform(String expected, String settings) {
    requirePlatform(expected, expected);
    String architecture = null;
    String operatingSystem = null;
    for (String line : settings.lines().toList()) {
      String[] property = line.strip().split("\\s*=\\s*", 2);
      if (property.length != 2) continue;
      if ("os.arch".equals(property[0])) {
        if (architecture != null) throw new IllegalStateException("Repeated JVM os.arch property");
        architecture = property[1];
      } else if ("os.name".equals(property[0])) {
        if (operatingSystem != null)
          throw new IllegalStateException("Repeated JVM os.name property");
        operatingSystem = property[1];
      }
    }
    String jvmArchitecture =
        switch (architecture) {
          case "amd64", "x86_64" -> "amd64";
          case "aarch64", "arm64" -> "arm64";
          case null -> throw new IllegalStateException("JVM os.arch property is missing");
          default ->
              throw new IllegalStateException("Unsupported JVM architecture: " + architecture);
        };
    if (!"Linux".equals(operatingSystem) || !expected.equals("linux/" + jvmArchitecture))
      throw new IllegalStateException(
          "Inspected image platform "
              + expected
              + " disagrees with JVM "
              + operatingSystem
              + "/"
              + architecture);
  }

  private GoldenLauncher.Result inspect(String reference, String selectedPlatform, String format)
      throws java.io.IOException, InterruptedException {
    List<String> arguments = new ArrayList<>(List.of("image", "inspect"));
    if (!selectedPlatform.isEmpty()) arguments.addAll(List.of("--platform", selectedPlatform));
    arguments.addAll(List.of("--format", format, reference));
    return command(60, arguments.toArray(String[]::new));
  }

  String inspectImage(String format) throws java.io.IOException, InterruptedException {
    return assertSuccess(inspect(image, platform, format));
  }

  String inspectedPlatform() throws java.io.IOException, InterruptedException {
    return requirePlatform(platform, inspectImage("{{.Os}}/{{.Architecture}}").strip());
  }

  private static List<String> anonymousClient(Path workspace)
      throws java.io.IOException, InterruptedException {
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
    return List.of("docker", "--config", config.toString(), "--host", endpoint);
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
    if (!ownsImage) throw new IllegalStateException("Cannot rebuild an externally owned image");
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
    return command(120, containerCommand("run", options, arguments).toArray(String[]::new));
  }

  String create(List<String> options, String... arguments)
      throws java.io.IOException, InterruptedException {
    return assertSuccess(
            command(60, containerCommand("create", options, arguments).toArray(String[]::new)))
        .strip();
  }

  private List<String> containerCommand(
      String operation, List<String> options, String... arguments) {
    String container = "gridgrind-test-" + UUID.randomUUID();
    containers.add(container);
    List<String> invocation =
        new ArrayList<>(List.of(operation, "--pull=never", "--name", container));
    if ("run".equals(operation)) invocation.add("--rm");
    if (!platform.isEmpty()) invocation.addAll(List.of("--platform", platform));
    invocation.addAll(options);
    invocation.add(image);
    invocation.addAll(List.of(arguments));
    return invocation;
  }

  GoldenLauncher.Result runWorkbook(Path root, String... arguments)
      throws java.io.IOException, InterruptedException {
    return runWorkbook(List.of(), root, arguments);
  }

  GoldenLauncher.Result runWorkbook(List<String> options, Path root, String... arguments)
      throws java.io.IOException, InterruptedException {
    Path cwd = Files.createDirectories(root.resolve("unrelated working directory"));
    List<String> invocationOptions = new ArrayList<>(options);
    invocationOptions.addAll(
        List.of("--user", user, "--volume", root + ":" + root, "--workdir", cwd.toString()));
    return run(invocationOptions, arguments);
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
    if (ownsImage) command(60, "image", "rm", "--force", image);
  }
}
