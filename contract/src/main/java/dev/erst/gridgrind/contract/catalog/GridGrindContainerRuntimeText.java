package dev.erst.gridgrind.contract.catalog;

/** Container-runtime wording shared by help surfaces that teach Docker execution. */
public final class GridGrindContainerRuntimeText {
  private static final String CONTAINER_WORKDIR_PATH = "/work";
  private static final String DOCKER_MOUNTED_WORKDIR_USER_ARGUMENT = "--user \"$(id -u):$(id -g)\"";
  private static final String DOCKER_MOUNTED_WORKDIR_VOLUME_ARGUMENT =
      "-v \"$(pwd)\":" + CONTAINER_WORKDIR_PATH;

  private GridGrindContainerRuntimeText() {}

  /** Canonical in-container working directory prepared by the published runtime image. */
  public static String containerWorkdirPath() {
    return CONTAINER_WORKDIR_PATH;
  }

  /** Canonical volume argument for mounting the host working directory into the runtime image. */
  public static String dockerMountedWorkdirVolumeArgument() {
    return DOCKER_MOUNTED_WORKDIR_VOLUME_ARGUMENT;
  }

  /** Canonical user-mapping argument for bind-mounted Docker execution. */
  public static String dockerMountedWorkdirUserArgument() {
    return DOCKER_MOUNTED_WORKDIR_USER_ARGUMENT;
  }

  /** Canonical request/response Docker command for bind-mounted execution. */
  public static String dockerMountedWorkdirExecutionCommand(String containerTag) {
    return "docker run --rm -i "
        + DOCKER_MOUNTED_WORKDIR_USER_ARGUMENT
        + " "
        + DOCKER_MOUNTED_WORKDIR_VOLUME_ARGUMENT
        + " "
        + containerTag
        + " --request request.json --response response.json";
  }

  /** Stable wording for the mounted-directory Docker execution pattern. */
  public static String dockerMountedWorkdirSummary() {
    return "Mount the host working directory at "
        + CONTAINER_WORKDIR_PATH
        + " and rely on the image's prepared WORKDIR so relative CLI paths resolve inside that"
        + " mounted directory without a separate -w override. Pass "
        + DOCKER_MOUNTED_WORKDIR_USER_ARGUMENT
        + " on ordinary bind mounts so response and workbook files stay owned by the calling"
        + " host user; omit it only when Docker Desktop or a rootless runtime already remaps"
        + " bind-mount ownership for you.";
  }
}
