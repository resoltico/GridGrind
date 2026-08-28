package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.CliTransportNotice;
import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;
import java.nio.file.Files;
import java.nio.file.LinkOption;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.Objects;

/** One create-new response destination held from execution start through payload delivery. */
final class CliResponseReservation implements AutoCloseable {
  private final Path responsePath;
  private final FileChannel channel;

  CliResponseReservation(Path responsePath, FileChannel channel) {
    this.responsePath = Objects.requireNonNull(responsePath, "responsePath must not be null");
    this.channel = Objects.requireNonNull(channel, "channel must not be null");
  }

  static CliResponseReservation reserve(Path requestedPath) throws ResponseReservationException {
    Path target = CliResponseTransportSupport.responseTargetPath(requestedPath);
    try {
      Path parent =
          Objects.requireNonNull(target.getParent(), "responsePath must not be a filesystem root");
      Files.createDirectories(parent);
      FileChannel channel =
          FileChannel.open(
              target,
              StandardOpenOption.CREATE_NEW,
              StandardOpenOption.WRITE,
              LinkOption.NOFOLLOW_LINKS);
      return new CliResponseReservation(target, channel);
    } catch (java.nio.file.FileAlreadyExistsException exception) {
      throw new ResponseReservationException(reservationFailureReason(target), target, exception);
    } catch (IOException exception) {
      throw new ResponseReservationException(
          CliTransportNotice.Reason.RESPONSE_PATH_UNWRITABLE, target, exception);
    }
  }

  private static CliTransportNotice.Reason reservationFailureReason(Path target) {
    if (!Files.exists(target, LinkOption.NOFOLLOW_LINKS)) {
      return CliTransportNotice.Reason.RESPONSE_PATH_UNWRITABLE;
    }
    return Files.isDirectory(target, LinkOption.NOFOLLOW_LINKS)
        ? CliTransportNotice.Reason.RESPONSE_PATH_DIRECTORY
        : CliTransportNotice.Reason.RESPONSE_PATH_EXISTS;
  }

  void write(byte[] payload) throws IOException {
    Objects.requireNonNull(payload, "payload must not be null");
    ByteBuffer buffer = ByteBuffer.wrap(payload);
    while (buffer.hasRemaining()) {
      channel.write(buffer);
    }
    channel.force(true);
  }

  Path responsePath() {
    return responsePath;
  }

  @Override
  public void close() throws IOException {
    channel.close();
  }

  /** Checked reservation failure that preserves one deterministic transport reason. */
  static final class ResponseReservationException extends IOException {
    private static final long serialVersionUID = 1L;

    private final CliTransportNotice.Reason reason;
    private final Path responsePath;

    ResponseReservationException(
        CliTransportNotice.Reason reason, Path responsePath, IOException cause) {
      super("Could not reserve response path: " + responsePath, cause);
      this.reason = Objects.requireNonNull(reason, "reason must not be null");
      this.responsePath = Objects.requireNonNull(responsePath, "responsePath must not be null");
    }

    CliTransportNotice.Reason reason() {
      return reason;
    }

    Path responsePath() {
      return responsePath;
    }
  }
}
