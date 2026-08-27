package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.cli.discovery.CliTransportNotice;
import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResults;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.PosixFilePermission;
import java.util.Set;
import java.util.concurrent.atomic.AtomicInteger;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies response-file reservation blocks execution before any workbook side effect. */
class GridGrindCliResponseReservationTest {
  @TempDir Path temporaryDirectory;

  @Test
  void existingResponsePathPreventsWorkbookExecutionAndLeavesEveryTargetUntouched()
      throws Exception {
    Path responsePath = temporaryDirectory.resolve("response.json");
    Files.writeString(responsePath, "sentinel\n");
    AtomicInteger executorCalls = new AtomicInteger();
    GridGrindCli cli =
        GridGrindCli.forTesting(
            (request, bindings, progress) -> {
              executorCalls.incrementAndGet();
              throw new AssertionError("response reservation must prevent execution");
            });
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        cli.run(
            new String[] {
              "--execution-root",
              temporaryDirectory.toString(),
              "--response",
              responsePath.toString()
            },
            new ByteArrayInputStream(
                """
                {
                  "protocolVersion": "V2",
                  "source": { "type": "NEW" },
                  "persistence": { "type": "SAVE_AS", "path": "output.xlsx", "ifExists": "REJECT" },
                  "steps": []
                }
                """
                    .getBytes(StandardCharsets.UTF_8)),
            stdout,
            stderr);

    CliTransportNotice notice =
        GridGrindCliJson.readBytes(stderr.toByteArray(), CliTransportNotice.class);
    assertEquals(1, exitCode);
    assertEquals(0, executorCalls.get());
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertEquals("sentinel\n", Files.readString(responsePath));
    assertFalse(Files.exists(temporaryDirectory.resolve("output.xlsx")));
    assertEquals(CliTransportNotice.Reason.RESPONSE_PATH_EXISTS, notice.reason());
    assertEquals(CliTransportNotice.Destination.NOT_DELIVERED, notice.wroteTo());
    assertEquals(
        java.util.Optional.of(responsePath.toAbsolutePath().toString()), notice.responsePath());
  }

  @Test
  void leadingUtf8BomSurvivesAsOneStructuredExecutionWarning() throws Exception {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    byte[] body =
        """
        {
          "protocolVersion": "V2",
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "steps": []
        }
        """
            .getBytes(StandardCharsets.UTF_8);
    byte[] withBom = new byte[body.length + 3];
    withBom[0] = (byte) 0xEF;
    withBom[1] = (byte) 0xBB;
    withBom[2] = (byte) 0xBF;
    System.arraycopy(body, 0, withBom, 3, body.length);

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--execution-root", temporaryDirectory.toString()},
                new ByteArrayInputStream(withBom),
                stdout,
                stderr);

    dev.erst.gridgrind.contract.dto.WorkbookResult.Success response =
        org.junit.jupiter.api.Assertions.assertInstanceOf(
            dev.erst.gridgrind.contract.dto.WorkbookResult.Success.class,
            dev.erst.gridgrind.contract.json.GridGrindJson.readWorkbookResult(
                stdout.toByteArray()));
    assertEquals(0, exitCode);
    assertEquals(1, response.warnings().size());
    assertEquals(
        dev.erst.gridgrind.contract.dto.GridGrindWarningCode.UTF8_BOM_IGNORED,
        response.warnings().getFirst().code());
    assertEquals("", stderr.toString(StandardCharsets.UTF_8));
  }

  @Test
  void leadingUtf8BomSurvivesAsOneStructuredDoctorWarning() throws Exception {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    byte[] body =
        """
        {
          "protocolVersion": "V2",
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "steps": []
        }
        """
            .getBytes(StandardCharsets.UTF_8);
    byte[] withBom = new byte[body.length + 3];
    withBom[0] = (byte) 0xEF;
    withBom[1] = (byte) 0xBB;
    withBom[2] = (byte) 0xBF;
    System.arraycopy(body, 0, withBom, 3, body.length);

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--doctor-request", "--execution-root", temporaryDirectory.toString()
                },
                new ByteArrayInputStream(withBom),
                stdout,
                stderr);

    dev.erst.gridgrind.contract.dto.RequestDoctorReport report =
        dev.erst.gridgrind.contract.json.GridGrindJson.readRequestDoctorReport(
            stdout.toByteArray());
    assertEquals(0, exitCode);
    assertEquals(1, report.warnings().size());
    assertEquals(
        dev.erst.gridgrind.contract.dto.GridGrindWarningCode.UTF8_BOM_IGNORED,
        report.warnings().getFirst().code());
    assertEquals("", stderr.toString(StandardCharsets.UTF_8));
  }

  @Test
  void reservationClassifiesUnwritableParentsAndToleratesAnUnavailableNoticeStream()
      throws Exception {
    Path parentFile = Files.createTempFile(temporaryDirectory, "response-parent-", ".txt");
    CliResponseReservation.ResponseReservationException unwritable =
        assertThrows(
            CliResponseReservation.ResponseReservationException.class,
            () -> CliResponseReservation.reserve(parentFile.resolve("response.json")));
    assertEquals(CliTransportNotice.Reason.RESPONSE_PATH_UNWRITABLE, unwritable.reason());

    Path readOnlyDirectory = Files.createDirectory(temporaryDirectory.resolve("read-only"));
    Files.setPosixFilePermissions(
        readOnlyDirectory,
        Set.of(
            PosixFilePermission.OWNER_READ,
            PosixFilePermission.OWNER_EXECUTE,
            PosixFilePermission.GROUP_READ,
            PosixFilePermission.GROUP_EXECUTE,
            PosixFilePermission.OTHERS_READ,
            PosixFilePermission.OTHERS_EXECUTE));
    try {
      CliResponseReservation.ResponseReservationException denied =
          assertThrows(
              CliResponseReservation.ResponseReservationException.class,
              () -> CliResponseReservation.reserve(readOnlyDirectory.resolve("response.json")));
      assertEquals(CliTransportNotice.Reason.RESPONSE_PATH_UNWRITABLE, denied.reason());
    } finally {
      Files.setPosixFilePermissions(
          readOnlyDirectory,
          Set.of(
              PosixFilePermission.OWNER_READ,
              PosixFilePermission.OWNER_WRITE,
              PosixFilePermission.OWNER_EXECUTE));
    }

    Path responsePath = temporaryDirectory.resolve("existing-response.json");
    Files.writeString(responsePath, "sentinel\n");
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--execution-root",
                  temporaryDirectory.toString(),
                  "--response",
                  responsePath.toString()
                },
                new ByteArrayInputStream(validRequestBytes()),
                new ByteArrayOutputStream(),
                failingOutputStream());
    assertEquals(1, exitCode);
    assertEquals("sentinel\n", Files.readString(responsePath));
  }

  @Test
  void executionAndDoctorResponsesRetainBomWarningsWhenTheirPrimaryResultIsInvalid()
      throws Exception {
    GridGrindCli failingCli =
        GridGrindCli.forTesting(
            (request, bindings, progress) ->
                CliExecutionFailureSupport.failure(request, new IllegalStateException("failed")));
    ByteArrayOutputStream executionStdout = new ByteArrayOutputStream();
    int executionExitCode =
        failingCli.run(
            new String[] {"--execution-root", temporaryDirectory.toString()},
            new ByteArrayInputStream(withBom(validRequestBytes())),
            executionStdout,
            new ByteArrayOutputStream());
    WorkbookResult.Failure failure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            dev.erst.gridgrind.contract.json.GridGrindJson.readWorkbookResult(
                executionStdout.toByteArray()));
    assertEquals(1, executionExitCode);
    assertEquals(
        dev.erst.gridgrind.contract.dto.GridGrindWarningCode.UTF8_BOM_IGNORED,
        failure.warnings().getFirst().code());

    ByteArrayOutputStream doctorStdout = new ByteArrayOutputStream();
    int doctorExitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--doctor-request", "--execution-root", temporaryDirectory.toString()
                },
                new ByteArrayInputStream(withBom("{}".getBytes(StandardCharsets.UTF_8))),
                doctorStdout,
                new ByteArrayOutputStream());
    dev.erst.gridgrind.contract.dto.RequestDoctorReport doctor =
        dev.erst.gridgrind.contract.json.GridGrindJson.readRequestDoctorReport(
            doctorStdout.toByteArray());
    assertEquals(1, doctorExitCode);
    assertFalse(doctor.valid());
    assertEquals(
        dev.erst.gridgrind.contract.dto.GridGrindWarningCode.UTF8_BOM_IGNORED,
        doctor.warnings().getFirst().code());
  }

  @Test
  void bindingConstructionFailureUsesTheAlreadyReservedResponseDescriptor() throws Exception {
    Path tempRootFile = Files.createTempFile(temporaryDirectory, "not-a-directory-", ".tmp");
    Path responsePath = temporaryDirectory.resolve("response.json");
    AtomicInteger executorCalls = new AtomicInteger();
    GridGrindCli cli =
        GridGrindCli.forTesting(
            (request, bindings, progress) -> {
              executorCalls.incrementAndGet();
              return WorkbookResults.success(
                  java.util.List.of(), java.util.List.of(), java.util.List.of());
            });

    int exitCode =
        cli.run(
            new String[] {
              "--execution-root",
              temporaryDirectory.toString(),
              "--temp-root",
              tempRootFile.toString(),
              "--response",
              responsePath.toString()
            },
            new ByteArrayInputStream(validRequestBytes()),
            new ByteArrayOutputStream(),
            new ByteArrayOutputStream());

    WorkbookResult.Failure failure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            dev.erst.gridgrind.contract.json.GridGrindJson.readWorkbookResult(
                Files.readAllBytes(responsePath)));
    assertEquals(1, exitCode);
    assertEquals(0, executorCalls.get());
    assertEquals("EXECUTE_REQUEST", failure.problem().context().stage());
  }

  private static byte[] validRequestBytes() {
    return """
        {
          "protocolVersion": "V2",
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "steps": []
        }
        """
        .getBytes(StandardCharsets.UTF_8);
  }

  private static byte[] withBom(byte[] body) {
    byte[] withBom = new byte[body.length + 3];
    withBom[0] = (byte) 0xEF;
    withBom[1] = (byte) 0xBB;
    withBom[2] = (byte) 0xBF;
    System.arraycopy(body, 0, withBom, 3, body.length);
    return withBom;
  }

  private static OutputStream failingOutputStream() {
    return new OutputStream() {
      @Override
      public void write(int ignored) throws IOException {
        throw new IOException("test output failure");
      }
    };
  }
}
