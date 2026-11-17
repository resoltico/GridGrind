package dev.erst.gridgrind.json.codec;

import dev.erst.gridgrind.core.catalog.Catalog;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.RequestDoctorReport;
import dev.erst.gridgrind.core.problem.RequestIntake;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.json.intake.GridGrindRequestDecoder;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.Objects;

/** Shared JSON codec for the GridGrind protocol. */
public final class GridGrindJson {
  private GridGrindJson() {}

  /** Reads a request from an input stream without closing the caller-owned stream. */
  public static WorkbookPlan readRequest(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    byte[] bytes = readRequestBytes(inputStream);
    return readRequest(bytes);
  }

  /** Reads a request from a byte array. */
  public static WorkbookPlan readRequest(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return GridGrindRequestDecoder.read(bytes);
  }

  /** Reads a request from one in-memory JSON string without exposing a checked I/O seam. */
  public static WorkbookPlan readRequest(String json) {
    Objects.requireNonNull(json, "json must not be null");
    return GridGrindRequestDecoder.read(json);
  }

  /** Analyses one request byte stream without discarding valid sibling fragments after a defect. */
  public static RequestIntake analyzeRequest(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return analyzeRequest(readRequestBytes(inputStream));
  }

  /** Reads at most the request limit plus one overflow byte, borrowing the supplied stream. */
  public static byte[] readRequestBytes(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    var bytes = new ByteArrayOutputStream();
    byte[] buffer = new byte[8192];
    long remaining = maxRequestDocumentBytes();
    while (remaining > 0) {
      int read = inputStream.read(buffer, 0, (int) Math.min(buffer.length, remaining));
      if (read < 0) {
        return bytes.toByteArray();
      }
      if (read == 0) {
        int next = inputStream.read();
        if (next < 0) {
          return bytes.toByteArray();
        }
        bytes.write(next);
        remaining--;
      } else {
        bytes.write(buffer, 0, read);
        remaining -= read;
      }
    }
    if (inputStream.read() >= 0) {
      requireSupportedRequestLength(maxRequestDocumentBytes() + 1);
    }
    return bytes.toByteArray();
  }

  /** Analyses one UTF-8 request document into bound fragments and every structural problem. */
  public static RequestIntake analyzeRequest(byte[] bytes) {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return GridGrindRequestDecoder.analyze(bytes);
  }

  /** Reads one explicit authority document using the supplied root for its relative paths. */
  public static ExecutionGrant readGrant(byte[] bytes, Path root) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    Objects.requireNonNull(root, "root must not be null");
    var mapper = GridGrindJsonMapperSupport.GRANT_JSON_MAPPER;
    try (var parser = mapper.createParser(bytes)) {
      return readAuthorityValue(parser, mapper, root);

    } catch (dev.erst.gridgrind.core.problem.InvalidAuthorityException known) {
      throw known;
    } catch (UnexpectedAuthorityReadException unexpected) {
      throw unexpected;
    } catch (RuntimeException unexpected) {
      throw new UnexpectedAuthorityReadException(unexpected);
    }
  }

  private static ExecutionGrant readAuthorityValue(
      tools.jackson.core.JsonParser parser,
      tools.jackson.databind.json.JsonMapper mapper,
      Path root) {
    tools.jackson.core.JsonToken token;
    try {
      token = parser.nextToken();
    } catch (RuntimeException failure) {
      throw AuthorityReadFailures.translate(failure, parser);
    }
    if (token == null || token == tools.jackson.core.JsonToken.VALUE_NULL) {
      throw AuthorityReadFailures.rootRefusal(parser);
    }
    try {
      return Objects.requireNonNull(
          mapper
              .readerFor(ExecutionGrant.class)
              .withAttribute(
                  GrantPathDeserializer.ROOT_ATTRIBUTE, root.toAbsolutePath().normalize())
              .readValue(parser),
          "Authority reader must not return null");
    } catch (RuntimeException failure) {
      throw AuthorityReadFailures.translate(failure, parser);
    }
  }

  /** Reads a response from an input stream without closing the caller-owned stream. */
  public static WorkbookResult readWorkbookResult(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return GridGrindJsonCodecSupport.readValue(
        inputStream,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        WorkbookResult.class,
        JsonReadFailures::syntax);
  }

  /** Reads a response from a byte array. */
  public static WorkbookResult readWorkbookResult(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return GridGrindJsonCodecSupport.readValue(
        bytes,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        WorkbookResult.class,
        JsonReadFailures::syntax);
  }

  /** Reads a protocol catalog from an input stream without closing the caller-owned stream. */
  public static Catalog readProtocolCatalog(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return GridGrindJsonCodecSupport.readValue(
        inputStream,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        Catalog.class,
        JsonReadFailures::syntax);
  }

  /** Reads a protocol catalog from a byte array. */
  public static Catalog readProtocolCatalog(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return GridGrindJsonCodecSupport.readValue(
        bytes, GridGrindJsonMapperSupport.JSON_MAPPER, Catalog.class, JsonReadFailures::syntax);
  }

  /** Reads a request doctor report from an input stream without closing the caller-owned stream. */
  public static RequestDoctorReport readRequestDoctorReport(InputStream inputStream)
      throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return GridGrindJsonCodecSupport.readValue(
        inputStream,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        RequestDoctorReport.class,
        JsonReadFailures::syntax);
  }

  /** Reads a request doctor report from a byte array. */
  public static RequestDoctorReport readRequestDoctorReport(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return GridGrindJsonCodecSupport.readValue(
        bytes,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        RequestDoctorReport.class,
        JsonReadFailures::syntax);
  }

  /** Returns the maximum accepted JSON request document length in bytes. */
  public static long maxRequestDocumentBytes() {
    return GridGrindJsonMapperSupport.maxRequestDocumentBytes();
  }

  /** Rejects one request payload length that exceeds the documented transport limit. */
  public static void requireSupportedRequestLength(long lengthBytes) {
    GridGrindJsonMapperSupport.requireSupportedRequestLength(lengthBytes);
  }
}
