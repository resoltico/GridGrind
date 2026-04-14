package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;
import org.slf4j.LoggerFactory;

/** Verifies the Jazzer runtime ships one explicit SLF4J provider for deterministic stderr. */
class JazzerLoggingRuntimeTest {
  @Test
  void jazzerRuntime_usesTheCommittedSlf4jProvider() {
    assertEquals(
        "org.apache.logging.slf4j.Log4jLoggerFactory",
        LoggerFactory.getILoggerFactory().getClass().getName());
  }
}
