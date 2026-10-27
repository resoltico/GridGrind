package dev.erst.gridgrind.fuzz.support;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;
import org.slf4j.LoggerFactory;

/** Verifies the fuzz runtime ships one explicit SLF4J provider for deterministic stderr. */
class FuzzLoggingRuntimeTest {
  @Test
  void fuzzRuntime_usesTheCommittedSlf4jProvider() {
    assertEquals(
        "org.apache.logging.slf4j.Log4jLoggerFactory",
        LoggerFactory.getILoggerFactory().getClass().getName());
  }
}
