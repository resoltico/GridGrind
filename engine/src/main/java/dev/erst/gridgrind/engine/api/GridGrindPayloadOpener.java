package dev.erst.gridgrind.engine.api;

import java.io.IOException;
import java.io.InputStream;

/** Lazy host payload capability. An opened stream transfers to, and is closed by, one run. */
@FunctionalInterface
public interface GridGrindPayloadOpener {
  /** Opens the payload only after its read authority has been established. */
  InputStream open() throws IOException;
}
