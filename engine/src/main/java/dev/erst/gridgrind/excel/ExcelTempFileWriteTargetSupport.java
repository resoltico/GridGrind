package dev.erst.gridgrind.excel;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;

/** Converts reserved temp-file paths into fresh workbook write targets. */
public final class ExcelTempFileWriteTargetSupport {
  private ExcelTempFileWriteTargetSupport() {}

  /**
   * Temp-file factories reserve a unique path by creating the file. Workbook writers that save with
   * CREATE_NEW need the path to be absent again before they take ownership of it.
   */
  public static Path prepareCreateNewTarget(Path reservedPath) throws IOException {
    Objects.requireNonNull(reservedPath, "reservedPath must not be null");
    Files.deleteIfExists(reservedPath);
    return reservedPath;
  }
}
