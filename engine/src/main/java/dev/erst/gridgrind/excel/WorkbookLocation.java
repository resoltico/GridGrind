package dev.erst.gridgrind.excel;

import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Describes the filesystem location that relative workbook file hyperlinks resolve against. */
public sealed interface WorkbookLocation
    permits WorkbookLocation.StoredWorkbook, WorkbookLocation.UnsavedWorkbook {
  /** Returns the base directory used to resolve relative file hyperlinks, when one exists. */
  Optional<Path> baseDirectory();

  /** Represents one workbook persisted to one concrete `.xlsx` path on disk. */
  public record StoredWorkbook(Path workbookPath) implements WorkbookLocation {
    public StoredWorkbook {
      Objects.requireNonNull(workbookPath, "workbookPath must not be null");
      workbookPath = workbookPath.toAbsolutePath().normalize();
    }

    @Override
    public Optional<Path> baseDirectory() {
      return Optional.of(workbookPath.getParent());
    }
  }

  /** Represents one workbook that is not anchored to any filesystem path. */
  public record UnsavedWorkbook() implements WorkbookLocation {
    @Override
    public Optional<Path> baseDirectory() {
      return Optional.empty();
    }
  }
}
