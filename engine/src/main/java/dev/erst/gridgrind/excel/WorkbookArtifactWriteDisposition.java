package dev.erst.gridgrind.excel;

/** Explicit file-write ownership for persisted workbook artifacts. */
public enum WorkbookArtifactWriteDisposition {
  /** Create one new target file and fail if the path already exists. */
  CREATE_NEW,
  /** Create or replace the target file explicitly. */
  REPLACE_EXISTING
}
