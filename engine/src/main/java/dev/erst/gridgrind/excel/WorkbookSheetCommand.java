package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import java.util.List;
import java.util.Objects;

/** Sheet and workbook state commands that reshape workbook tabs and protection. */
public sealed interface WorkbookSheetCommand extends WorkbookCommand
    permits WorkbookSheetCommand.CreateSheet,
        WorkbookSheetCommand.RenameSheet,
        WorkbookSheetCommand.DeleteSheet,
        WorkbookSheetCommand.MoveSheet,
        WorkbookSheetCommand.CopySheet,
        WorkbookSheetCommand.SetActiveSheet,
        WorkbookSheetCommand.SetSelectedSheets,
        WorkbookSheetCommand.SetSheetVisibility,
        WorkbookSheetCommand.SetSheetProtection,
        WorkbookSheetCommand.ClearSheetProtection,
        WorkbookSheetCommand.SetWorkbookProtection,
        WorkbookSheetCommand.ClearWorkbookProtection {

  record CreateSheet(String sheetName) implements WorkbookSheetCommand {
    public CreateSheet {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Renames an existing sheet to a new destination name. */
  record RenameSheet(String sheetName, String newSheetName) implements WorkbookSheetCommand {
    public RenameSheet {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(newSheetName, "newSheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (newSheetName.isBlank()) {
        throw new IllegalArgumentException("newSheetName must not be blank");
      }
    }
  }

  /** Deletes an existing sheet from the workbook. */
  record DeleteSheet(String sheetName) implements WorkbookSheetCommand {
    public DeleteSheet {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Moves an existing sheet to a zero-based workbook position. */
  record MoveSheet(String sheetName, int targetIndex) implements WorkbookSheetCommand {
    public MoveSheet {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (targetIndex < 0) {
        throw new IllegalArgumentException("targetIndex must not be negative");
      }
    }
  }

  /** Copies one sheet into a new visible, unselected sheet at the requested workbook position. */
  record CopySheet(String sourceSheetName, String newSheetName, ExcelSheetCopyPosition position)
      implements WorkbookSheetCommand {
    public CopySheet {
      Objects.requireNonNull(sourceSheetName, "sourceSheetName must not be null");
      Objects.requireNonNull(newSheetName, "newSheetName must not be null");
      Objects.requireNonNull(position, "position must not be null");
      if (sourceSheetName.isBlank()) {
        throw new IllegalArgumentException("sourceSheetName must not be blank");
      }
      if (newSheetName.isBlank()) {
        throw new IllegalArgumentException("newSheetName must not be blank");
      }
    }
  }

  /** Sets the active sheet and ensures it is selected. */
  record SetActiveSheet(String sheetName) implements WorkbookSheetCommand {
    public SetActiveSheet {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Sets the selected visible sheet set. */
  record SetSelectedSheets(List<String> sheetNames) implements WorkbookSheetCommand {
    public SetSelectedSheets {
      Objects.requireNonNull(sheetNames, "sheetNames must not be null");
      sheetNames = List.copyOf(sheetNames);
      if (sheetNames.isEmpty()) {
        throw new IllegalArgumentException("sheetNames must not be empty");
      }
      for (String sheetName : sheetNames) {
        Objects.requireNonNull(sheetName, "sheetNames must not contain nulls");
        if (sheetName.isBlank()) {
          throw new IllegalArgumentException("sheetNames must not contain blank values");
        }
      }
      if (sheetNames.size() != new java.util.LinkedHashSet<>(sheetNames).size()) {
        throw new IllegalArgumentException("sheetNames must not contain duplicates");
      }
    }
  }

  /** Sets one sheet visibility. */
  record SetSheetVisibility(String sheetName, ExcelSheetVisibility visibility)
      implements WorkbookSheetCommand {
    public SetSheetVisibility {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(visibility, "visibility must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Enables sheet protection with the exact supported lock flags. */
  record SetSheetProtection(
      String sheetName, ExcelSheetProtectionSettings protection, String password)
      implements WorkbookSheetCommand {
    /** Enables sheet protection without applying a password hash. */
    public SetSheetProtection(String sheetName, ExcelSheetProtectionSettings protection) {
      this(sheetName, protection, null);
    }

    public SetSheetProtection {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(protection, "protection must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (password != null && password.isBlank()) {
        throw new IllegalArgumentException("password must not be blank");
      }
    }
  }

  /** Disables sheet protection entirely. */
  record ClearSheetProtection(String sheetName) implements WorkbookSheetCommand {
    public ClearSheetProtection {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Enables workbook-level protection and password hashes with authoritative settings. */
  record SetWorkbookProtection(ExcelWorkbookProtectionSettings protection)
      implements WorkbookSheetCommand {
    public SetWorkbookProtection {
      Objects.requireNonNull(protection, "protection must not be null");
    }
  }

  /** Clears workbook-level protection and password hashes entirely. */
  record ClearWorkbookProtection() implements WorkbookSheetCommand {}
}
