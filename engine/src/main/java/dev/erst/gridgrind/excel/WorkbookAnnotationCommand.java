package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Cell annotation commands for hyperlinks and comments. */
public sealed interface WorkbookAnnotationCommand extends WorkbookCommand
    permits WorkbookAnnotationCommand.SetHyperlink,
        WorkbookAnnotationCommand.ClearHyperlink,
        WorkbookAnnotationCommand.SetComment,
        WorkbookAnnotationCommand.ClearComment {

  /** Replaces the hyperlink attached to a single cell. */
  record SetHyperlink(String sheetName, String address, ExcelHyperlink target)
      implements WorkbookAnnotationCommand {
    public SetHyperlink {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(address, "address must not be null");
      Objects.requireNonNull(target, "target must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
    }
  }

  /** Removes any hyperlink attached to a single existing cell. */
  record ClearHyperlink(String sheetName, String address) implements WorkbookAnnotationCommand {
    public ClearHyperlink {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(address, "address must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
    }
  }

  /** Replaces the plain-text comment attached to a single cell. */
  record SetComment(String sheetName, String address, ExcelComment comment)
      implements WorkbookAnnotationCommand {
    public SetComment {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(address, "address must not be null");
      Objects.requireNonNull(comment, "comment must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
    }
  }

  /** Removes any comment attached to a single existing cell. */
  record ClearComment(String sheetName, String address) implements WorkbookAnnotationCommand {
    public ClearComment {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(address, "address must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
    }
  }
}
