package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Shares workbook-level sheet validation and view-state normalization helpers. */
final class ExcelWorkbookSheetSupport {
  private ExcelWorkbookSheetSupport() {}

  /**
   * Returns the active visible-or-primary sheet name for a workbook that already contains sheets.
   */
  static String activeSheetName(XSSFWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    if (workbook.getNumberOfSheets() == 0) {
      throw new IllegalArgumentException("workbook must contain at least one sheet");
    }
    int activeSheetIndex = workbook.getActiveSheetIndex();
    if (activeSheetIndex < 0 || activeSheetIndex >= workbook.getNumberOfSheets()) {
      return workbook.getSheetName(0);
    }
    return workbook.getSheetName(activeSheetIndex);
  }

  /** Returns the selected sheet names in workbook order, falling back to the active sheet. */
  static List<String> selectedSheetNames(XSSFWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    if (workbook.getNumberOfSheets() == 0) {
      return List.of();
    }

    List<String> selectedSheetNames = new ArrayList<>();
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      if (workbook.getSheetAt(sheetIndex).isSelected()) {
        selectedSheetNames.add(workbook.getSheetName(sheetIndex));
      }
    }
    if (!selectedSheetNames.isEmpty()) {
      return List.copyOf(selectedSheetNames);
    }
    return List.of(activeSheetName(workbook));
  }

  /** Normalizes active and selected sheet state so it remains valid and visible. */
  static void normalizeWorkbookViewState(XSSFWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");

    int activeSheetIndex = workbook.getActiveSheetIndex();
    if (activeSheetIndex < 0
        || activeSheetIndex >= workbook.getNumberOfSheets()
        || workbook.getSheetVisibility(activeSheetIndex) != SheetVisibility.VISIBLE) {
      activeSheetIndex = firstVisibleSheetIndex(workbook);
      workbook.setActiveSheet(activeSheetIndex);
    }

    boolean hasVisibleSelection = false;
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      if (workbook.getSheetVisibility(sheetIndex) != SheetVisibility.VISIBLE
          && workbook.getSheetAt(sheetIndex).isSelected()) {
        workbook.getSheetAt(sheetIndex).setSelected(false);
      }
      if (workbook.getSheetVisibility(sheetIndex) == SheetVisibility.VISIBLE
          && workbook.getSheetAt(sheetIndex).isSelected()) {
        hasVisibleSelection = true;
      }
    }

    workbook.getSheetAt(activeSheetIndex).setSelected(true);
    if (!hasVisibleSelection) {
      workbook.getSheetAt(activeSheetIndex).setSelected(true);
    }
  }

  /** Returns the first visible sheet index, or fails when the workbook has no visible sheets. */
  static int firstVisibleSheetIndex(XSSFWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      if (workbook.getSheetVisibility(sheetIndex) == SheetVisibility.VISIBLE) {
        return sheetIndex;
      }
    }
    throw new IllegalStateException("workbook must contain at least one visible sheet");
  }

  /** Returns the count of visible sheets in one workbook. */
  static int visibleSheetCount(XSSFWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    int visibleSheetCount = 0;
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      if (workbook.getSheetVisibility(sheetIndex) == SheetVisibility.VISIBLE) {
        visibleSheetCount++;
      }
    }
    return visibleSheetCount;
  }

  /** Fails when the referenced visible sheet is the only visible sheet left in the workbook. */
  static void requireNotLastVisibleSheet(XSSFWorkbook workbook, int sheetIndex, String message) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(message, "message must not be null");
    if (workbook.getSheetVisibility(sheetIndex) == SheetVisibility.VISIBLE
        && visibleSheetCount(workbook) == 1) {
      throw new IllegalArgumentException(message);
    }
  }

  /** Returns the requested primary selection sheet index after validating it exists. */
  static int requiredSelectedSheetIndex(XSSFWorkbook workbook, String requestedPrimarySheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    requireSheetName(requestedPrimarySheetName, "requestedPrimarySheetName");

    int sheetIndex = workbook.getSheetIndex(requestedPrimarySheetName);
    if (sheetIndex < 0) {
      throw new IllegalArgumentException(
          "selected sheets must contain at least one existing sheet");
    }
    return sheetIndex;
  }

  /** Fails when the referenced sheet is not visible. */
  static void requireVisibleSheet(XSSFWorkbook workbook, int sheetIndex, String message) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(message, "message must not be null");
    if (workbook.getSheetVisibility(sheetIndex) != SheetVisibility.VISIBLE) {
      throw new IllegalArgumentException(message);
    }
  }

  /** Returns a validated copy of the requested selected sheet names. */
  static List<String> copySelectedSheetNames(List<String> sheetNames) {
    Objects.requireNonNull(sheetNames, "sheetNames must not be null");
    List<String> copy = List.copyOf(sheetNames);
    if (copy.isEmpty()) {
      throw new IllegalArgumentException("sheetNames must not be empty");
    }
    Set<String> seen = new LinkedHashSet<>();
    for (String sheetName : copy) {
      requireSheetName(sheetName, "sheetNames");
      if (!seen.add(sheetName)) {
        throw new IllegalArgumentException("sheetNames must not contain duplicates");
      }
    }
    return copy;
  }

  /** Returns the existing sheet with the requested name. */
  static XSSFSheet requiredSheet(XSSFWorkbook workbook, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    requireSheetName(sheetName, "sheetName");

    XSSFSheet sheet = workbook.getSheet(sheetName);
    if (sheet == null) {
      throw new SheetNotFoundException(sheetName);
    }
    return sheet;
  }

  /** Returns the zero-based index of an existing sheet. */
  static int requiredSheetIndex(XSSFWorkbook workbook, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    requireSheetName(sheetName, "sheetName");

    int sheetIndex = workbook.getSheetIndex(sheetName);
    if (sheetIndex < 0) {
      throw new SheetNotFoundException(sheetName);
    }
    return sheetIndex;
  }

  /** Fails when another sheet already owns the requested new name. */
  static void requireSheetNameAvailable(
      XSSFWorkbook workbook, String newSheetName, int currentSheetIndex) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    requireSheetName(newSheetName, "newSheetName");

    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      if (sheetIndex == currentSheetIndex) {
        continue;
      }
      if (workbook.getSheetName(sheetIndex).equalsIgnoreCase(newSheetName)) {
        throw new IllegalArgumentException("Sheet already exists: " + newSheetName);
      }
    }
  }

  /** Fails when the requested target workbook index is out of bounds. */
  static void requireTargetIndex(XSSFWorkbook workbook, int targetIndex) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    int upperBound = workbook.getNumberOfSheets() - 1;
    if (targetIndex < 0 || targetIndex > upperBound) {
      throw new IllegalArgumentException(
          "targetIndex out of range: workbook has "
              + workbook.getNumberOfSheets()
              + " sheet(s), valid positions are 0 to "
              + upperBound
              + "; got "
              + targetIndex);
    }
  }

  /** Validates one workbook sheet name against the shared GridGrind contract. */
  static void requireSheetName(String value, String fieldName) {
    ExcelSheetNames.requireValid(value, fieldName);
  }
}
