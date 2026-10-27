package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.core.result.SheetSummaryReport;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetNames;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import dev.erst.gridgrind.core.vocabulary.SheetCopyPosition;
import dev.erst.gridgrind.engine.protection.ProtectionAdapter;
import dev.erst.gridgrind.engine.tables.TableAdapter;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Owns sheet lookup, lifecycle, selection and structural summaries for one workbook. */
@SuppressWarnings("PMD.CommentRequired")
public final class SheetCollectionAdapter {
  private final XssfWorkbook workbook;

  SheetCollectionAdapter(XssfWorkbook workbook) {
    this.workbook = Objects.requireNonNull(workbook, "workbook must not be null");
  }

  /** Renames an existing sheet to a new destination name. */
  public SheetCollectionAdapter renameSheet(String sheetName, String newSheetName) {
    SheetCollectionAdapter.requireSheetName(sheetName, "sheetName");
    SheetCollectionAdapter.requireSheetName(newSheetName, "newSheetName");

    int sheetIndex = SheetCollectionAdapter.requiredSheetIndex(workbook.xssfWorkbook(), sheetName);
    SheetCollectionAdapter.requireSheetNameAvailable(
        workbook.xssfWorkbook(), newSheetName, sheetIndex);
    if (sheetName.equals(newSheetName)) {
      return this;
    }
    TableAdapter.canonicalizeWorkbook(workbook.xssfWorkbook());
    workbook.xssfWorkbook().setSheetName(sheetIndex, newSheetName);
    return this;
  }

  /**
   * Deletes an existing sheet while preserving a valid visible active and selected workbook state.
   */
  public SheetCollectionAdapter deleteSheet(String sheetName) {
    SheetCollectionAdapter.requireSheetName(sheetName, "sheetName");

    int sheetIndex = SheetCollectionAdapter.requiredSheetIndex(workbook.xssfWorkbook(), sheetName);
    if (workbook.xssfWorkbook().getNumberOfSheets() == 1) {
      throw new IllegalArgumentException(
          "cannot delete sheet '" + sheetName + "': a workbook must contain at least one sheet");
    }
    SheetCollectionAdapter.requireNotLastVisibleSheet(
        workbook.xssfWorkbook(),
        sheetIndex,
        "cannot delete the last visible sheet '" + sheetName + "'");
    workbook.xssfWorkbook().removeSheetAt(sheetIndex);
    SheetCollectionAdapter.normalizeWorkbookViewState(workbook.xssfWorkbook());
    return this;
  }

  /** Moves an existing sheet to a zero-based workbook position. */
  public SheetCollectionAdapter moveSheet(String sheetName, int targetIndex) {
    SheetCollectionAdapter.requireSheetName(sheetName, "sheetName");

    SheetCollectionAdapter.requiredSheetIndex(workbook.xssfWorkbook(), sheetName);
    SheetCollectionAdapter.requireTargetIndex(workbook.xssfWorkbook(), targetIndex);
    workbook.xssfWorkbook().setSheetOrder(sheetName, targetIndex);
    SheetCollectionAdapter.normalizeWorkbookViewState(workbook.xssfWorkbook());
    return this;
  }

  /** Sets the active sheet and ensures it is also selected. */
  public SheetCollectionAdapter setActiveSheet(String sheetName) {
    SheetCollectionAdapter.requireSheetName(sheetName, "sheetName");

    int sheetIndex = SheetCollectionAdapter.requiredSheetIndex(workbook.xssfWorkbook(), sheetName);
    SheetCollectionAdapter.requireVisibleSheet(
        workbook.xssfWorkbook(), sheetIndex, "cannot activate hidden sheet '" + sheetName + "'");

    workbook.xssfWorkbook().setActiveSheet(sheetIndex);
    workbook.xssfWorkbook().getSheetAt(sheetIndex).setSelected(true);
    SheetCollectionAdapter.normalizeWorkbookViewState(workbook.xssfWorkbook());
    return this;
  }

  /**
   * Sets the selected visible sheet set and normalizes the active tab into that selection if
   * needed.
   */
  public SheetCollectionAdapter setSelectedSheets(List<String> sheetNames) {
    List<String> normalizedSheetNames = SheetCollectionAdapter.copySelectedSheetNames(sheetNames);

    Set<String> selectedNames = new LinkedHashSet<>(normalizedSheetNames);
    for (String sheetName : selectedNames) {
      int sheetIndex =
          SheetCollectionAdapter.requiredSheetIndex(workbook.xssfWorkbook(), sheetName);
      SheetCollectionAdapter.requireVisibleSheet(
          workbook.xssfWorkbook(), sheetIndex, "cannot select hidden sheet '" + sheetName + "'");
    }

    for (int sheetIndex = 0;
        sheetIndex < workbook.xssfWorkbook().getNumberOfSheets();
        sheetIndex++) {
      XSSFSheet sheet = workbook.xssfWorkbook().getSheetAt(sheetIndex);
      sheet.setSelected(selectedNames.contains(sheet.getSheetName()));
    }

    int activeSheetIndex = workbook.xssfWorkbook().getActiveSheetIndex();
    if (activeSheetIndex < 0
        || activeSheetIndex >= workbook.xssfWorkbook().getNumberOfSheets()
        || !selectedNames.contains(workbook.xssfWorkbook().getSheetName(activeSheetIndex))) {
      workbook
          .xssfWorkbook()
          .setActiveSheet(
              SheetCollectionAdapter.requiredSelectedSheetIndex(
                  workbook.xssfWorkbook(), normalizedSheetNames.getFirst()));
    }

    SheetCollectionAdapter.normalizeWorkbookViewState(workbook.xssfWorkbook());
    return this;
  }

  /** Sets one sheet visibility while preserving at least one visible active selected sheet. */
  public SheetCollectionAdapter setSheetVisibility(
      String sheetName, ExcelSheetVisibility visibility) {
    SheetCollectionAdapter.requireSheetName(sheetName, "sheetName");
    Objects.requireNonNull(visibility, "visibility must not be null");

    int sheetIndex = SheetCollectionAdapter.requiredSheetIndex(workbook.xssfWorkbook(), sheetName);
    ExcelSheetVisibility currentVisibility =
        SheetCollectionAdapter.fromPoiVisibility(
            workbook.xssfWorkbook().getSheetVisibility(sheetIndex));
    if (visibility != ExcelSheetVisibility.VISIBLE
        && currentVisibility == ExcelSheetVisibility.VISIBLE) {
      SheetCollectionAdapter.requireNotLastVisibleSheet(
          workbook.xssfWorkbook(),
          sheetIndex,
          "cannot hide the last visible sheet '" + sheetName + "'");
    }
    workbook
        .xssfWorkbook()
        .setSheetVisibility(sheetIndex, SheetCollectionAdapter.toPoiVisibility(visibility));
    SheetCollectionAdapter.normalizeWorkbookViewState(workbook.xssfWorkbook());
    return this;
  }

  /** Returns structural and state facts for one sheet. */
  public SheetSummaryReport summarizeSheet(String sheetName) {
    SheetCollectionAdapter.requireSheetName(sheetName, "sheetName");

    XSSFSheet sheet = SheetCollectionAdapter.requiredSheet(workbook.xssfWorkbook(), sheetName);
    XssfSheet excelSheet = workbook.sheet(sheetName);
    return new SheetSummaryReport(
        new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheetName),
        SheetCollectionAdapter.fromPoiVisibility(
            workbook
                .xssfWorkbook()
                .getSheetVisibility(workbook.xssfWorkbook().getSheetIndex(sheet))),
        ProtectionAdapter.snapshotSheet(sheet),
        excelSheet.rows().physicalCount(),
        excelSheet.rows().lastIndex(),
        excelSheet.columns().lastIndex());
  }

  public XssfSheet getOrCreateSheet(String sheetName) {
    SheetCollectionAdapter.requireSheetName(sheetName, "sheetName");

    Sheet sheet = workbook.xssfWorkbook().getSheet(sheetName);
    if (sheet == null) {
      sheet = workbook.xssfWorkbook().createSheet(sheetName);
    }
    return new XssfSheet(
        sheet, workbook.context().styleRegistry(), workbook.context().formulaRuntime());
  }

  public XssfSheet sheet(String sheetName) {
    return new XssfSheet(
        requiredSheet(workbook.xssfWorkbook(), sheetName),
        workbook.context().styleRegistry(),
        workbook.context().formulaRuntime());
  }

  public java.util.List<String> sheetNames() {
    java.util.List<String> sheetNames =
        new ArrayList<>(workbook.xssfWorkbook().getNumberOfSheets());
    for (int sheetIndex = 0;
        sheetIndex < workbook.xssfWorkbook().getNumberOfSheets();
        sheetIndex++) {
      sheetNames.add(workbook.xssfWorkbook().getSheetName(sheetIndex));
    }
    return java.util.List.copyOf(sheetNames);
  }

  /** Copies one sheet into a new visible, unselected sheet at the requested workbook position. */
  public SheetCollectionAdapter copySheet(
      String sourceSheetName, String newSheetName, SheetCopyPosition position) {
    new SheetCopyAdapter().copySheet(workbook, sourceSheetName, newSheetName, position);
    return this;
  }

  /** Returns the number of sheets in the workbook. */
  public int sheetCount() {
    return workbook.xssfWorkbook().getNumberOfSheets();
  }

  /**
   * Returns the active visible-or-primary sheet name for a workbook that already contains sheets.
   */
  public static String activeSheetName(XSSFWorkbook workbook) {
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
  public static List<String> selectedSheetNames(XSSFWorkbook workbook) {
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
  public static void normalizeWorkbookViewState(XSSFWorkbook workbook) {
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
  public static int firstVisibleSheetIndex(XSSFWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      if (workbook.getSheetVisibility(sheetIndex) == SheetVisibility.VISIBLE) {
        return sheetIndex;
      }
    }
    throw new IllegalStateException("workbook must contain at least one visible sheet");
  }

  /** Returns the count of visible sheets in one workbook. */
  public static int visibleSheetCount(XSSFWorkbook workbook) {
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
  public static void requireNotLastVisibleSheet(
      XSSFWorkbook workbook, int sheetIndex, String message) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(message, "message must not be null");
    if (workbook.getSheetVisibility(sheetIndex) == SheetVisibility.VISIBLE
        && visibleSheetCount(workbook) == 1) {
      throw new IllegalArgumentException(message);
    }
  }

  /** Returns the requested primary selection sheet index after validating it exists. */
  public static int requiredSelectedSheetIndex(
      XSSFWorkbook workbook, String requestedPrimarySheetName) {
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
  public static void requireVisibleSheet(XSSFWorkbook workbook, int sheetIndex, String message) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(message, "message must not be null");
    if (workbook.getSheetVisibility(sheetIndex) != SheetVisibility.VISIBLE) {
      throw new IllegalArgumentException(message);
    }
  }

  /** Returns a validated copy of the requested selected sheet names. */
  public static List<String> copySelectedSheetNames(List<String> sheetNames) {
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
  public static XSSFSheet requiredSheet(XSSFWorkbook workbook, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    requireSheetName(sheetName, "sheetName");

    XSSFSheet sheet = workbook.getSheet(sheetName);
    if (sheet == null) {
      throw new SheetNotFoundException(sheetName);
    }
    return sheet;
  }

  /** Returns the zero-based index of an existing sheet. */
  public static int requiredSheetIndex(XSSFWorkbook workbook, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    requireSheetName(sheetName, "sheetName");

    int sheetIndex = workbook.getSheetIndex(sheetName);
    if (sheetIndex < 0) {
      throw new SheetNotFoundException(sheetName);
    }
    return sheetIndex;
  }

  /** Fails when another sheet already owns the requested new name. */
  public static void requireSheetNameAvailable(
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
  public static void requireTargetIndex(XSSFWorkbook workbook, int targetIndex) {
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
  public static void requireSheetName(String value, String fieldName) {
    ExcelSheetNames.requireValid(value, fieldName);
  }

  static ExcelSheetVisibility fromPoiVisibility(SheetVisibility visibility) {
    Objects.requireNonNull(visibility, "visibility must not be null");
    return ExcelSheetVisibility.valueOf(visibility.name());
  }

  static SheetVisibility toPoiVisibility(ExcelSheetVisibility visibility) {
    Objects.requireNonNull(visibility, "visibility must not be null");
    return SheetVisibility.valueOf(visibility.name());
  }
}
