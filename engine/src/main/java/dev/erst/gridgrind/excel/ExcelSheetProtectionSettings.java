package dev.erst.gridgrind.excel;

/** Supported sheet-protection lock flags that GridGrind can author and report honestly. */
public record ExcelSheetProtectionSettings(
    boolean autoFilterLocked,
    boolean deleteColumnsLocked,
    boolean deleteRowsLocked,
    boolean formatCellsLocked,
    boolean formatColumnsLocked,
    boolean formatRowsLocked,
    boolean insertColumnsLocked,
    boolean insertHyperlinksLocked,
    boolean insertRowsLocked,
    boolean objectsLocked,
    boolean pivotTablesLocked,
    boolean scenariosLocked,
    boolean selectLockedCellsLocked,
    boolean selectUnlockedCellsLocked,
    boolean sortLocked) {}
