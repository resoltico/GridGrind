package dev.erst.gridgrind.core.model.protection;

import dev.erst.gridgrind.core.metadata.CatalogType;

/** Supported sheet-protection lock flags that GridGrind can author and report honestly. */
@CatalogType(
    id = "SheetProtectionSettings",
    summary = "Supported sheet-protection lock flags authored and reported by GridGrind.",
    displayOrder = 42)
public record SheetProtectionSettings(
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
