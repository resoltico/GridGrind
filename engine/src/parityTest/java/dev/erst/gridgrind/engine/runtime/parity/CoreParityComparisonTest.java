package dev.erst.gridgrind.engine.runtime.parity;

import static dev.erst.gridgrind.engine.runtime.parity.XlsxParityCoreProbeGroup.matchesCoreLayout;
import static dev.erst.gridgrind.engine.runtime.parity.XlsxParityCoreProbeGroup.matchesCorePrint;
import static dev.erst.gridgrind.engine.runtime.parity.XlsxParityCoreProbeGroup.matchesCoreProtection;
import static dev.erst.gridgrind.engine.runtime.parity.XlsxParityCoreProbeGroup.matchesCoreStyle;
import static dev.erst.gridgrind.engine.runtime.parity.XlsxParityProbeRegistry.matchesColorDescriptor;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.layout.PaneReport;
import dev.erst.gridgrind.core.model.layout.SheetLayoutReport;
import dev.erst.gridgrind.core.model.print.PrintLayoutReport;
import dev.erst.gridgrind.core.model.print.PrintScalingReport;
import dev.erst.gridgrind.core.model.protection.SheetProtectionReport;
import dev.erst.gridgrind.core.model.protection.SheetProtectionSettings;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Counterexamples ensure the parity comparators detect observed-field drift. */
class CoreParityComparisonTest {
  @Test
  void independentlyChangedStyleLayoutPrintAndProtectionFactsAreRejected(@TempDir Path root) {
    var observation =
        XlsxParityCoreProbeGroup.readCoreObservation(
            new XlsxParityProbeRegistry.ProbeContext(root));
    var direct = observation.direct();
    var style =
        observation.cells().cells().stream()
            .filter(cell -> "B3".equals(cell.address().value()))
            .findFirst()
            .orElseThrow()
            .style()
            .orElseThrow();
    var expectedStyle = direct.styles().get("B3");
    assertTrue(matchesCoreStyle(style, expectedStyle));
    var changedStyle =
        new CellStyleReport(
            "__parity_number_format_drift__",
            style.alignment(),
            style.font(),
            style.fill(),
            style.border(),
            style.protection());
    assertFalse(matchesCoreStyle(changedStyle, expectedStyle));

    var layout = observation.layout().layout();
    assertTrue(matchesCoreLayout(layout, direct));
    var pane = (PaneReport.Frozen) layout.pane();
    var changedPane =
        new PaneReport.Frozen(
            pane.splitColumn(), pane.splitRow(), pane.leftmostColumn(), pane.topRow() + 1);
    assertFalse(
        matchesCoreLayout(
            new SheetLayoutReport(
                layout.sheetName(),
                changedPane,
                layout.zoomPercent(),
                layout.presentation(),
                layout.columns(),
                layout.rows()),
            direct));
    var changedColumns =
        layout.columns().stream()
            .map(
                column ->
                    column.columnIndex() != 0
                        ? column
                        : new dev.erst.gridgrind.core.model.layout.ColumnLayoutReport(
                            column.columnIndex(),
                            column.widthCharacters() + 1,
                            column.hidden(),
                            column.outlineLevel(),
                            column.collapsed()))
            .toList();
    assertFalse(
        matchesCoreLayout(
            new SheetLayoutReport(
                layout.sheetName(),
                layout.pane(),
                layout.zoomPercent(),
                layout.presentation(),
                changedColumns,
                layout.rows()),
            direct));
    var changedRows =
        layout.rows().stream()
            .map(
                row ->
                    row.rowIndex() != 0
                        ? row
                        : new dev.erst.gridgrind.core.model.layout.RowLayoutReport(
                            row.rowIndex(),
                            row.heightPoints() + 1,
                            row.hidden(),
                            row.outlineLevel(),
                            row.collapsed()))
            .toList();
    assertFalse(
        matchesCoreLayout(
            new SheetLayoutReport(
                layout.sheetName(),
                layout.pane(),
                layout.zoomPercent(),
                layout.presentation(),
                layout.columns(),
                changedRows),
            direct));

    var print = observation.print().layout();
    assertTrue(matchesCorePrint(print, direct.print()));
    var fit = (PrintScalingReport.Fit) print.scaling();
    assertFalse(
        matchesCorePrint(
            new PrintLayoutReport(
                print.sheetName(),
                print.printArea(),
                print.orientation(),
                new PrintScalingReport.Fit(fit.widthPages() + 1, fit.heightPages()),
                print.repeatingRows(),
                print.repeatingColumns(),
                print.header(),
                print.footer(),
                print.setup()),
            direct.print()));
    var page = direct.print();
    assertFalse(
        matchesCorePrint(
            print,
            new XlsxParityOracle.CorePrintSnapshot(
                page.landscape(),
                page.fitToPage(),
                page.widthPages(),
                page.heightPages(),
                page.firstTitleRow(),
                page.lastTitleRow(),
                page.firstTitleColumn(),
                page.lastTitleColumn(),
                !page.horizontallyCentered(),
                page.verticallyCentered())));

    var protection = observation.opsSummary().sheet().protection();
    assertTrue(matchesCoreProtection(protection, direct.sheetProtection()));
    var settings = ((SheetProtectionReport.Protected) protection).settings();
    var changedSettings =
        new SheetProtectionSettings(
            !settings.autoFilterLocked(),
            settings.deleteColumnsLocked(),
            settings.deleteRowsLocked(),
            settings.formatCellsLocked(),
            settings.formatColumnsLocked(),
            settings.formatRowsLocked(),
            settings.insertColumnsLocked(),
            settings.insertHyperlinksLocked(),
            settings.insertRowsLocked(),
            settings.objectsLocked(),
            settings.pivotTablesLocked(),
            settings.scenariosLocked(),
            settings.selectLockedCellsLocked(),
            settings.selectUnlockedCellsLocked(),
            settings.sortLocked());
    assertFalse(
        matchesCoreProtection(
            new SheetProtectionReport.Protected(changedSettings), direct.sheetProtection()));
  }

  @Test
  void colorDescriptorsCompareRgbBytesAndRejectUnknownOrMissingColorClaims() {
    var rgb = CellColorReport.rgb(new RgbHex("#D9EAF7"));
    assertTrue(matchesColorDescriptor(rgb, "rgb=d9eaf7"));
    assertTrue(matchesColorDescriptor(rgb, "rgb=D9EAF7"));
    assertFalse(matchesColorDescriptor(rgb, "rgb=D9EAF8"));
    assertTrue(matchesColorDescriptor(null, "none"));
    assertFalse(matchesColorDescriptor(rgb, "none"));
    assertFalse(matchesColorDescriptor(null, "rgb=D9EAF7"));
    assertFalse(matchesColorDescriptor(rgb, "unknown"));
    assertFalse(matchesColorDescriptor(rgb, "rgb=D9EAF7|unknown"));
    assertFalse(matchesColorDescriptor(rgb, "rgb=D9EAF7|"));
    assertFalse(matchesColorDescriptor(CellColorReport.theme(4), "theme=not-a-number"));
  }
}
