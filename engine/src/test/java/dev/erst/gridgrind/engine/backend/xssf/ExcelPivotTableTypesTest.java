package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotTableNaming;
import dev.erst.gridgrind.engine.pivots.PivotAdapter;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for pivot-table value objects, validation helpers, and enum mappings. */
class ExcelPivotTableTypesTest {
  @Test
  void validatesNamingSelectionsDefinitionsAndSnapshots() {
    assertEquals("Sales Pivot 2026", ExcelPivotTableNaming.validateName("Sales Pivot 2026"));
    assertThrows(NullPointerException.class, () -> ExcelPivotTableNaming.validateName(null));
    assertThrows(IllegalArgumentException.class, () -> ExcelPivotTableNaming.validateName("   "));

    List<String> mutableNames = new ArrayList<>(List.of("Sales Pivot 2026", "Ops Pivot"));
    PivotTableSelector.ByNames selection = new PivotTableSelector.ByNames(mutableNames);
    mutableNames.clear();
    assertEquals(List.of("Sales Pivot 2026", "Ops Pivot"), selection.names());
    assertThrows(
        IllegalArgumentException.class,
        () -> new PivotTableSelector.ByNames(List.of("Sales Pivot 2026", "sales pivot 2026")));

    PivotTableReport.Supported supported =
        new PivotTableReport.Supported(
            "Sales Pivot 2026",
            new dev.erst.gridgrind.core.vocabulary.SheetName("Report"),
            new PivotTableReport.Anchor(
                new dev.erst.gridgrind.core.vocabulary.CellAddress("C5"),
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("C5:F9")),
            new PivotTableReport.PivotSourceReport.Range(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:D5")),
            List.of(new PivotTableReport.Field(0, "Region")),
            List.of(new PivotTableReport.Field(1, "Stage")),
            List.of(new PivotTableReport.Field(2, "Owner")),
            List.of(
                new PivotTableReport.DataField(
                    3,
                    "Amount",
                    ExcelPivotDataConsolidateFunction.SUM,
                    "Total Amount",
                    Optional.of("#,##0.00"))),
            true);
    PivotTableReport.Unsupported unsupported =
        new PivotTableReport.Unsupported(
            "Broken Pivot",
            new dev.erst.gridgrind.core.vocabulary.SheetName("Report"),
            new PivotTableReport.Anchor(
                new dev.erst.gridgrind.core.vocabulary.CellAddress("A3"),
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A3:C6")),
            "Pivot cache source no longer resolves cleanly.");

    assertTrue(supported.valuesAxisOnColumns());
    assertEquals("Amount", supported.dataFields().getFirst().sourceColumnName());
    assertEquals("Pivot cache source no longer resolves cleanly.", unsupported.detail());
    assertThrows(IllegalArgumentException.class, () -> new PivotTableReport.Field(-1, "Region"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableReport.DataField(
                0, "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total", Optional.of(" ")));
  }

  @Test
  void pivotDataConsolidateFunctionsRoundTripSupportedPoiValues() {
    for (ExcelPivotDataConsolidateFunction function : ExcelPivotDataConsolidateFunction.values()) {
      assertEquals(
          function,
          PivotAdapter.fromPoi(PivotAdapter.toPoi(function)),
          () -> "pivot function must round-trip through the POI bridge: " + function);
    }
    assertThrows(IllegalArgumentException.class, () -> PivotAdapter.fromPoi(null));
    assertThrows(IllegalArgumentException.class, () -> PivotAdapter.toPoi(null));
  }

  @Test
  void pivotTypesRejectAdditionalInvalidInputs() {
    assertThrows(NullPointerException.class, () -> new PivotTableSelector.ByNames(null));
    assertThrows(IllegalArgumentException.class, () -> new PivotTableSelector.ByNames(List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new PivotTableSelector.ByNames(
                new ArrayList<>(java.util.Arrays.asList("Sales Pivot 2026", null))));

    PivotTableReport.Anchor anchor =
        new PivotTableReport.Anchor(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("C5"),
            new dev.erst.gridgrind.core.vocabulary.RangeLabel("C5:F9"));
    PivotTableReport.PivotSourceReport.Range source =
        new PivotTableReport.PivotSourceReport.Range(
            new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
            new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:D5"));
    PivotTableReport.Field rowField = new PivotTableReport.Field(0, "Region");
    PivotTableReport.DataField dataField =
        new PivotTableReport.DataField(
            3, "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total Amount", Optional.empty());

    assertThrows(
        NullPointerException.class,
        () ->
            new PivotTableReport.Supported(
                "Sales Pivot 2026",
                new dev.erst.gridgrind.core.vocabulary.SheetName("Report"),
                anchor,
                source,
                null,
                List.of(),
                List.of(),
                List.of(dataField),
                false));
    assertThrows(
        NullPointerException.class,
        () ->
            new PivotTableReport.Supported(
                "Sales Pivot 2026",
                new dev.erst.gridgrind.core.vocabulary.SheetName("Report"),
                anchor,
                source,
                List.of(rowField),
                List.of(),
                List.of(),
                new ArrayList<>(java.util.Arrays.asList((PivotTableReport.DataField) null)),
                false));
    assertThrows(
        NullPointerException.class,
        () ->
            new PivotTableReport.DataField(
                3, "Amount", ExcelPivotDataConsolidateFunction.SUM, null, Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableReport.DataField(
                3, "Amount", ExcelPivotDataConsolidateFunction.SUM, " ", Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableReport.Unsupported(
                "Broken Pivot",
                new dev.erst.gridgrind.core.vocabulary.SheetName("Report"),
                anchor,
                " "));
    assertThrows(
        NullPointerException.class,
        () ->
            new PivotTableReport.DataField(
                3, null, ExcelPivotDataConsolidateFunction.SUM, "Total Amount", Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableReport.DataField(
                3, " ", ExcelPivotDataConsolidateFunction.SUM, "Total Amount", Optional.empty()));
    assertThrows(
        NullPointerException.class,
        () -> new PivotTableReport.DataField(3, "Amount", null, "Total Amount", Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableReport.DataField(
                -1,
                "Amount",
                ExcelPivotDataConsolidateFunction.SUM,
                "Total Amount",
                Optional.empty()));
  }
}
