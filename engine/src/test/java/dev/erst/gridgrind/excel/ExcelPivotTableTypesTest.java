package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotTableNaming;
import dev.erst.gridgrind.excel.pivot.ExcelPivotDataPoiBridge;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableSelection;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableSnapshot;
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
    ExcelPivotTableSelection.ByNames selection = new ExcelPivotTableSelection.ByNames(mutableNames);
    mutableNames.clear();
    assertEquals(List.of("Sales Pivot 2026", "Ops Pivot"), selection.names());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPivotTableSelection.ByNames(List.of("Sales Pivot 2026", "sales pivot 2026")));

    ExcelPivotTableSnapshot.Supported supported =
        new ExcelPivotTableSnapshot.Supported(
            "Sales Pivot 2026",
            "Report",
            new ExcelPivotTableSnapshot.Anchor("C5", "C5:F9"),
            new ExcelPivotTableSnapshot.Source.Range("Data", "A1:D5"),
            List.of(new ExcelPivotTableSnapshot.Field(0, "Region")),
            List.of(new ExcelPivotTableSnapshot.Field(1, "Stage")),
            List.of(new ExcelPivotTableSnapshot.Field(2, "Owner")),
            List.of(
                new ExcelPivotTableSnapshot.DataField(
                    3,
                    "Amount",
                    ExcelPivotDataConsolidateFunction.SUM,
                    "Total Amount",
                    Optional.of("#,##0.00"))),
            true);
    ExcelPivotTableSnapshot.Unsupported unsupported =
        new ExcelPivotTableSnapshot.Unsupported(
            "Broken Pivot",
            "Report",
            new ExcelPivotTableSnapshot.Anchor("A3", "A3:C6"),
            "Pivot cache source no longer resolves cleanly.");

    assertTrue(supported.valuesAxisOnColumns());
    assertEquals("Amount", supported.dataFields().getFirst().sourceColumnName());
    assertEquals("Pivot cache source no longer resolves cleanly.", unsupported.detail());
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelPivotTableSnapshot.Field(-1, "Region"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPivotTableSnapshot.DataField(
                0, "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total", Optional.of(" ")));
  }

  @Test
  void pivotDataConsolidateFunctionsRoundTripSupportedPoiValues() {
    for (ExcelPivotDataConsolidateFunction function : ExcelPivotDataConsolidateFunction.values()) {
      assertEquals(
          function,
          ExcelPivotDataPoiBridge.fromPoi(ExcelPivotDataPoiBridge.toPoi(function)),
          () -> "pivot function must round-trip through the POI bridge: " + function);
    }
    assertThrows(IllegalArgumentException.class, () -> ExcelPivotDataPoiBridge.fromPoi(null));
    assertThrows(IllegalArgumentException.class, () -> ExcelPivotDataPoiBridge.toPoi(null));
  }

  @Test
  void pivotTypesRejectAdditionalInvalidInputs() {
    assertThrows(NullPointerException.class, () -> new ExcelPivotTableSelection.ByNames(null));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelPivotTableSelection.ByNames(List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPivotTableSelection.ByNames(
                new ArrayList<>(java.util.Arrays.asList("Sales Pivot 2026", null))));

    ExcelPivotTableSnapshot.Anchor anchor = new ExcelPivotTableSnapshot.Anchor("C5", "C5:F9");
    ExcelPivotTableSnapshot.Source.Range source =
        new ExcelPivotTableSnapshot.Source.Range("Data", "A1:D5");
    ExcelPivotTableSnapshot.Field rowField = new ExcelPivotTableSnapshot.Field(0, "Region");
    ExcelPivotTableSnapshot.DataField dataField =
        new ExcelPivotTableSnapshot.DataField(
            3, "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total Amount", Optional.empty());

    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPivotTableSnapshot.Supported(
                "Sales Pivot 2026",
                "Report",
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
            new ExcelPivotTableSnapshot.Supported(
                "Sales Pivot 2026",
                "Report",
                anchor,
                source,
                List.of(rowField),
                List.of(),
                List.of(),
                new ArrayList<>(java.util.Arrays.asList((ExcelPivotTableSnapshot.DataField) null)),
                false));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPivotTableSnapshot.DataField(
                3, "Amount", ExcelPivotDataConsolidateFunction.SUM, null, Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPivotTableSnapshot.DataField(
                3, "Amount", ExcelPivotDataConsolidateFunction.SUM, " ", Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelPivotTableSnapshot.Unsupported("Broken Pivot", "Report", anchor, " "));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPivotTableSnapshot.DataField(
                3, null, ExcelPivotDataConsolidateFunction.SUM, "Total Amount", Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPivotTableSnapshot.DataField(
                3, " ", ExcelPivotDataConsolidateFunction.SUM, "Total Amount", Optional.empty()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPivotTableSnapshot.DataField(
                3, "Amount", null, "Total Amount", Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPivotTableSnapshot.DataField(
                -1,
                "Amount",
                ExcelPivotDataConsolidateFunction.SUM,
                "Total Amount",
                Optional.empty()));
  }
}
