package dev.erst.gridgrind.core.model.pivot;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Authored pivot constraints belong to the shared model. */
class PivotTableSpecTest {
  @Test
  void validatesAuthoredDefinition() {
    PivotTableSpec definition =
        new PivotTableSpec(
            "Sales Pivot 2026",
            new dev.erst.gridgrind.core.vocabulary.SheetName("Report"),
            new PivotTableSpec.PivotSourceSpec.Table(
                new dev.erst.gridgrind.core.vocabulary.TableName("SalesTable")),
            new PivotTableSpec.Anchor(new dev.erst.gridgrind.core.vocabulary.CellAddress("C5")),
            List.of("Region"),
            List.of("Stage"),
            List.of("Owner"),
            List.of(
                new PivotTableSpec.DataField(
                    "Amount",
                    ExcelPivotDataConsolidateFunction.SUM,
                    "Amount",
                    Optional.of("#,##0.00"))));

    assertEquals("Sales Pivot 2026", definition.name());
    assertEquals("Amount", definition.dataFields().getFirst().displayName());
    assertEquals(Optional.of("#,##0.00"), definition.dataFields().getFirst().valueFormat());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableSpec.DataField(
                "Amount", ExcelPivotDataConsolidateFunction.SUM, " ", Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableSpec.PivotSourceSpec.NamedRange(
                new dev.erst.gridgrind.core.vocabulary.DefinedName("Sales Pivot 2026")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableSpec.PivotSourceSpec.Table(
                new dev.erst.gridgrind.core.vocabulary.TableName("Sales Table With Spaces")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableSpec(
                "Sales Pivot 2026",
                new dev.erst.gridgrind.core.vocabulary.SheetName("Report"),
                new PivotTableSpec.PivotSourceSpec.Range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5")),
                new PivotTableSpec.Anchor(new dev.erst.gridgrind.core.vocabulary.CellAddress("A1")),
                List.of("Region"),
                List.of("Stage"),
                List.of("Owner"),
                List.of(
                    new PivotTableSpec.DataField(
                        "Region",
                        ExcelPivotDataConsolidateFunction.SUM,
                        "Total",
                        Optional.empty()))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableSpec.DataField(
                "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total", Optional.of(" ")));
  }

  @Test
  void rejectsInvalidAuthoredFieldsAndAxisAssignments() {
    assertThrows(
        NullPointerException.class,
        () -> new PivotTableSpec.Anchor(new dev.erst.gridgrind.core.vocabulary.CellAddress(null)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PivotTableSpec.Anchor(new dev.erst.gridgrind.core.vocabulary.CellAddress(" ")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableSpec.PivotSourceSpec.Range(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                new dev.erst.gridgrind.core.vocabulary.RangeRef(" ")));
    assertThrows(
        NullPointerException.class,
        () ->
            new PivotTableSpec.DataField(
                null, ExcelPivotDataConsolidateFunction.SUM, "Total", Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableSpec.DataField(
                " ", ExcelPivotDataConsolidateFunction.SUM, "Total", Optional.empty()));
    assertThrows(
        NullPointerException.class,
        () -> new PivotTableSpec.DataField("Amount", null, "Total", Optional.empty()));
    assertThrows(
        NullPointerException.class,
        () ->
            new PivotTableSpec(
                "Missing Data",
                new dev.erst.gridgrind.core.vocabulary.SheetName("Report"),
                new PivotTableSpec.PivotSourceSpec.Range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5")),
                new PivotTableSpec.Anchor(new dev.erst.gridgrind.core.vocabulary.CellAddress("A3")),
                List.of(),
                List.of(),
                List.of(),
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableSpec(
                "Empty Data",
                new dev.erst.gridgrind.core.vocabulary.SheetName("Report"),
                new PivotTableSpec.PivotSourceSpec.Range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5")),
                new PivotTableSpec.Anchor(new dev.erst.gridgrind.core.vocabulary.CellAddress("A3")),
                List.of(),
                List.of(),
                List.of(),
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new PivotTableSpec(
                "Null Data Field",
                new dev.erst.gridgrind.core.vocabulary.SheetName("Report"),
                new PivotTableSpec.PivotSourceSpec.Range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5")),
                new PivotTableSpec.Anchor(new dev.erst.gridgrind.core.vocabulary.CellAddress("A3")),
                List.of(),
                List.of(),
                List.of(),
                new ArrayList<>(java.util.Arrays.asList((PivotTableSpec.DataField) null))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableSpec(
                "Axis Clash",
                new dev.erst.gridgrind.core.vocabulary.SheetName("Report"),
                new PivotTableSpec.PivotSourceSpec.Range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5")),
                new PivotTableSpec.Anchor(new dev.erst.gridgrind.core.vocabulary.CellAddress("A3")),
                List.of("Owner"),
                List.of("Region"),
                List.of("region"),
                List.of(
                    new PivotTableSpec.DataField(
                        "Amount",
                        ExcelPivotDataConsolidateFunction.SUM,
                        "Total",
                        Optional.empty()))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableSpec(
                "Duplicate Axis Names",
                new dev.erst.gridgrind.core.vocabulary.SheetName("Report"),
                new PivotTableSpec.PivotSourceSpec.Range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5")),
                new PivotTableSpec.Anchor(new dev.erst.gridgrind.core.vocabulary.CellAddress("A3")),
                List.of(),
                List.of("Region", "region"),
                List.of(),
                List.of(
                    new PivotTableSpec.DataField(
                        "Amount",
                        ExcelPivotDataConsolidateFunction.SUM,
                        "Total",
                        Optional.empty()))));
    assertThrows(
        NullPointerException.class,
        () ->
            new PivotTableSpec(
                "Null Axis Name",
                new dev.erst.gridgrind.core.vocabulary.SheetName("Report"),
                new PivotTableSpec.PivotSourceSpec.Range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5")),
                new PivotTableSpec.Anchor(new dev.erst.gridgrind.core.vocabulary.CellAddress("A3")),
                new ArrayList<>(java.util.Arrays.asList((String) null)),
                List.of(),
                List.of(),
                List.of(
                    new PivotTableSpec.DataField(
                        "Amount",
                        ExcelPivotDataConsolidateFunction.SUM,
                        "Total",
                        Optional.empty()))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableSpec(
                "Blank Axis",
                new dev.erst.gridgrind.core.vocabulary.SheetName("Report"),
                new PivotTableSpec.PivotSourceSpec.Range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5")),
                new PivotTableSpec.Anchor(new dev.erst.gridgrind.core.vocabulary.CellAddress("A3")),
                List.of(" "),
                List.of(),
                List.of(),
                List.of(
                    new PivotTableSpec.DataField(
                        "Amount",
                        ExcelPivotDataConsolidateFunction.SUM,
                        "Total",
                        Optional.empty()))));
  }
}
