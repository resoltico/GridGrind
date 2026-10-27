package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelAutofilterSortMethod;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.engine.autofilter.AutofilterAdapter;
import java.util.List;
import java.util.Optional;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSortCondition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STSortBy;

/** Residual coverage for autofilter OOXML translation and copied sort conditions. */
class ExcelAutofilterResidualCoverageTest {
  @Test
  void replaceSortStateRoundsTripsStrokeAndPinyinFontColorAndIconConditions() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      CTAutoFilter autoFilter = CTAutoFilter.Factory.newInstance();
      autoFilter.addNewSortState().setRef("A1:A2");

      AutofilterAdapter.replaceSortState(workbook, autoFilter, Optional.empty());
      assertFalse(autoFilter.isSetSortState());

      AutofilterSortStateSpec sortState =
          new AutofilterSortStateSpec(
              new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:C5"),
              false,
              false,
              Optional.of(ExcelAutofilterSortMethod.STROKE),
              List.of(
                  new AutofilterSortConditionSpec.FontColor(
                      new dev.erst.gridgrind.core.vocabulary.RangeRef("B2:B5"),
                      true,
                      ColorSpec.rgb(new RgbHex("#112233"))),
                  new AutofilterSortConditionSpec.Icon(
                      new dev.erst.gridgrind.core.vocabulary.RangeRef("C2:C5"), false, 3)));

      AutofilterAdapter.replaceSortState(workbook, autoFilter, Optional.of(sortState));

      AutofilterSortStateReport snapshot =
          AutofilterAdapter.sortState(workbook, autoFilter).orElseThrow();
      assertEquals(Optional.of(ExcelAutofilterSortMethod.STROKE), snapshot.sortMethod());
      assertEquals(
          new AutofilterSortConditionReport.FontColor(
              new dev.erst.gridgrind.core.vocabulary.RangeLabel("B2:B5"),
              true,
              CellColorReport.rgb(new RgbHex("#112233"))),
          snapshot.conditions().get(0));
      assertEquals(
          new AutofilterSortConditionReport.Icon(
              new dev.erst.gridgrind.core.vocabulary.RangeLabel("C2:C5"), false, 3),
          snapshot.conditions().get(1));

      AutofilterAdapter.replaceSortState(
          workbook,
          autoFilter,
          Optional.of(
              new AutofilterSortStateSpec(
                  new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:B3"),
                  false,
                  false,
                  Optional.of(ExcelAutofilterSortMethod.PINYIN),
                  List.of(
                      new AutofilterSortConditionSpec.Value(
                          new dev.erst.gridgrind.core.vocabulary.RangeRef("A2:A3"), false)))));
      assertEquals(
          Optional.of(ExcelAutofilterSortMethod.PINYIN),
          AutofilterAdapter.sortState(workbook, autoFilter).orElseThrow().sortMethod());
    }
  }

  @Test
  void sortStateAndSortConditionRejectMissingRefsAndColors() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      CTAutoFilter autoFilter = CTAutoFilter.Factory.newInstance();
      autoFilter.addNewSortState();

      IllegalArgumentException missingSortStateRef =
          assertThrows(
              IllegalArgumentException.class,
              () -> AutofilterAdapter.sortState(workbook, autoFilter));
      assertEquals("autofilter sort state is missing ref", missingSortStateRef.getMessage());

      CTSortCondition missingSortConditionRef = CTSortCondition.Factory.newInstance();
      IllegalArgumentException missingSortConditionRange =
          assertThrows(
              IllegalArgumentException.class,
              () -> AutofilterAdapter.sortConditionSnapshot(workbook, missingSortConditionRef));
      assertEquals(
          "autofilter sort condition is missing ref", missingSortConditionRange.getMessage());

      CTSortCondition cellColorSort = CTSortCondition.Factory.newInstance();
      cellColorSort.setRef("A1:A3");
      cellColorSort.setSortBy(STSortBy.CELL_COLOR);
      cellColorSort.setDxfId(0L);

      IllegalArgumentException missingCellColor =
          assertThrows(
              IllegalArgumentException.class,
              () -> AutofilterAdapter.sortConditionSnapshot(workbook, cellColorSort));
      assertEquals(
          "autofilter cell-color sort condition is missing dxf color",
          missingCellColor.getMessage());

      CTSortCondition fontColorSort = CTSortCondition.Factory.newInstance();
      fontColorSort.setRef("B1:B3");
      fontColorSort.setSortBy(STSortBy.FONT_COLOR);
      fontColorSort.setDxfId(0L);

      IllegalArgumentException missingFontColor =
          assertThrows(
              IllegalArgumentException.class,
              () -> AutofilterAdapter.sortConditionSnapshot(workbook, fontColorSort));
      assertEquals(
          "autofilter font-color sort condition is missing dxf color",
          missingFontColor.getMessage());

      CTSortCondition explicitValueSort = CTSortCondition.Factory.newInstance();
      explicitValueSort.setRef("C1:C3");
      explicitValueSort.setSortBy(STSortBy.VALUE);
      assertEquals(
          new AutofilterSortConditionReport.Value(
              new dev.erst.gridgrind.core.vocabulary.RangeLabel("C1:C3"), false),
          AutofilterAdapter.sortConditionSnapshot(workbook, explicitValueSort));
    }
  }

  @Test
  void sheetCopySupportCopiesFontColorAndIconSortConditions() {
    AutofilterSortConditionSpec.FontColor fontColor =
        assertInstanceOf(
            AutofilterSortConditionSpec.FontColor.class,
            SheetCopyAdapter.copyableSortCondition(
                new AutofilterSortConditionReport.FontColor(
                    new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:A2"),
                    true,
                    CellColorReport.rgb(new RgbHex("#AABBCC")))));
    assertEquals("A1:A2", fontColor.range().value());
    assertTrue(fontColor.descending());
    assertEquals(ColorSpec.rgb(new RgbHex("#AABBCC")), fontColor.color());

    AutofilterSortConditionSpec.Value value =
        assertInstanceOf(
            AutofilterSortConditionSpec.Value.class,
            SheetCopyAdapter.copyableSortCondition(
                new AutofilterSortConditionReport.Value(
                    new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:A2"), false)));
    assertEquals("A1:A2", value.range().value());
    assertFalse(value.descending());

    AutofilterSortConditionSpec.Icon icon =
        assertInstanceOf(
            AutofilterSortConditionSpec.Icon.class,
            SheetCopyAdapter.copyableSortCondition(
                new AutofilterSortConditionReport.Icon(
                    new dev.erst.gridgrind.core.vocabulary.RangeLabel("B1:B2"), false, 4)));
    assertEquals("B1:B2", icon.range().value());
    assertFalse(icon.descending());
    assertEquals(4, icon.iconId());
  }
}
