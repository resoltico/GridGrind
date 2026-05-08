package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.excel.foundation.ExcelAutofilterSortMethod;
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

      ExcelAutofilterOoxmlSupport.replaceSortState(workbook, autoFilter, Optional.empty());
      assertFalse(autoFilter.isSetSortState());

      ExcelAutofilterSortState sortState =
          new ExcelAutofilterSortState(
              "A1:C5",
              false,
              false,
              Optional.of(ExcelAutofilterSortMethod.STROKE),
              List.of(
                  new ExcelAutofilterSortCondition.FontColor(
                      "B2:B5", true, ExcelColor.rgb("#112233")),
                  new ExcelAutofilterSortCondition.Icon("C2:C5", false, 3)));

      ExcelAutofilterOoxmlSupport.replaceSortState(workbook, autoFilter, Optional.of(sortState));

      ExcelAutofilterSortStateSnapshot snapshot =
          ExcelAutofilterOoxmlSupport.sortState(workbook, autoFilter).orElseThrow();
      assertEquals(Optional.of(ExcelAutofilterSortMethod.STROKE), snapshot.sortMethod());
      assertEquals(
          new ExcelAutofilterSortConditionSnapshot.FontColor(
              "B2:B5", true, ExcelColorSnapshot.rgb("#112233")),
          snapshot.conditions().get(0));
      assertEquals(
          new ExcelAutofilterSortConditionSnapshot.Icon("C2:C5", false, 3),
          snapshot.conditions().get(1));

      ExcelAutofilterOoxmlSupport.replaceSortState(
          workbook,
          autoFilter,
          Optional.of(
              new ExcelAutofilterSortState(
                  "A1:B3",
                  false,
                  false,
                  Optional.of(ExcelAutofilterSortMethod.PINYIN),
                  List.of(new ExcelAutofilterSortCondition.Value("A2:A3", false)))));
      assertEquals(
          Optional.of(ExcelAutofilterSortMethod.PINYIN),
          ExcelAutofilterOoxmlSupport.sortState(workbook, autoFilter).orElseThrow().sortMethod());
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
              () -> ExcelAutofilterOoxmlSupport.sortState(workbook, autoFilter));
      assertEquals("autofilter sort state is missing ref", missingSortStateRef.getMessage());

      CTSortCondition missingSortConditionRef = CTSortCondition.Factory.newInstance();
      IllegalArgumentException missingSortConditionRange =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  ExcelAutofilterOoxmlSupport.sortConditionSnapshot(
                      workbook, missingSortConditionRef));
      assertEquals(
          "autofilter sort condition is missing ref", missingSortConditionRange.getMessage());

      CTSortCondition cellColorSort = CTSortCondition.Factory.newInstance();
      cellColorSort.setRef("A1:A3");
      cellColorSort.setSortBy(STSortBy.CELL_COLOR);
      cellColorSort.setDxfId(0L);

      IllegalArgumentException missingCellColor =
          assertThrows(
              IllegalArgumentException.class,
              () -> ExcelAutofilterOoxmlSupport.sortConditionSnapshot(workbook, cellColorSort));
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
              () -> ExcelAutofilterOoxmlSupport.sortConditionSnapshot(workbook, fontColorSort));
      assertEquals(
          "autofilter font-color sort condition is missing dxf color",
          missingFontColor.getMessage());

      CTSortCondition explicitValueSort = CTSortCondition.Factory.newInstance();
      explicitValueSort.setRef("C1:C3");
      explicitValueSort.setSortBy(STSortBy.VALUE);
      assertEquals(
          new ExcelAutofilterSortConditionSnapshot.Value("C1:C3", false),
          ExcelAutofilterOoxmlSupport.sortConditionSnapshot(workbook, explicitValueSort));
    }
  }

  @Test
  void sheetCopySupportCopiesFontColorAndIconSortConditions() {
    ExcelAutofilterSortCondition.FontColor fontColor =
        assertInstanceOf(
            ExcelAutofilterSortCondition.FontColor.class,
            ExcelSheetCopySupport.copyableSortCondition(
                new ExcelAutofilterSortConditionSnapshot.FontColor(
                    "A1:A2", true, ExcelColorSnapshot.rgb("#AABBCC"))));
    assertEquals("A1:A2", fontColor.range());
    assertTrue(fontColor.descending());
    assertEquals(ExcelColor.rgb("#AABBCC"), fontColor.color());

    ExcelAutofilterSortCondition.Value value =
        assertInstanceOf(
            ExcelAutofilterSortCondition.Value.class,
            ExcelSheetCopySupport.copyableSortCondition(
                new ExcelAutofilterSortConditionSnapshot.Value("A1:A2", false)));
    assertEquals("A1:A2", value.range());
    assertFalse(value.descending());

    ExcelAutofilterSortCondition.Icon icon =
        assertInstanceOf(
            ExcelAutofilterSortCondition.Icon.class,
            ExcelSheetCopySupport.copyableSortCondition(
                new ExcelAutofilterSortConditionSnapshot.Icon("B1:B2", false, 4)));
    assertEquals("B1:B2", icon.range());
    assertFalse(icon.descending());
    assertEquals(4, icon.iconId());
  }
}
