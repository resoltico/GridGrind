package dev.erst.gridgrind.core.plan;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.formula.ResolvedRawFormula;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.ResolvedRichText;
import dev.erst.gridgrind.core.model.style.ResolvedRichTextRun;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for CellValueSpec record construction and ResolvedCellWrite conversion. */
class CellValuePreparationTest {
  @Test
  void convertsAllSupportedInputTypesToExcelValues() {
    assertInstanceOf(
        CellValue.Blank.class, MutationPreparation.toResolvedCellWrite(new CellValueSpec.Blank()));

    dev.erst.gridgrind.core.model.value.ResolvedCellText textValue =
        assertInstanceOf(
            dev.erst.gridgrind.core.model.value.ResolvedCellText.class,
            MutationPreparation.toResolvedCellWrite(
                new CellValueSpec.Text(TextSourceSpec.inline("Budget"))));
    assertEquals("Budget", textValue.value());

    var inputFont =
        new dev.erst.gridgrind.core.model.style.CellFontPatch(
            java.util.Optional.of(true),
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            java.util.Optional.of(ColorSpec.rgb(new RgbHex("#AABBCC"))),
            java.util.Optional.empty(),
            java.util.Optional.empty());
    ResolvedRichText richTextValue =
        assertInstanceOf(
            ResolvedRichText.class,
            MutationPreparation.toResolvedCellWrite(
                new CellValueSpec.RichText(
                    List.of(
                        new dev.erst.gridgrind.core.model.style.RichTextRunSpec(
                            TextSourceSpec.inline("Budget"), java.util.Optional.empty()),
                        new dev.erst.gridgrind.core.model.style.RichTextRunSpec(
                            TextSourceSpec.inline(" FY26"), java.util.Optional.of(inputFont))))));
    assertEquals(
        new ResolvedRichText(
            List.of(
                new ResolvedRichTextRun("Budget", java.util.Optional.empty()),
                new ResolvedRichTextRun(
                    " FY26",
                    java.util.Optional.of(
                        new dev.erst.gridgrind.core.model.style.CellFontPatch(
                            java.util.Optional.of(Boolean.TRUE),
                            java.util.Optional.empty(),
                            java.util.Optional.empty(),
                            java.util.Optional.empty(),
                            java.util.Optional.of(ColorSpec.rgb(new RgbHex("#AABBCC"))),
                            java.util.Optional.empty(),
                            java.util.Optional.empty()))))),
        richTextValue);

    CellValueSpec.NumberValue numberValue =
        assertInstanceOf(
            CellValueSpec.NumberValue.class,
            MutationPreparation.toResolvedCellWrite(new CellValueSpec.NumberValue(42.5)));
    assertEquals(42.5, numberValue.value());

    CellValue.BooleanValue booleanValue =
        assertInstanceOf(
            CellValue.BooleanValue.class,
            MutationPreparation.toResolvedCellWrite(new CellValueSpec.BooleanValue(true)));
    assertTrue(booleanValue.value());

    ResolvedFormula formulaValue =
        assertInstanceOf(
            ResolvedFormula.class,
            MutationPreparation.toResolvedCellWrite(formulaCell("SUM(B2:B4)")));
    assertEquals("SUM(B2:B4)", formulaValue.value());

    assertThrows(
        IllegalArgumentException.class,
        () -> MutationPreparation.toResolvedCellWrite(formulaCell("=SUM(B2:B4)")));

    ResolvedRawFormula rawFormulaValue =
        assertInstanceOf(
            ResolvedRawFormula.class,
            MutationPreparation.toResolvedCellWrite(
                new CellValueSpec.RawFormula(TextSourceSpec.inline("LAMBDA(x,x+1)(A1)"))));
    assertEquals("LAMBDA(x,x+1)(A1)", rawFormulaValue.value());

    CellValueSpec.Date dateValue =
        assertInstanceOf(
            CellValueSpec.Date.class,
            MutationPreparation.toResolvedCellWrite(
                new CellValueSpec.Date(LocalDate.of(2026, 3, 23))));
    assertEquals(LocalDate.of(2026, 3, 23), dateValue.value());

    CellValueSpec.DateTime dateTimeValue =
        assertInstanceOf(
            CellValueSpec.DateTime.class,
            MutationPreparation.toResolvedCellWrite(
                new CellValueSpec.DateTime(LocalDateTime.of(2026, 3, 23, 10, 15, 30))));
    assertEquals(LocalDateTime.of(2026, 3, 23, 10, 15, 30), dateTimeValue.value());
  }

  @Test
  void validatesTypedInputRequirements() {
    assertThrows(IllegalArgumentException.class, () -> new CellValueSpec.Text(null));
    assertThrows(IllegalArgumentException.class, () -> new CellValueSpec.RichText(null));
    assertThrows(IllegalArgumentException.class, () -> new CellValueSpec.RichText(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellValueSpec.RichText(
                List.of(
                    new dev.erst.gridgrind.core.model.style.RichTextRunSpec(
                        TextSourceSpec.inline(""), java.util.Optional.empty()))));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellValueSpec.NumberValue(Double.POSITIVE_INFINITY));
    assertThrows(IllegalArgumentException.class, () -> new CellValueSpec.NumberValue(Double.NaN));
    assertThrows(IllegalArgumentException.class, () -> new CellValueSpec.Formula(null));
    assertThrows(IllegalArgumentException.class, () -> formulaCell("="));
    assertThrows(IllegalArgumentException.class, () -> formulaCell("=   "));
    assertThrows(IllegalArgumentException.class, () -> new CellValueSpec.Date(null));
    assertThrows(IllegalArgumentException.class, () -> new CellValueSpec.DateTime(null));
  }

  @Test
  void rejectsDdeFormulaInjectionInAllCaseForms() {
    IllegalArgumentException ddeUpper =
        assertThrows(
            IllegalArgumentException.class, () -> formulaCell("DDE(\"cmd\",\"/C calc\",\"\")"));
    assertTrue(ddeUpper.getMessage().contains("DDE"));

    assertThrows(
        IllegalArgumentException.class, () -> formulaCell("dde(\"cmd\",\"/C calc\",\"\")"));
    assertThrows(
        IllegalArgumentException.class, () -> formulaCell("Dde(\"cmd\",\"/C calc\",\"\")"));
    assertThrows(
        IllegalArgumentException.class, () -> formulaCell("=DDE(\"cmd\",\"/C calc\",\"\")"));
    assertThrows(
        IllegalArgumentException.class, () -> formulaCell("=dde(\"cmd\",\"/C calc\",\"\")"));
  }

  private static CellValueSpec.Formula formulaCell(String formula) {
    return new CellValueSpec.Formula(TextSourceSpec.inline(formula));
  }
}
