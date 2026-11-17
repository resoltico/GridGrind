package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import java.util.List;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

/** Document-link analysis applies concrete geometry rules without rewriting observed targets. */
class DocumentHyperlinkGeometryTest {
  @ParameterizedTest
  @ValueSource(strings = {"A0", "XFE1", "A1048577", "B2:A1", "A1:B2:C3"})
  void invalidDocumentCoordinatesProduceFindingsWhileRawTargetsRemainVisible(String reference)
      throws Exception {
    try (var workbook = XssfWorkbooks.create()) {
      var sheet = workbook.getOrCreateSheet("Grid");
      String target = "Grid!" + reference;
      sheet.links().setHyperlink(new CellAddress(0, 0), new HyperlinkTarget.Document(target));
      var findings = sheet.links().hyperlinkHealthFindings();
      assertEquals(1, findings.size());
      assertEquals(
          AnalysisFindingCode.HYPERLINK_INVALID_DOCUMENT_TARGET, findings.getFirst().code());
      assertEquals(target, sheet.xssfSheet().getRow(0).getCell(0).getHyperlink().getAddress());
    }
  }

  @ParameterizedTest
  @ValueSource(strings = {"A1", "$xFd$1048576", "$a$1:$b$2"})
  void validDocumentCoordinatesAcceptAnchorsCaseAndTheLastCell(String reference) throws Exception {
    try (var workbook = XssfWorkbooks.create()) {
      var sheet = workbook.getOrCreateSheet("Grid");
      String target = "Grid!" + reference;
      sheet.links().setHyperlink(new CellAddress(0, 0), new HyperlinkTarget.Document(target));
      assertEquals(List.of(), sheet.links().hyperlinkHealthFindings());
      assertEquals(target, sheet.xssfSheet().getRow(0).getCell(0).getHyperlink().getAddress());
    }
  }
}
