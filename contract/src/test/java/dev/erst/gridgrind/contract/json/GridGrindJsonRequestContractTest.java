package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.ExecutionJournalLevel;
import dev.erst.gridgrind.contract.dto.OoxmlEncryptionInput;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceEncryptionInput;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceSignatureInput;
import dev.erst.gridgrind.contract.dto.OoxmlSignatureInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureDigestAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlWriteCipher;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlWriteHash;
import java.nio.charset.StandardCharsets;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Covers the request JSON contract, including top-level omission defaults. */
class GridGrindJsonRequestContractTest {
  @Test
  void requestRequiresExplicitProtocolVersion() {
    InvalidRequestShapeException exception =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "execution": {
                        "mode": {"type": "FULL_XSSF"},
                        "journal": { "level": "NORMAL" },
                        "calculation": {
                          "strategy": { "type": "DO_NOT_CALCULATE" },
                          "markRecalculateOnOpen": false
                        }
                      },
                      "formulaEnvironment": {
                        "externalWorkbooks": [],
                        "missingWorkbookPolicy": "ERROR",
                        "udfToolpacks": []
                      },
                      "steps": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertTrue(exception.getMessage().contains("protocolVersion"));
  }

  @Test
  void requestRequiresExplicitPersistenceAndSteps() {
    InvalidRequestShapeException exception =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "execution": {
                        "mode": {"type": "FULL_XSSF"},
                        "journal": { "level": "NORMAL" },
                        "calculation": {
                          "strategy": { "type": "DO_NOT_CALCULATE" },
                          "markRecalculateOnOpen": false
                        }
                      },
                      "formulaEnvironment": {
                        "externalWorkbooks": [],
                        "missingWorkbookPolicy": "ERROR",
                        "udfToolpacks": []
                      }
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertTrue(
        exception.getMessage().contains("persistence") || exception.getMessage().contains("steps"));
  }

  @Test
  void requestAllowsOmittedTopLevelExecutionAndFormulaEnvironment() throws Exception {
    WorkbookPlan request =
        GridGrindJson.readRequest(
            """
            {
              "protocolVersion": "V2",
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "steps": []
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertTrue(request.execution().isDefault());
    assertEquals(ExecutionJournalLevel.SUMMARY, request.journalLevel());
    assertTrue(request.formulaEnvironment().isEmpty());
  }

  @Test
  void requestAllowsSparseExecutionFormulaEnvironmentAndOoxmlSecurityBlocks() throws Exception {
    WorkbookPlan request =
        GridGrindJson.readRequest(
            """
            {
              "protocolVersion": "V2",
              "source": { "type": "NEW" },
	              "persistence": {
	                "type": "SAVE_AS",
	                "path": "secured.xlsx",
                "ifExists": "REJECT",
                "security": {
                  "encryption": {
                    "type": "ENCRYPT",
                    "encryption": {
                      "password": "persist-pass"
                    }
                  },
                  "signature": {
                    "type": "SIGN",
                    "signature": {
                      "pkcs12Path": "keys/signing.p12",
                      "keystorePassword": "store-pass"
                    }
                  }
                }
              },
              "execution": {
                "calculation": {}
              },
              "formulaEnvironment": {},
              "steps": []
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertTrue(request.execution().isDefault());
    assertEquals(ExecutionJournalLevel.SUMMARY, request.journalLevel());
    assertTrue(request.formulaEnvironment().isEmpty());
    OoxmlPersistenceSecurityInput security =
        ((WorkbookPlan.WorkbookPersistence.SaveAs) request.persistence()).security().orElseThrow();
    OoxmlEncryptionInput encryption =
        assertInstanceOf(OoxmlPersistenceEncryptionInput.Encrypt.class, security.encryption())
            .encryption();
    assertEquals(ExcelOoxmlWriteCipher.AES_256, encryption.cipher());
    assertEquals(ExcelOoxmlWriteHash.SHA_512, encryption.hash());
    OoxmlSignatureInput signature =
        assertInstanceOf(OoxmlPersistenceSignatureInput.Sign.class, security.signature())
            .signature();
    assertEquals("store-pass", signature.keyPassword());
    assertEquals(ExcelOoxmlSignatureDigestAlgorithm.SHA256, signature.digestAlgorithm());
  }

  @Test
  void requestRequiresBothExplicitPersistenceSecurityAxes() {
    InvalidRequestShapeException missingSignature =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": {
                        "type": "SAVE_AS",
                        "path": "secured.xlsx",
                        "ifExists": "REJECT",
                        "security": { "encryption": { "type": "NONE" } }
                      },
                      "steps": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));
    InvalidRequestShapeException missingEncryption =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": {
                        "type": "SAVE_AS",
                        "path": "secured.xlsx",
                        "ifExists": "REJECT",
                        "security": { "signature": { "type": "NONE" } }
                      },
                      "steps": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertTrue(missingSignature.getMessage().contains("signature"));
    assertTrue(missingEncryption.getMessage().contains("encryption"));
  }

  @Test
  void requestRejectsLegacyEncryptionModeFieldNowThatWriteEncryptionIsModeLess() {
    InvalidRequestShapeException exception =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": {
                        "type": "SAVE_AS",
                        "path": "secured.xlsx",
                        "ifExists": "REJECT",
                        "security": {
                          "encryption": {
                            "type": "ENCRYPT",
                            "encryption": {
                              "password": "persist-pass",
                              "mode": "AGILE"
                            }
                          }
                        }
                      },
                      "steps": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertTrue(exception.getMessage().contains("mode"));
  }

  @Test
  void requestRejectsLegacyConditionalFormattingRangesInsideTheActionBody() {
    InvalidRequestShapeException exception =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "steps": [ {
                        "stepId": "conditional-formatting",
                        "target": {
                          "type": "RANGE_BY_RANGE",
                          "sheetName": "Budget",
                          "range": "B2:B5"
                        },
                        "action": {
                          "type": "SET_CONDITIONAL_FORMATTING",
                          "conditionalFormatting": {
                            "ranges": [ "B2:B5" ],
                            "rules": [ {
                              "type": "FORMULA_RULE",
                              "formula": "B2>0",
                              "stopIfTrue": false
                            } ]
                          }
                        }
                      } ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertEquals(Optional.of("steps[0].action.conditionalFormatting.ranges"), exception.jsonPath());
    assertTrue(exception.getMessage().contains("ranges"));
  }

  @Test
  void requestRejectsUnsupportedWriteCipherValuesThatRemainReadableOnTheReportSurface() {
    InvalidRequestShapeException exception =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": {
                        "type": "SAVE_AS",
                        "path": "secured.xlsx",
                        "ifExists": "REJECT",
                        "security": {
                          "encryption": {
                            "type": "ENCRYPT",
                            "encryption": {
                              "password": "persist-pass",
                              "cipher": "AES_128"
                            }
                          }
                        }
                      },
                      "steps": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertEquals(
        "Unsupported value 'AES_128' for field 'persistence.security.encryption.encryption.cipher'; expected one of: AES_256, AES_192",
        exception.getMessage());
  }

  @Test
  void requestRejectsUnsupportedWriteHashValuesThatRemainReadableOnTheReportSurface() {
    InvalidRequestShapeException exception =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": {
                        "type": "SAVE_AS",
                        "path": "secured.xlsx",
                        "ifExists": "REJECT",
                        "security": {
                          "encryption": {
                            "type": "ENCRYPT",
                            "encryption": {
                              "password": "persist-pass",
                              "hash": "SHA_1"
                            }
                          }
                        }
                      },
                      "steps": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertEquals(
        "Unsupported value 'SHA_1' for field 'persistence.security.encryption.encryption.hash'; expected one of: SHA_512, SHA_384, SHA_256",
        exception.getMessage());
  }

  @Test
  void requestRejectsSparsePrintLayoutPayloadsThatPreviouslyReliedOnImplicitDefaults() {
    InvalidRequestShapeException exception =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "execution": {
                        "mode": {"type": "FULL_XSSF"},
                        "journal": { "level": "NORMAL" },
                        "calculation": {
                          "strategy": { "type": "DO_NOT_CALCULATE" },
                          "markRecalculateOnOpen": false
                        }
                      },
                      "formulaEnvironment": {
                        "externalWorkbooks": [],
                        "missingWorkbookPolicy": "ERROR",
                        "udfToolpacks": []
                      },
                      "steps": [ {
                        "stepId": "layout",
                        "target": { "type": "SHEET_BY_NAME", "name": "Sheet1" },
                        "action": {
                          "type": "SET_PRINT_LAYOUT",
                          "printLayout": {
                            "printArea": { "type": "NONE" },
                            "orientation": "PORTRAIT",
                            "scaling": { "type": "AUTOMATIC" },
                            "repeatingRows": { "type": "NONE" },
                            "repeatingColumns": { "type": "NONE" },
                            "header": {
                              "left": { "type": "INLINE", "text": "" },
                              "center": { "type": "INLINE", "text": "" },
                              "right": { "type": "INLINE", "text": "" }
                            },
                            "footer": {
                              "left": { "type": "INLINE", "text": "" },
                              "center": { "type": "INLINE", "text": "" },
                              "right": { "type": "INLINE", "text": "" }
                            }
                          }
                        }
                      } ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertTrue(exception.getMessage().contains("setup"));
  }

  @Test
  void requestRejectsSparseChartPayloadsThatPreviouslyReliedOnImplicitDefaults() {
    InvalidRequestShapeException exception =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "execution": {
                        "mode": {"type": "FULL_XSSF"},
                        "journal": { "level": "NORMAL" },
                        "calculation": {
                          "strategy": { "type": "DO_NOT_CALCULATE" },
                          "markRecalculateOnOpen": false
                        }
                      },
                      "formulaEnvironment": {
                        "externalWorkbooks": [],
                        "missingWorkbookPolicy": "ERROR",
                        "udfToolpacks": []
                      },
                      "steps": [ {
                        "stepId": "chart",
                        "target": { "type": "SHEET_BY_NAME", "name": "Sheet1" },
                        "action": {
                          "type": "SET_CHART",
                          "chart": {
                            "name": "RevenueChart",
                            "anchor": {
                              "type": "TWO_CELL",
                              "to": { "columnIndex": 8, "rowIndex": 14, "dx": 0, "dy": 0 },
                              "behavior": "MOVE_AND_RESIZE"
                            },
                            "title": { "type": "NONE" },
                            "legend": { "type": "VISIBLE", "position": "RIGHT" },
                            "displayBlanksAs": "GAP",
                            "plotOnlyVisibleCells": true,
                            "plots": [ {
                              "type": "LINE",
                              "varyColors": false,
                              "grouping": "STANDARD",
                              "axes": [
                                {
                                  "kind": "CATEGORY",
                                  "position": "BOTTOM",
                                  "crosses": "AUTO_ZERO",
                                  "visible": true
                                },
                                {
                                  "kind": "VALUE",
                                  "position": "LEFT",
                                  "crosses": "AUTO_ZERO",
                                  "visible": true
                                }
                              ],
                              "series": [ {
                                "title": { "type": "NONE" },
                                "categories": { "type": "REFERENCE", "formula": "Sheet1!$A$2:$A$5" },
                                "values": { "type": "REFERENCE", "formula": "Sheet1!$B$2:$B$5" }
                              } ]
                            } ]
                          }
                        }
                      } ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertTrue(
        exception
            .getMessage()
            .contains("Missing required field 'steps[0].action.chart.anchor.from'"),
        exception::getMessage);
  }
}
