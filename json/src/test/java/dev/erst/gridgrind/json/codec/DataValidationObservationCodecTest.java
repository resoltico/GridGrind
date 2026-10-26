package dev.erst.gridgrind.json.codec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertReport;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptReport;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import org.junit.jupiter.api.Test;

/** Observed validation messages serialize as text, without artificial authored sources. */
class DataValidationObservationCodecTest {
  @Test
  void promptAndErrorMessagesRoundTripAsPlainObservedText() {
    var mapper = GridGrindJsonMapperSupport.WIRE_JSON_MAPPER;
    var prompt = new DataValidationPromptReport("Select", "Choose a value", true);
    String promptJson = mapper.writeValueAsString(prompt);
    var promptTree = mapper.readTree(promptJson);
    assertEquals("Select", promptTree.path("title").asText());
    assertEquals("Choose a value", promptTree.path("text").asText());
    assertEquals(prompt, mapper.readValue(promptJson, DataValidationPromptReport.class));
    var error =
        new DataValidationErrorAlertReport(
            ExcelDataValidationErrorStyle.STOP, "Invalid", "Try again", true);
    String errorJson = mapper.writeValueAsString(error);
    assertEquals("Invalid", mapper.readTree(errorJson).path("title").asText());
    assertEquals(error, mapper.readValue(errorJson, DataValidationErrorAlertReport.class));
  }
}
