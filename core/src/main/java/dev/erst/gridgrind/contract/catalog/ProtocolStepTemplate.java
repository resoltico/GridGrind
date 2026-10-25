package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.Objects;
import tools.jackson.databind.node.ObjectNode;

/** Machine-readable example showing how one operation is placed inside a workbook step. */
public record ProtocolStepTemplate(
    String stepKind, String bodyField, ObjectNode template, List<String> notes) {
  public ProtocolStepTemplate {
    stepKind = CatalogRecordValidation.requireNonBlank(stepKind, "stepKind");
    bodyField = CatalogRecordValidation.requireNonBlank(bodyField, "bodyField");
    template = Objects.requireNonNull(template, "template must not be null").deepCopy();
    notes = CatalogRecordValidation.copyStrings(notes, "notes");
  }
}
