package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;
import java.util.Objects;

/** Machine-readable starter scaffold for authoring one task against the exact protocol surface. */
public record TaskPlanTemplate(
    GridGrindProtocolVersion protocolVersion,
    TaskEntry task,
    WorkbookPlan requestTemplate,
    List<String> authoringNotes) {
  public TaskPlanTemplate {
    Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
    Objects.requireNonNull(task, "task must not be null");
    Objects.requireNonNull(requestTemplate, "requestTemplate must not be null");
    authoringNotes = CatalogRecordValidation.copyStrings(authoringNotes, "authoringNotes");
  }
}
