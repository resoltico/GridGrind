package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;
import java.util.Objects;

/** Internal task-definition record that couples discovery metadata with a starter plan. */
record TaskDefinition(
    TaskEntry task,
    List<String> searchTerms,
    WorkbookPlan requestTemplate,
    List<String> authoringNotes) {
  TaskDefinition {
    Objects.requireNonNull(task, "task must not be null");
    searchTerms = CatalogRecordValidation.copyStrings(searchTerms, "searchTerms");
    Objects.requireNonNull(requestTemplate, "requestTemplate must not be null");
    authoringNotes = CatalogRecordValidation.copyStrings(authoringNotes, "authoringNotes");
  }

  TaskPlanTemplate starterTemplate() {
    return new TaskPlanTemplate(
        GridGrindProtocolVersion.current(), task, requestTemplate, authoringNotes);
  }
}
