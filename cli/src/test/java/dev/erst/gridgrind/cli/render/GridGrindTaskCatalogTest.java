package dev.erst.gridgrind.cli.render;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.discovery.TaskCatalog;
import dev.erst.gridgrind.core.discovery.TaskCatalogs;
import java.io.IOException;
import org.junit.jupiter.api.Test;

/** Tests for the high-level task catalog layered on top of the exact protocol catalog. */
class GridGrindTaskCatalogTest {
  @Test
  void exposesDeterministicTaskCatalogEntries() throws IOException {
    TaskCatalog catalog = TaskCatalogs.catalog();
    TaskCatalog decoded =
        GridGrindCliJson.readBytes(GridGrindCliJson.writeBytes(catalog), TaskCatalog.class);

    assertFalse(catalog.tasks().isEmpty());
    assertEquals(catalog, decoded);
    assertTrue(TaskCatalogs.entryFor("DASHBOARD").isPresent());
    assertEquals(
        "inspectionQueryTypes:GET_CHARTS",
        TaskCatalogs.entryFor("DASHBOARD")
            .orElseThrow()
            .workflow()
            .phases()
            .getLast()
            .capabilityRefs()
            .getFirst()
            .qualifiedId());
    assertTrue(TaskCatalogs.entryFor("TABULAR_REPORT").isPresent());
    assertTrue(TaskCatalogs.entryFor("AUDIT_EXISTING_WORKBOOK").isPresent());
    assertTrue(TaskCatalogs.entryFor("DATA_ENTRY_WORKFLOW").isPresent());
    assertTrue(TaskCatalogs.entryFor("CUSTOM_XML_WORKFLOW").isPresent());
    assertTrue(TaskCatalogs.entryFor("PIVOT_REPORT").isPresent());
    assertTrue(TaskCatalogs.entryFor("DRAWING_AND_SIGNATURE_WORKFLOW").isPresent());
    assertTrue(TaskCatalogs.entryFor("WORKBOOK_MAINTENANCE").isPresent());
    assertTrue(TaskCatalogs.entryFor("BOGUS_TASK").isEmpty());
  }
}
