package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.dto.OoxmlPersistenceEncryptionInput;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceSignatureInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies staged-artifact reopening preconditions and source-security preservation boundaries. */
class ExecutionWorkbookSupportTest {
  @TempDir Path root;

  @Test
  void rejectsNonFileStagedMaterializationBeforeAttemptingWorkbookReopen() throws Exception {
    IOException failure =
        assertThrows(
            IOException.class,
            () ->
                ExecutionWorkbookSupport.requireRegularFile(
                    Files.createDirectory(root.resolve("staged"))));

    assertEquals(
        "staged workbook materialization did not produce a regular file", failure.getMessage());
  }

  @Test
  void refusesToReopenPreservedSourceEncryptionWithoutTheVerifiedSourcePassword() {
    WorkbookPlan.WorkbookPersistence persistence =
        new WorkbookPlan.WorkbookPersistence.SaveAs(
            "output.xlsx",
            WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
            new OoxmlPersistenceSecurityInput(
                new OoxmlPersistenceEncryptionInput.PreserveSource(),
                new OoxmlPersistenceSignatureInput.None()));
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            root, root.resolve("scratch"), ExecutionGrantTestSupport.noPublication());

    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                ExecutionWorkbookSupport.stagedArtifactOpenOptions(
                    ExecutionRequestPaths.persistenceOptions(persistence, bindings),
                    new WorkbookPlan.WorkbookSource.New(),
                    bindings));

    assertEquals(
        "preserved source encryption requires a verified source password", failure.getMessage());
  }
}
