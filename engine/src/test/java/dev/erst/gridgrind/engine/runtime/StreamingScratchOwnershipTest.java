package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.security.OoxmlEncryptionSpec;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecuritySpec;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceSignatureSpec;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.source.SecretReference;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.engine.api.GridGrindEngine;
import dev.erst.gridgrind.engine.api.GridGrindProgressSink;
import dev.erst.gridgrind.engine.api.GridGrindRequestInputs;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

/** Encryption copies the serialized package; its plaintext lifetime remains owned by the run. */
class StreamingScratchOwnershipTest {
  @TempDir Path root;

  @ParameterizedTest
  @ValueSource(booleans = {false, true})
  void finalizationRemovesEveryOwnedScratchFileAndPreservesHostFiles(boolean encrypted)
      throws Exception {
    Path scratch = Files.createDirectory(root.resolve("scratch"));
    Path hostFile = Files.writeString(scratch.resolve("host-owned.txt"), "preserve host material");
    SecretReference password = new SecretReference("encryption");
    Path output = root.resolve("output.xlsx");
    var security =
        new OoxmlPersistenceSecuritySpec(
            encrypted
                ? new OoxmlPersistenceEncryptionSpec.Encrypt(OoxmlEncryptionSpec.strong(password))
                : new OoxmlPersistenceEncryptionSpec.None(),
            new OoxmlPersistenceSignatureSpec.None());
    var plan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                "output.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT, security),
            ExecutionPolicySpec.mode(ExecutionModeSpec.streamingWrite()),
            FormulaEnvironmentSpec.empty(),
            List.of(
                new MutationStep(
                    new StepId("ensure"),
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new WorkbookMutationAction.EnsureSheet())));
    var grant =
        new ExecutionGrant(
            0,
            List.of(),
            List.of("ENSURE_SHEET"),
            new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new ExecutionGrant.PublicationAuthority.SaveAs(
                output, WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            encrypted ? List.of(password) : List.of(),
            AcceptancePolicy.minimum());
    var inputs =
        new GridGrindRequestInputs(
            root, scratch, grant, reference -> "synthetic-test-password".toCharArray());
    assertInstanceOf(
        WorkbookResult.Success.class,
        GridGrindEngine.requestExecutor().execute(plan, inputs, GridGrindProgressSink.NOOP));
    try (var paths = Files.walk(scratch)) {
      assertEquals(List.of(scratch, hostFile), paths.sorted().toList());
    }
    assertEquals("preserve host material", Files.readString(hostFile));
    if (encrypted) {
      try (var filesystem = new POIFSFileSystem(output.toFile(), true)) {
        var decryptor = Decryptor.getInstance(new EncryptionInfo(filesystem));
        assertTrue(decryptor.verifyPassword("synthetic-test-password"));
        try (var plaintext = decryptor.getDataStream(filesystem);
            var workbook = new XSSFWorkbook(plaintext)) {
          assertEquals("Budget", workbook.getSheetName(0));
        }
      }
    } else {
      try (var workbook = new XSSFWorkbook(output.toFile())) {
        assertEquals("Budget", workbook.getSheetName(0));
      }
    }
  }
}
