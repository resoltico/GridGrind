package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.Preservation;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Enumeration;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/**
 * Proves host-required part identity is evaluated on the verified staged package before publish.
 */
class OpaqueOoxmlPartPreservationExecutionTest {
  @TempDir Path root;

  @Test
  void preservesAnExplicitOpaqueCustomPartAcrossAFullXssfSaveAs() throws Exception {
    Path source = sourceWithCustomPart();
    Path output = root.resolve("verified-output.xlsx");
    WorkbookPlan request = updatePlan(source, output);

    WorkbookResult.Success success =
        assertInstanceOf(
            WorkbookResult.Success.class,
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request,
                    ExecutionInputBindingsFixtureSupport.bindings(
                        root, request, preservationPolicy("/customXml/item1.xml")),
                    ExecutionProgressSink.NOOP));

    Preservation.ByteEstablished established =
        assertInstanceOf(
            Preservation.ByteEstablished.class, success.evidence().preservation().getFirst());
    assertEquals(List.of("/customXml/item1.xml"), established.partNames());
    assertArrayEquals(
        partBytes(source, "customXml/item1.xml"), partBytes(output, "customXml/item1.xml"));
  }

  @Test
  void blocksPublicationWhenARequiredPartChangesDuringSerialization() throws Exception {
    Path source = sourceWithCustomPart();
    Path output = root.resolve("rejected-output.xlsx");
    WorkbookPlan request = updatePlan(source, output);

    WorkbookResult.Failure failure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request,
                    ExecutionInputBindingsFixtureSupport.bindings(
                        root, request, preservationPolicy("/xl/worksheets/sheet1.xml")),
                    ExecutionProgressSink.NOOP));

    assertEquals(GridGrindProblemCode.PRESERVATION_FAILED, failure.problem().code());
    assertFalse(Files.exists(output));
    Preservation.Failed preservation =
        assertInstanceOf(Preservation.Failed.class, failure.evidence().preservation().getFirst());
    assertEquals(Preservation.Kind.BYTE, preservation.kind());
    assertEquals(List.of("/xl/worksheets/sheet1.xml"), preservation.identifiers());
  }

  private Path sourceWithCustomPart() throws IOException {
    Path source = root.resolve("source.xlsx");
    try (XSSFWorkbook workbook = new XSSFWorkbook();
        var output = Files.newOutputStream(source)) {
      workbook.createSheet("Report").createRow(0).createCell(0).setCellValue(1);
      workbook.write(output);
    }
    appendPart(source, "customXml/item1.xml", "<report><version>1</version></report>");
    return source;
  }

  private static WorkbookPlan updatePlan(Path source, Path output) {
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.ExistingFile(source.toString()),
        new WorkbookPlan.WorkbookPersistence.SaveAs(
            output.toString(),
            WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
            dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecuritySpec.none()),
        ExecutionPolicySpec.defaults(),
        FormulaEnvironmentSpec.empty(),
        List.of(
            new MutationStep(
                new StepId("update-report"),
                new CellSelector.ByAddress(new SheetName("Report"), new CellAddress("A1")),
                new CellMutationAction.SetCell(new CellValueSpec.NumberValue(2)))));
  }

  private static AcceptancePolicy preservationPolicy(String partName) {
    return new AcceptancePolicy.RequireAll(
        List.of(new AcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts(List.of(partName))));
  }

  private void appendPart(Path archive, String partName, String contents) throws IOException {
    Path rewritten = Files.createTempFile(root, "source-rewrite-", ".xlsx");
    try (ZipFile source = new ZipFile(archive.toFile());
        ZipOutputStream target = new ZipOutputStream(Files.newOutputStream(rewritten))) {
      Enumeration<? extends ZipEntry> entries = source.entries();
      while (entries.hasMoreElements()) {
        copyEntry(source, target, entries.nextElement());
      }
      target.putNextEntry(new ZipEntry(partName));
      target.write(contents.getBytes(java.nio.charset.StandardCharsets.UTF_8));
      target.closeEntry();
    }
    Files.move(rewritten, archive, StandardCopyOption.REPLACE_EXISTING);
  }

  private static void copyEntry(ZipFile source, ZipOutputStream target, ZipEntry entry)
      throws IOException {
    target.putNextEntry(new ZipEntry(entry.getName()));
    try (InputStream input = source.getInputStream(entry)) {
      input.transferTo(target);
    }
    target.closeEntry();
  }

  private static byte[] partBytes(Path archive, String partName) throws IOException {
    try (ZipFile zip = new ZipFile(archive.toFile());
        InputStream input = zip.getInputStream(zip.getEntry(partName))) {
      return input.readAllBytes();
    }
  }
}
