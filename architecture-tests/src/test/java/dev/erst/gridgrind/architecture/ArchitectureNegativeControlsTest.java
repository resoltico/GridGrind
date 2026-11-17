package dev.erst.gridgrind.architecture;

import static com.tngtech.archunit.lang.syntax.ArchRuleDefinition.classes;
import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.tngtech.archunit.core.importer.ClassFileImporter;
import dev.erst.gridgrind.architecture.fixture.ArchitectureCatalogWithoutSemantics;
import dev.erst.gridgrind.architecture.fixture.ArchitectureMethodReferenceViolationFixture;
import dev.erst.gridgrind.architecture.fixture.ArchitectureNonWireSealedClass;
import dev.erst.gridgrind.architecture.fixture.ArchitectureNonWireSealedInterface;
import dev.erst.gridgrind.architecture.fixture.ArchitectureRuntimeLeakFixture;
import dev.erst.gridgrind.architecture.fixture.ArchitectureWireUnionVariants;
import org.junit.jupiter.api.Test;

/** One focused counterexample for each architecture property beyond JPMS enforcement. */
class ArchitectureNegativeControlsTest {
  @Test
  void unionTagsRejectOtherNamesAndUnownedStructuralExceptions() {
    AssertionError failure =
        assertThrows(
            AssertionError.class,
            () ->
                ProductDomainShapeArchitectureRules.unionTagsUseTypeOrStatus()
                    .check(
                        new ClassFileImporter()
                            .importClasses(
                                dev.erst.gridgrind.architecture.fixture
                                    .ArchitectureDiscriminatorViolations.WrongKind.class,
                                dev.erst.gridgrind.architecture.fixture
                                    .ArchitectureDiscriminatorViolations.WrongEmpty.class)));
    assertTrue(failure.getMessage().contains("WrongKind"));
    assertTrue(failure.getMessage().contains("WrongEmpty"));
  }

  @Test
  void coreRejectsAdapterDependencies() {
    AssertionError failure =
        assertThrows(
            AssertionError.class,
            () ->
                ProductDependencyArchitectureRules.coreDoesNotDependOnAdapters()
                    .check(
                        new ClassFileImporter()
                            .importClasses(
                                dev.erst.gridgrind.core.fixture.ArchitectureCoreAdapterDependency
                                    .class)));
    assertTrue(failure.getMessage().contains("GridGrindRequestInputs"));
  }

  @Test
  void coreRejectsWorkbookAndJsonLibraryDependencies() {
    AssertionError failure =
        assertThrows(
            AssertionError.class,
            () ->
                ProductDependencyArchitectureRules.coreDoesNotDependOnWorkbookOrJsonLibraries()
                    .check(
                        new ClassFileImporter()
                            .importClasses(
                                dev.erst.gridgrind.core.fixture.ArchitectureCoreLibraryLeak
                                    .class)));
    assertTrue(failure.getMessage().contains("org.apache.poi.ss.usermodel.Workbook"));
    assertTrue(failure.getMessage().contains("tools.jackson.databind.JsonNode"));
  }

  @Test
  void runtimeRejectsPoiAndRawSpreadsheetXmlDependencies() {
    AssertionError failure =
        assertThrows(
            AssertionError.class,
            () ->
                ProductDependencyArchitectureRules.poiDependenciesStayInNativeAdapters()
                    .check(
                        new ClassFileImporter()
                            .importClasses(
                                dev.erst.gridgrind.engine.runtime.fixture
                                    .ArchitectureNativeLibraryLeak.class)));
    assertTrue(failure.getMessage().contains("org.apache.poi.ss.usermodel.Workbook"));
    assertTrue(
        failure
            .getMessage()
            .contains("org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCell"));
  }

  @Test
  void workbookMechanicsRejectRuntimeDependencies() {
    AssertionError failure =
        assertThrows(
            AssertionError.class,
            () ->
                ProductDependencyArchitectureRules
                    .workbookImplementationDoesNotDependOnExecutionRuntime()
                    .check(
                        new ClassFileImporter()
                            .importPackages("dev.erst.gridgrind.engine.backend.xssf")));
    assertTrue(failure.getMessage().contains("ArchitectureRuntimeDependencyFixture"));
  }

  @Test
  void apiSignaturesRejectInternalAndPoiTypesIncludingGenerics() {
    AssertionError failure =
        assertThrows(
            AssertionError.class,
            () ->
                classes()
                    .should(
                        ProductDependencyArchitectureRules
                            .exportedApiMustNotExposeImplementationTypes())
                    .check(
                        new ClassFileImporter()
                            .importClasses(ArchitectureRuntimeLeakFixture.class)));
    assertTrue(failure.getMessage().contains("GridGrindRequestDoctor"));
    assertTrue(failure.getMessage().contains("XSSFWorkbook"));
  }

  @Test
  void formulaWritesRejectMethodReferencesOutsideTheGateway() {
    assertThrows(
        AssertionError.class,
        () ->
            classes()
                .should(ProductToolingSeamArchitectureRules.directPoiFormulaWritesAreForbidden())
                .check(
                    new ClassFileImporter()
                        .importClasses(ArchitectureMethodReferenceViolationFixture.class)));
  }

  @Test
  void privateAccessRejectsMethodReferencesOutsideTheGateway() {
    assertThrows(
        AssertionError.class,
        () ->
            classes()
                .should(ProductToolingSeamArchitectureRules.privateReflectionIsForbidden())
                .check(
                    new ClassFileImporter()
                        .importClasses(ArchitectureMethodReferenceViolationFixture.class)));
  }

  @Test
  void domainNamesAreIndependentOfDependencyBoundaries() {
    assertDoesNotThrow(() ->
        ProductDependencyArchitectureRules.coreDoesNotDependOnWorkbookOrJsonLibraries()
            .check(
                new ClassFileImporter()
                    .importClasses(
                        dev.erst.gridgrind.core.model.fixture.ArchitectureModelNames.AuthoredInput
                            .class,
                        dev.erst.gridgrind.core.model.fixture.ArchitectureModelNames.AuthoredDto
                            .class,
                        dev.erst.gridgrind.core.model.fixture.ArchitectureModelNames.CellSnapshot
                            .class,
                        dev.erst.gridgrind.core.model.fixture.ArchitectureModelNames.CellDefinition
                            .class)));
  }

  @Test
  void nonWireRepresentationsMayUseOrdinaryOrSealedClasses() {
    assertDoesNotThrow(() ->
        ProductDomainShapeArchitectureRules.wireUnionsUseRecordVariants()
            .check(
                new ClassFileImporter()
                    .importClasses(
                        ArchitectureNonWireSealedInterface.class,
                        ArchitectureNonWireSealedClass.class,
                        ArchitectureWireUnionVariants.class,
                        ArchitectureWireUnionVariants.RecordBranch.class,
                        ArchitectureWireUnionVariants.RecordVariant.class)));
  }

  @Test
  void wireUnionsAcceptRecordLeavesAndRejectEnumOrOpenAlternatives() {
    ClassFileImporter importer = new ClassFileImporter();
    assertDoesNotThrow(() ->
        ProductDomainShapeArchitectureRules.wireUnionsUseRecordVariants()
            .check(
                importer.importClasses(
                    ArchitectureWireUnionVariants.class,
                    ArchitectureWireUnionVariants.RecordBranch.class,
                    ArchitectureWireUnionVariants.RecordVariant.class)));
    AssertionError failure =
        assertThrows(
            AssertionError.class,
            () ->
                ProductDomainShapeArchitectureRules.wireUnionsUseRecordVariants()
                    .check(
                        importer.importClasses(
                            ArchitectureWireUnionVariants.WithEnum.class,
                            ArchitectureWireUnionVariants.EnumVariant.class,
                            ArchitectureWireUnionVariants.OpenBranch.class)));
    assertTrue(failure.getMessage().contains("EnumVariant"));
    assertTrue(failure.getMessage().contains("OpenBranch"));
  }

  @Test
  void catalogProjectionRequiresCanonicalSemantics() {
    assertThrows(
        AssertionError.class,
        () ->
            classes()
                .should(ProductAssuranceArchitectureRules.canonicalOperationContractsAreRequired())
                .check(
                    new ClassFileImporter()
                        .importClasses(ArchitectureCatalogWithoutSemantics.class)));
  }

  @Test
  void publicationRejectsAnAdditionalRuntimeConsumer() {
    AssertionError failure =
        assertThrows(
            AssertionError.class,
            () ->
                ProductAssuranceArchitectureRules.publicationIsCentralizedBehindRequestPathAccess()
                    .check(
                        new ClassFileImporter()
                            .importPackages("dev.erst.gridgrind.engine.runtime")));
    assertTrue(failure.getMessage().contains("ArchitecturePublicationBypassFixture"));
  }
}
