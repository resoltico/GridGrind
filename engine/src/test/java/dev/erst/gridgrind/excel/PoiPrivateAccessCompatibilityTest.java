package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertAll;
import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.lang.invoke.MethodHandles;
import org.junit.jupiter.api.Test;

/** Fails fast when a POI upgrade breaks one of GridGrind's explicit private engine seams. */
class PoiPrivateAccessCompatibilityTest {
  @Test
  void currentPoiVersionSatisfiesEveryRegisteredPrivateEngineContract() {
    MethodHandles.Lookup lookup = MethodHandles.lookup();

    assertAll(
        () -> assertDoesNotThrow(() -> PoiRelationRemoval.removePoiRelationInvoker(lookup)),
        () -> assertDoesNotThrow(() -> StylesTableFillRegistryAccess.requireFillsField(lookup)),
        () ->
            assertDoesNotThrow(
                () -> ExcelSheetClonePreparationSupport.requireHyperlinksField(lookup)),
        () ->
            assertDoesNotThrow(
                () -> ExcelSheetClonePreparationSupport.requireHyperlinkConstructor(lookup)),
        () ->
            assertDoesNotThrow(() -> ExcelWorkbookImageCatalogSupport.requirePicturesField(lookup)),
        () ->
            assertDoesNotThrow(
                () -> ExcelWorkbookImageCatalogSupport.requirePictureConstructor(lookup)));
  }

  @Test
  void contractFailureMessagesNameTheBrokenPoiSurfaceAndAffectedFeature() {
    IllegalStateException relationFailure =
        assertThrows(
            IllegalStateException.class,
            () -> PoiRelationRemoval.removePoiRelationInvoker(MethodHandles.publicLookup()));
    IllegalStateException hyperlinkFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                ExcelSheetClonePreparationSupport.requireHyperlinksField(
                    MethodHandles.publicLookup()));
    IllegalStateException pictureFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                ExcelWorkbookImageCatalogSupport.requirePicturesField(
                    MethodHandles.publicLookup()));

    assertAll(
        () ->
            assertTrue(
                relationFailure.getMessage().contains("Apache POI private contract unavailable")),
        () ->
            assertTrue(
                relationFailure
                    .getMessage()
                    .contains(PoiRelationRemoval.REMOVE_RELATION_CONTRACT.affectedSurface())),
        () ->
            assertTrue(
                relationFailure
                    .getMessage()
                    .contains(PoiRelationRemoval.REMOVE_RELATION_CONTRACT.memberSignature())),
        () ->
            assertTrue(
                hyperlinkFailure
                    .getMessage()
                    .contains(
                        ExcelSheetClonePreparationSupport.HYPERLINKS_FIELD_CONTRACT
                            .affectedSurface())),
        () ->
            assertTrue(
                hyperlinkFailure
                    .getMessage()
                    .contains(
                        ExcelSheetClonePreparationSupport.HYPERLINKS_FIELD_CONTRACT
                            .memberSignature())),
        () ->
            assertTrue(
                pictureFailure
                    .getMessage()
                    .contains(
                        ExcelWorkbookImageCatalogSupport.PICTURES_FIELD_CONTRACT
                            .affectedSurface())),
        () ->
            assertTrue(
                pictureFailure
                    .getMessage()
                    .contains(
                        ExcelWorkbookImageCatalogSupport.PICTURES_FIELD_CONTRACT
                            .memberSignature())));
  }

  @Test
  void failureMessagesCoverPoiVersionFallbackPaths() {
    PoiPrivateContract moduleVersionContract =
        PoiPrivateContract.field(String.class, "value", "module-version fallback");
    PoiPrivateContract noPackageContract =
        PoiPrivateContract.field(byte[].class, "length", "no-package fallback");
    PoiPrivateContract unknownVersionContract =
        PoiPrivateContract.field(
            PoiPrivateAccessCompatibilityTest.class, "class", "unknown-version fallback");

    assertAll(
        () ->
            assertTrue(
                moduleVersionContract
                    .failureMessage()
                    .contains("Apache POI private contract unavailable")),
        () -> assertTrue(moduleVersionContract.failureMessage().contains("Apache POI 26")),
        () -> assertTrue(noPackageContract.failureMessage().contains("Apache POI 26")),
        () -> assertTrue(unknownVersionContract.failureMessage().contains("Apache POI unknown")));
  }
}
