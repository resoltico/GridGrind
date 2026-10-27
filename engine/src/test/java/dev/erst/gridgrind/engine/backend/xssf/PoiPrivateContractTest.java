package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertAll;
import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.engine.poi.FillRegistryAccess;
import dev.erst.gridgrind.engine.poi.ImageCatalog;
import dev.erst.gridgrind.engine.poi.PrivateContract;
import java.lang.invoke.MethodHandles;
import java.util.List;
import org.apache.poi.Version;
import org.junit.jupiter.api.Test;

/** Fails fast when a POI upgrade breaks one of GridGrind's explicit private engine seams. */
class PoiPrivateContractTest {
  @Test
  void currentPoiVersionSatisfiesEveryRegisteredPrivateEngineContract() {
    MethodHandles.Lookup lookup = MethodHandles.lookup();

    assertAll(
        () -> assertDoesNotThrow(() -> PoiRelationRemoval.removePoiRelationInvoker(lookup)),
        () -> assertDoesNotThrow(() -> FillRegistryAccess.requireFillsField(lookup)),
        () -> assertDoesNotThrow(() -> ImageCatalog.requirePicturesField(lookup)),
        () -> assertDoesNotThrow(() -> ImageCatalog.requirePictureConstructor(lookup)));
  }

  @Test
  void contractFailureMessagesNameTheBrokenPoiSurfaceAndAffectedFeature() {
    IllegalStateException relationFailure =
        assertThrows(
            IllegalStateException.class,
            () -> PoiRelationRemoval.removePoiRelationInvoker(MethodHandles.publicLookup()));
    IllegalStateException pictureFailure =
        assertThrows(
            IllegalStateException.class,
            () -> ImageCatalog.requirePicturesField(MethodHandles.publicLookup()));

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
                pictureFailure
                    .getMessage()
                    .contains(ImageCatalog.PICTURES_FIELD_CONTRACT.affectedSurface())),
        () ->
            assertTrue(
                pictureFailure
                    .getMessage()
                    .contains(ImageCatalog.PICTURES_FIELD_CONTRACT.memberSignature())));
  }

  @Test
  void failureMessagesUseTheActualPoiVersionForEveryContractOwner() {
    for (Class<?> owner : List.of(String.class, byte[].class, PoiPrivateContractTest.class)) {
      var contract = PrivateContract.field(owner, "missing", "diagnostic identity");
      assertTrue(contract.failureMessage().contains("Apache POI " + Version.getVersion()));
    }
  }
}
