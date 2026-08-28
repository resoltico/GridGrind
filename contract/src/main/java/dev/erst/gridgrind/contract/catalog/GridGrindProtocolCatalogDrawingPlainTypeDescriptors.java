package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.ArrayFormulaInput;
import dev.erst.gridgrind.contract.dto.ChartAxisInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.ChartSeriesInput;
import dev.erst.gridgrind.contract.dto.CustomXmlImportInput;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator;
import dev.erst.gridgrind.contract.dto.DrawingMarkerInput;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.SignatureLineInput;
import java.util.List;

/** Drawing, charting, media, and custom-XML authoring plain type descriptors. */
final class GridGrindProtocolCatalogDrawingPlainTypeDescriptors {
  private GridGrindProtocolCatalogDrawingPlainTypeDescriptors() {}

  static final List<CatalogPlainTypeDescriptor> DESCRIPTORS =
      List.of(
          plainTypeDescriptor(
              "drawingMarkerInputType",
              DrawingMarkerInput.class,
              "DrawingMarkerInput",
              "One zero-based drawing marker with explicit column, row, and in-cell offsets."),
          plainTypeDescriptor(
              "arrayFormulaInputType",
              ArrayFormulaInput.class,
              "ArrayFormulaInput",
              "One authored array formula bound to a contiguous single-cell or multi-cell range."
                  + " Leading = or {=...} wrappers normalize away for inline sources."),
          plainTypeDescriptor(
              "customXmlMappingLocatorType",
              CustomXmlMappingLocator.class,
              "CustomXmlMappingLocator",
              "One locator for an existing workbook custom-XML mapping."
                  + " Supply mapId, name, or both; the locator must resolve to exactly one"
                  + " existing mapping."),
          plainTypeDescriptor(
              "customXmlImportInputType",
              CustomXmlImportInput.class,
              "CustomXmlImportInput",
              "One custom-XML import payload targeting an existing workbook mapping plus the XML"
                  + " content to import."),
          plainTypeDescriptor(
              "chartInputType",
              ChartInput.class,
              "ChartInput",
              "One authored chart with an explicit drawing anchor, chart-level presentation"
                  + " state, and one or more plots."
                  + " Use ChartInput.withStandardDisplay(...) from Java authoring when the"
                  + " standard legend and blank-cell display policy is intended."),
          plainTypeDescriptor(
              "chartAxisInputType",
              ChartAxisInput.class,
              "ChartAxisInput",
              "One authored chart axis used by a chart plot with explicit visibility state."),
          plainTypeDescriptor(
              "chartSeriesInputType",
              ChartSeriesInput.class,
              "ChartSeriesInput",
              "One authored chart series with a title plus category and value data sources."
                  + " smooth, marker, and explosion fields are optional by chart family."),
          plainTypeDescriptor(
              "pictureDataInputType",
              PictureDataInput.class,
              "PictureDataInput",
              "One picture payload with explicit format and base64-encoded binary data."),
          plainTypeDescriptor(
              "pictureInputType",
              PictureInput.class,
              "PictureInput",
              "Named picture-authoring payload for SET_PICTURE."),
          plainTypeDescriptor(
              "signatureLineInputType",
              SignatureLineInput.class,
              "SignatureLineInput",
              "Named signature-line authoring payload for SET_SIGNATURE_LINE."
                  + " allowComments is explicit and plainSignature is optional,"
                  + " but caption or suggested signer metadata must still be present."),
          plainTypeDescriptor(
              "embeddedObjectInputType",
              EmbeddedObjectInput.class,
              "EmbeddedObjectInput",
              "Named embedded-object authoring payload for SET_EMBEDDED_OBJECT."
                  + " base64Data holds the embedded package bytes and previewImage holds the"
                  + " visible preview raster."));

  private static CatalogPlainTypeDescriptor plainTypeDescriptor(
      String group, Class<? extends Record> recordType, String id, String summary) {
    return CatalogTypeEntryFactory.plainTypeDescriptor(group, recordType, id, summary);
  }
}
