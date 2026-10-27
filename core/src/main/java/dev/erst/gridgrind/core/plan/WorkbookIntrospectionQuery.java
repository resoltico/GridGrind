package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingLocator;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.nio.charset.StandardCharsets;
import java.util.Objects;

/** Workbook-scoped factual inspection queries. */
public sealed interface WorkbookIntrospectionQuery extends InspectionQuery.Introspection
    permits WorkbookIntrospectionQuery.GetWorkbookSummary,
        WorkbookIntrospectionQuery.GetPackageSecurity,
        WorkbookIntrospectionQuery.GetWorkbookProtection,
        WorkbookIntrospectionQuery.GetCustomXmlMappings,
        WorkbookIntrospectionQuery.ExportCustomXmlMapping,
        WorkbookIntrospectionQuery.GetNamedRanges {

  @Variant(
      id = "GET_WORKBOOK_SUMMARY",
      summary = "Return workbook-level summary facts.",
      targetSelectors = {WorkbookSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record GetWorkbookSummary() implements WorkbookIntrospectionQuery {}

  @Variant(
      id = "GET_PACKAGE_SECURITY",
      summary = "Return OOXML package-encryption and package-signature facts.",
      targetSelectors = {WorkbookSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record GetPackageSecurity() implements WorkbookIntrospectionQuery {}

  @Variant(
      id = "GET_WORKBOOK_PROTECTION",
      summary = "Return workbook-level protection facts.",
      targetSelectors = {WorkbookSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record GetWorkbookProtection() implements WorkbookIntrospectionQuery {}

  @Variant(
      id = "GET_CUSTOM_XML_MAPPINGS",
      summary = "Return workbook custom-XML mapping metadata.",
      targetSelectors = {WorkbookSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record GetCustomXmlMappings() implements WorkbookIntrospectionQuery {}

  @Variant(
      id = "EXPORT_CUSTOM_XML_MAPPING",
      summary = "Export one existing workbook custom-XML mapping as serialized XML.",
      targetSelectors = {WorkbookSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record ExportCustomXmlMapping(
      CustomXmlMappingLocator mapping,
      @ProtocolField(optional = true) @WireDefault("false") boolean validateSchema,
      @ProtocolField(optional = true) @WireDefault("UTF-8") String encoding)
      implements WorkbookIntrospectionQuery {
    /** Creates one UTF-8 export request without requiring the caller to pass the encoding. */
    public ExportCustomXmlMapping(CustomXmlMappingLocator mapping, boolean validateSchema) {
      this(mapping, validateSchema, StandardCharsets.UTF_8.name());
    }

    public ExportCustomXmlMapping {
      Objects.requireNonNull(mapping, "mapping must not be null");
      encoding = TextValidation.requireNonBlank(encoding, "encoding");
    }
  }

  @Variant(
      id = "GET_NAMED_RANGES",
      summary = "Return named ranges matched by the supplied selection.",
      targetSelectors = {NamedRangeSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record GetNamedRanges() implements WorkbookIntrospectionQuery {}
}
