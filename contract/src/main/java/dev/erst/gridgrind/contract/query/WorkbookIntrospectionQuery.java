package dev.erst.gridgrind.contract.query;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadata;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
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

  @ProtocolTypeMetadata(
      id = "GET_WORKBOOK_SUMMARY",
      summary = "Return workbook-level summary facts.",
      targetSelectors = {WorkbookSelector.class})
  record GetWorkbookSummary() implements WorkbookIntrospectionQuery {}

  @ProtocolTypeMetadata(
      id = "GET_PACKAGE_SECURITY",
      summary = "Return OOXML package-encryption and package-signature facts.",
      targetSelectors = {WorkbookSelector.class})
  record GetPackageSecurity() implements WorkbookIntrospectionQuery {}

  @ProtocolTypeMetadata(
      id = "GET_WORKBOOK_PROTECTION",
      summary = "Return workbook-level protection facts.",
      targetSelectors = {WorkbookSelector.class})
  record GetWorkbookProtection() implements WorkbookIntrospectionQuery {}

  @ProtocolTypeMetadata(
      id = "GET_CUSTOM_XML_MAPPINGS",
      summary = "Return workbook custom-XML mapping metadata.",
      targetSelectors = {WorkbookSelector.class})
  record GetCustomXmlMappings() implements WorkbookIntrospectionQuery {}

  @ProtocolTypeMetadata(
      id = "EXPORT_CUSTOM_XML_MAPPING",
      summary = "Export one existing workbook custom-XML mapping as serialized XML.",
      optionalFields = {"validateSchema", "encoding"},
      targetSelectors = {WorkbookSelector.class})
  record ExportCustomXmlMapping(
      CustomXmlMappingLocator mapping, boolean validateSchema, String encoding)
      implements WorkbookIntrospectionQuery {
    /** Creates one UTF-8 export request without requiring the caller to pass the encoding. */
    public ExportCustomXmlMapping(CustomXmlMappingLocator mapping, boolean validateSchema) {
      this(mapping, validateSchema, StandardCharsets.UTF_8.name());
    }

    /** Reads one export request while defaulting omitted flags and encoding. */
    @JsonCreator
    public ExportCustomXmlMapping(
        @JsonProperty("mapping") CustomXmlMappingLocator mapping,
        @JsonProperty("validateSchema") Boolean validateSchema,
        @JsonProperty("encoding") String encoding) {
      this(
          mapping,
          Boolean.TRUE.equals(validateSchema),
          encoding == null ? StandardCharsets.UTF_8.name() : encoding);
    }

    public ExportCustomXmlMapping {
      Objects.requireNonNull(mapping, "mapping must not be null");
      Objects.requireNonNull(encoding, "encoding must not be null");
      if (encoding.isBlank()) {
        throw new IllegalArgumentException("encoding must not be blank");
      }
    }
  }

  @ProtocolTypeMetadata(
      id = "GET_NAMED_RANGES",
      summary = "Return named ranges matched by the supplied selection.",
      targetSelectors = {NamedRangeSelector.class})
  record GetNamedRanges() implements WorkbookIntrospectionQuery {}
}
