package dev.erst.gridgrind.contract.json;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.catalog.GridGrindRequestSurfaceContractText;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import org.jspecify.annotations.Nullable;
import tools.jackson.core.StreamReadConstraints;
import tools.jackson.core.StreamReadFeature;
import tools.jackson.core.StreamWriteFeature;
import tools.jackson.core.json.JsonFactory;
import tools.jackson.core.json.JsonFactoryBuilder;
import tools.jackson.databind.DeserializationFeature;
import tools.jackson.databind.SerializationFeature;
import tools.jackson.databind.cfg.CoercionAction;
import tools.jackson.databind.cfg.CoercionInputShape;
import tools.jackson.databind.json.JsonMapper;
import tools.jackson.databind.type.LogicalType;

/** Owns JSON mapper construction and request-size policy for the protocol seam. */
final class GridGrindJsonMapperSupport {
  static final JsonMapper JSON_MAPPER = buildMapper(unlimitedJsonFactory());
  static final JsonMapper WIRE_JSON_MAPPER = buildMapper(unlimitedJsonFactory(), true, false);
  static final JsonMapper PRETTY_WIRE_JSON_MAPPER = buildMapper(unlimitedJsonFactory(), true, true);
  static final JsonMapper REQUEST_JSON_MAPPER = buildMapper(requestJsonFactory());

  private GridGrindJsonMapperSupport() {}

  static long maxRequestDocumentBytes() {
    return GridGrindRequestSurfaceContractText.requestDocumentLimitBytes();
  }

  static void requireSupportedRequestLength(long lengthBytes) {
    // LIM-021
    if (lengthBytes > maxRequestDocumentBytes()) {
      throw requestTooLarge(null);
    }
  }

  static InvalidRequestException requestTooLarge(@Nullable Throwable cause) {
    return new InvalidRequestException(
        new MessageInvariant(
            GridGrindRequestSurfaceContractText.requestDocumentTooLargeMessage(),
            java.util.Optional.empty()),
        java.util.Optional.empty(),
        java.util.Optional.empty(),
        java.util.Optional.empty(),
        cause);
  }

  private static JsonMapper buildMapper(JsonFactory jsonFactory) {
    return buildMapper(jsonFactory, false, false);
  }

  private static JsonMapper buildMapper(
      JsonFactory jsonFactory, boolean omitNullValues, boolean prettyPrint) {
    JsonMapper.Builder builder =
        JsonMapper.builder(jsonFactory)
            .disable(StreamReadFeature.AUTO_CLOSE_SOURCE)
            .disable(StreamWriteFeature.AUTO_CLOSE_TARGET)
            .enable(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES)
            .registerSubtypes(GridGrindJsonSubtypeSupport.namedLeafSubtypes(MutationAction.class))
            .registerSubtypes(GridGrindJsonSubtypeSupport.namedLeafSubtypes(Assertion.class))
            .registerSubtypes(GridGrindJsonSubtypeSupport.namedLeafSubtypes(InspectionQuery.class))
            .registerSubtypes(GridGrindJsonSubtypeSupport.namedLeafSubtypes(InspectionResult.class))
            .withCoercionConfig(
                LogicalType.Integer,
                config -> config.setCoercion(CoercionInputShape.Float, CoercionAction.Fail));
    if (prettyPrint) {
      builder.enable(SerializationFeature.INDENT_OUTPUT);
    }
    if (omitNullValues) {
      builder.changeDefaultPropertyInclusion(
          value -> value.withValueInclusion(JsonInclude.Include.NON_ABSENT));
    }
    return builder.build();
  }

  private static JsonFactory unlimitedJsonFactory() {
    return new JsonFactoryBuilder().build();
  }

  private static JsonFactory requestJsonFactory() {
    return new JsonFactoryBuilder()
        .streamReadConstraints(
            StreamReadConstraints.builder().maxDocumentLength(maxRequestDocumentBytes()).build())
        .build();
  }
}
