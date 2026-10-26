package dev.erst.gridgrind.json.codec;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.core.catalog.GridGrindRequestSurfaceContractText;
import dev.erst.gridgrind.core.plan.Assertion;
import dev.erst.gridgrind.core.plan.InspectionQuery;
import dev.erst.gridgrind.core.plan.InspectionResult;
import dev.erst.gridgrind.core.plan.MutationAction;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.problem.MessageInvariant;
import dev.erst.gridgrind.core.vocabulary.WireString;
import dev.erst.gridgrind.json.intake.WorkbookStepJsonDeserializer;
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
import tools.jackson.databind.module.SimpleModule;
import tools.jackson.databind.type.LogicalType;

/** Owns JSON mapper construction and request-size policy for the protocol seam. */
public final class GridGrindJsonMapperSupport {
  public static final JsonMapper JSON_MAPPER = buildMapper(unlimitedJsonFactory());
  public static final JsonMapper WIRE_JSON_MAPPER = buildMapper(unlimitedJsonFactory(), false);
  public static final JsonMapper PRETTY_WIRE_JSON_MAPPER =
      buildMapper(unlimitedJsonFactory(), true);
  public static final JsonMapper REQUEST_JSON_MAPPER = buildMapper(requestJsonFactory());

  private GridGrindJsonMapperSupport() {}

  static long maxRequestDocumentBytes() {
    return GridGrindRequestSurfaceContractText.requestDocumentLimitBytes();
  }

  /** Rejects a document size outside the request transport limit. */
  public static void requireSupportedRequestLength(long lengthBytes) {
    // LIM-021
    if (lengthBytes > maxRequestDocumentBytes()) {
      throw requestTooLarge(null);
    }
  }

  /** Classifies an exceeded request transport limit. */
  public static InvalidRequestException requestTooLarge(@Nullable Throwable cause) {
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
    return buildMapper(jsonFactory, false);
  }

  private static JsonMapper buildMapper(JsonFactory jsonFactory, boolean prettyPrint) {
    JsonMapper.Builder builder =
        JsonMapper.builder(jsonFactory)
            .annotationIntrospector(new CoreMetadataIntrospector())
            .enable(tools.jackson.databind.MapperFeature.INFER_RECORD_GETTERS_FROM_COMPONENTS_ONLY)
            .addModule(coreCodecs())
            .disable(StreamReadFeature.AUTO_CLOSE_SOURCE)
            .disable(StreamWriteFeature.AUTO_CLOSE_TARGET)
            .enable(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES)
            // RequestStructuralAnalyzer owns V3 field presence. The typed mapper only
            // normalizes already shape-valid fragments, including Optional fields.
            .disable(DeserializationFeature.FAIL_ON_MISSING_CREATOR_PROPERTIES)
            .disable(DeserializationFeature.FAIL_ON_NULL_CREATOR_PROPERTIES)
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
    builder.changeDefaultPropertyInclusion(
        value -> value.withValueInclusion(JsonInclude.Include.NON_ABSENT));
    return builder.build();
  }

  private static SimpleModule coreCodecs() {
    SimpleModule module =
        new SimpleModule()
            .setValueInstantiators(new RecordDefaultInstantiators())
            .setSerializerModifier(new RecordDefaultSerialization())
            .setDeserializers(new WireStringDeserializers())
            .addSerializer(WireString.class, new WireStringSerializer())
            .addSerializer(Selector.class, new SelectorJsonSerializer());
    registerStepCodec(module, WorkbookStep.class);
    for (Class<?> type : WorkbookStep.class.getPermittedSubclasses()) {
      registerStepCodec(module, type.asSubclass(WorkbookStep.class));
    }
    return module;
  }

  private static <T extends WorkbookStep> void registerStepCodec(
      SimpleModule module, Class<T> type) {
    module.addDeserializer(type, new WorkbookStepJsonDeserializer<>(type));
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
