package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.ModelGraph;
import dev.erst.gridgrind.core.metadata.ProjectedBy;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.plan.Assertion;
import dev.erst.gridgrind.core.plan.InspectionQuery;
import dev.erst.gridgrind.core.plan.MutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.result.ExecutionProgressEvent;
import dev.erst.gridgrind.core.result.WorkbookResult;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;

/** Derives catalog membership and descriptors from the public model graph. */
final class CatalogTypes {
  private static final List<Class<?>> MODEL_TYPES =
      ModelGraph.reachableTypes(
          WorkbookPlan.class, WorkbookResult.class, Catalog.class, ExecutionProgressEvent.class);
  static final List<CatalogTypeDescriptor> SOURCE_TYPES =
      descriptors(WorkbookPlan.WorkbookSource.class);
  static final List<CatalogTypeDescriptor> PERSISTENCE_TYPES =
      descriptors(WorkbookPlan.WorkbookPersistence.class);
  static final List<CatalogTypeDescriptor> STEP_TYPES = descriptors(WorkbookStep.class);
  static final List<CatalogTypeDescriptor> MUTATION_ACTION_TYPES =
      descriptors(MutationAction.class);
  static final List<CatalogTypeDescriptor> ASSERTION_TYPES = descriptors(Assertion.class);
  static final List<CatalogTypeDescriptor> INSPECTION_QUERY_TYPES =
      descriptors(InspectionQuery.class);
  static final List<CatalogTopLevelTypeDescriptorGroup> TOP_LEVEL_GROUPS =
      List.of(
          topLevel(WorkbookPlan.WorkbookSource.class),
          topLevel(WorkbookPlan.WorkbookPersistence.class),
          topLevel(WorkbookStep.class),
          topLevel(MutationAction.class),
          topLevel(Assertion.class),
          topLevel(InspectionQuery.class));
  static final List<CatalogTypeDescriptor> ALL_TYPES =
      TOP_LEVEL_GROUPS.stream().flatMap(group -> group.typeDescriptors().stream()).toList();
  static final List<CatalogNestedTypeDescriptor> NESTED_TYPE_GROUPS =
      MODEL_TYPES.stream()
          .filter(
              type ->
                  type.isAnnotationPresent(Union.class)
                      && type.getAnnotation(Union.class).catalog())
          .map(
              type ->
                  new CatalogNestedTypeDescriptor(
                      groupId(type),
                      type.getAnnotation(Union.class).discriminator(),
                      type,
                      descriptors(type)))
          .sorted(
              Comparator.comparingInt(
                  group -> group.sealedType().getAnnotation(Union.class).displayOrder()))
          .toList();
  static final List<CatalogPlainTypeDescriptor> PLAIN_TYPE_DESCRIPTORS =
      MODEL_TYPES.stream()
          .filter(
              type ->
                  type.isRecord()
                      && type != WorkbookPlan.class
                      && type.isAnnotationPresent(CatalogType.class))
          .map(
              type -> {
                CatalogType metadata = type.getAnnotation(CatalogType.class);
                return new CatalogPlainTypeDescriptor(
                    plainGroupId(type),
                    type.asSubclass(Record.class),
                    metadata.id(),
                    metadata.summary(),
                    List.of(metadata.noteRefs()));
              })
          .sorted(
              Comparator.comparingInt(
                  group -> group.recordType().getAnnotation(CatalogType.class).displayOrder()))
          .toList();

  private CatalogTypes() {}

  static String groupId(Class<?> union) {
    return decapitalize(union.getSimpleName()) + "Types";
  }

  static String plainGroupId(Class<?> recordType) {
    return decapitalize(recordType.getAnnotation(CatalogType.class).id()) + "Type";
  }

  static CatalogTypeDescriptor descriptor(Class<?> recordType) {
    Variant metadata = ProtocolTypeMetadataSupport.requiredMetadata(recordType);
    return new CatalogTypeDescriptor(
        recordType.asSubclass(Record.class),
        metadata.id(),
        ProtocolTypeMetadataSupport.summary(recordType),
        List.of(metadata.noteRefs()),
        projections(recordType));
  }

  static List<CatalogTypeDescriptor> descriptors(Class<?> union) {
    return ProtocolTypeMetadataSupport.catalogDescriptorsFor(union);
  }

  static List<CatalogProjectedField> projections(Class<?> recordType) {
    return Arrays.stream(recordType.getRecordComponents())
        .filter(component -> component.isAnnotationPresent(ProjectedBy.class))
        .map(
            component ->
                new CatalogProjectedField(
                    component.getName(),
                    List.of(component.getAnnotation(ProjectedBy.class).value())))
        .toList();
  }

  private static CatalogTopLevelTypeDescriptorGroup topLevel(Class<?> union) {
    return new CatalogTopLevelTypeDescriptorGroup(groupId(union), union, descriptors(union));
  }

  private static String decapitalize(String value) {
    return Character.toLowerCase(value.charAt(0)) + value.substring(1);
  }
}
