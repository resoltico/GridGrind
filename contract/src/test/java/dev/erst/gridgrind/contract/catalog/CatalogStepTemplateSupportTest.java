package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.node.ObjectNode;

/**
 * Focused behavior coverage for executable step-template synthesis inside the published catalog.
 */
class CatalogStepTemplateSupportTest {
  @Test
  void attachBuildsExecutableTemplatesWithPreferredTargetsAndTypedPlaceholders() {
    Catalog attached = CatalogStepTemplateSupport.attach(richFixtureCatalog());
    TypeEntry action = attached.mutationActionTypes().getFirst();
    ProtocolStepTemplate template = action.stepTemplate().orElseThrow();
    ObjectNode step = template.template();
    ObjectNode target = (ObjectNode) step.path("target");
    ObjectNode actionBody = (ObjectNode) step.path("action");

    assertEquals("MUTATION", template.stepKind());
    assertEquals("action", template.bodyField());
    assertEquals("step-custom-action", step.path("stepId").stringValue());
    assertEquals("CELL_BY_ADDRESS", target.path("type").stringValue());
    assertEquals("A1", target.path("address").stringValue());
    assertEquals("CUSTOM_ACTION", actionBody.path("type").stringValue());
    assertEquals("Sheet1", actionBody.path("sheetName").stringValue());
    assertEquals("2", actionBody.path("formula2").stringValue());
    assertEquals("sample-path", actionBody.path("path").stringValue());
    assertEquals("certificate.p12", actionBody.path("pkcs12Path").stringValue());
    assertEquals("signer@example.com", actionBody.path("email").stringValue());
    assertEquals("rId1", actionBody.path("relationshipId").stringValue());
    assertEquals("Sample title", actionBody.path("title").stringValue());
    assertEquals("Sample description", actionBody.path("description").stringValue());
    assertEquals("sample.bin", actionBody.path("fileName").stringValue());
    assertEquals("sample-custom-name", actionBody.path("custom_name").stringValue());
    assertEquals("sample-plan", actionBody.path("planId").stringValue());
    assertEquals("sample-step", actionBody.path("stepId").stringValue());
    assertEquals(0, actionBody.path("rowIndex").intValue());
    assertEquals(1, actionBody.path("count").intValue());
    assertTrue(actionBody.path("visible").booleanValue());
    assertFalse(actionBody.path("customFlag").booleanValue());
    assertEquals("TEXT", actionBody.path("source").path("type").stringValue());
    assertEquals("SAVE_AS", actionBody.path("persistence").path("type").stringValue());
    assertEquals("MUTATION", actionBody.path("stepPrototype").path("type").stringValue());
    assertEquals(
        "WORKBOOK_TARGET_ACTION", actionBody.path("actionPrototype").path("type").stringValue());
    assertEquals(
        "ASSERTION_ALPHA", actionBody.path("assertionPrototype").path("type").stringValue());
    assertEquals("QUERY_ALPHA", actionBody.path("queryPrototype").path("type").stringValue());
    assertFalse(actionBody.path("plainBlock").has("type"));
    assertEquals("Sample label", actionBody.path("plainBlock").path("label").stringValue());
    assertEquals("RECURSIVE", actionBody.path("recursive").path("self").path("kind").stringValue());
    assertEquals(
        "UNION_RECURSIVE_ALT",
        actionBody.path("recursiveUnion").path("self").path("kind").stringValue());
    assertEquals("INLINE_TEXT", actionBody.path("nested").path("kind").stringValue());
    assertEquals("Sample text", actionBody.path("nested").path("text").stringValue());
    assertEquals("ENUM_ALPHA", actionBody.path("enumChoice").stringValue());
    assertEquals("sample-values", actionBody.path("values").path(0).stringValue());
    assertTrue(
        template
            .notes()
            .contains(
                "Replace the placeholder target when needed; CUSTOM_ACTION accepts CELL_BY_ADDRESSES, CELL_BY_ADDRESS, CUSTOM_TARGET."));
    assertTrue(
        template
            .notes()
            .contains(
                "Replace the recursive placeholder for RECURSIVE with one concrete non-recursive variant."));
  }

  @Test
  void differentialStyleTemplateDefaultProducesAnExecutableStylePlaceholder() {
    ObjectNode style = tools.jackson.databind.node.JsonNodeFactory.instance.objectNode();
    CatalogStepTemplateDefaults.buildTypeSpecificDefaults()
        .get("DifferentialStyleInput")
        .apply(
            GridGrindProtocolCatalog.catalog(),
            style,
            new java.util.HashSet<>(),
            new java.util.ArrayList<>());

    assertTrue(style.path("bold").booleanValue());
  }

  @Test
  void attachUsesTargetRulesAndRejectsBrokenCatalogReferences() {
    Catalog targetRuleCatalog = CatalogStepTemplateSupport.attach(targetRuleCatalog());
    ProtocolStepTemplate targetRuleTemplate =
        targetRuleCatalog.assertionTypes().getFirst().stepTemplate().orElseThrow();

    assertEquals(
        "WORKBOOK_CURRENT",
        targetRuleTemplate.template().path("target").path("type").stringValue());
    assertTrue(targetRuleTemplate.notes().contains("Target rule: Matches the nested target."));

    assertEquals(
        "Catalog target selector ids did not resolve for BROKEN_TARGET",
        assertThrows(
                IllegalStateException.class,
                () -> CatalogStepTemplateSupport.attach(brokenTargetCatalog()))
            .getMessage());
    assertEquals(
        "Catalog is missing WORKBOOK_CURRENT for dynamic target placeholder",
        assertThrows(
                IllegalStateException.class,
                () -> CatalogStepTemplateSupport.attach(missingWorkbookCurrentCatalog()))
            .getMessage());
    assertEquals(
        "Unknown nested type group: missingNested",
        assertThrows(
                IllegalArgumentException.class,
                () -> CatalogStepTemplateSupport.attach(missingNestedGroupCatalog()))
            .getMessage());
    assertEquals(
        "Unknown plain type group: missingPlain",
        assertThrows(
                IllegalArgumentException.class,
                () -> CatalogStepTemplateSupport.attach(missingPlainGroupCatalog()))
            .getMessage());
    assertEquals(
        "Unsupported top-level type set: bogusTypes",
        assertThrows(
                IllegalArgumentException.class,
                () -> CatalogStepTemplateSupport.attach(bogusTopLevelTypeSetCatalog()))
            .getMessage());
    assertEquals(
        "Catalog group emptyGroup did not expose any entries",
        assertThrows(
                IllegalStateException.class,
                () -> CatalogStepTemplateSupport.attach(emptyNestedGroupCatalog()))
            .getMessage());
    assertEquals(
        "Catalog is missing type ANALYZE_WORKBOOK_FINDINGS",
        assertThrows(
                IllegalStateException.class,
                () -> CatalogStepTemplateSupport.attach(missingTypeSpecificDefaultCatalog()))
            .getMessage());
  }

  @Test
  void attachPrefersWorkbookCurrentTargetsWhenAvailable() {
    Catalog attached = CatalogStepTemplateSupport.attach(richFixtureCatalog());
    ProtocolStepTemplate workbookTargetTemplate =
        attached.mutationActionTypes().get(1).stepTemplate().orElseThrow();

    assertEquals(
        "WORKBOOK_CURRENT",
        workbookTargetTemplate.template().path("target").path("type").stringValue());
  }

  private static Catalog richFixtureCatalog() {
    TypeEntry workbookCurrent =
        new TypeEntry(
            "WORKBOOK_CURRENT",
            "Current workbook target.",
            List.of(field("name", new FieldShape.Scalar(ScalarType.STRING))));
    TypeEntry cellByAddress =
        new TypeEntry(
            "CELL_BY_ADDRESS",
            "Cell target by address.",
            List.of(field("address", new FieldShape.Scalar(ScalarType.STRING))));
    TypeEntry cellByAddresses =
        new TypeEntry(
            "CELL_BY_ADDRESSES",
            "Multiple cell targets.",
            List.of(
                field(
                    "addresses",
                    new FieldShape.ListShape(new FieldShape.Scalar(ScalarType.STRING)))));
    TypeEntry customTarget =
        new TypeEntry(
            "CUSTOM_TARGET",
            "Low-priority custom target.",
            List.of(field("name", new FieldShape.Scalar(ScalarType.STRING))));
    TypeEntry nestedValue =
        new TypeEntry(
            "INLINE_TEXT",
            "Inline nested value.",
            List.of(field("text", new FieldShape.Scalar(ScalarType.STRING))));
    TypeEntry recursive =
        new TypeEntry(
            "RECURSIVE",
            "Recursive placeholder.",
            List.of(
                field("self", new FieldShape.NestedTypeGroupUnionRef(List.of("recursiveGroup")))));
    TypeEntry plainLabel =
        new TypeEntry(
            "PLAIN_LABEL",
            "Plain label block.",
            List.of(field("label", new FieldShape.Scalar(ScalarType.STRING))));
    TypeEntry sourceText =
        new TypeEntry(
            "TEXT",
            "Text source.",
            List.of(field("path", new FieldShape.Scalar(ScalarType.STRING))));
    TypeEntry sourceCustom =
        new TypeEntry(
            "CUSTOM_SOURCE",
            "Fallback source.",
            List.of(field("path", new FieldShape.Scalar(ScalarType.STRING))));
    TypeEntry persistenceSaveAs =
        new TypeEntry(
            "SAVE_AS",
            "Save workbook as a file.",
            List.of(field("path", new FieldShape.Scalar(ScalarType.STRING))));
    TypeEntry stepMutation =
        new TypeEntry(
            "MUTATION",
            "Mutation step wrapper.",
            List.of(field("stepId", new FieldShape.Scalar(ScalarType.STRING))));
    TypeEntry assertionAlpha =
        new TypeEntry(
            "ASSERTION_ALPHA",
            "Assertion wrapper.",
            List.of(field("label", new FieldShape.Scalar(ScalarType.STRING))));
    TypeEntry queryAlpha =
        new TypeEntry(
            "QUERY_ALPHA",
            "Inspection wrapper.",
            List.of(field("label", new FieldShape.Scalar(ScalarType.STRING))));
    TypeEntry action =
        new TypeEntry(
            "CUSTOM_ACTION",
            "Synthetic action for step-template coverage.",
            List.of(
                field("sheetName", new FieldShape.Scalar(ScalarType.STRING)),
                field("formula2", new FieldShape.Scalar(ScalarType.STRING)),
                field("path", new FieldShape.Scalar(ScalarType.STRING)),
                field("pkcs12Path", new FieldShape.Scalar(ScalarType.STRING)),
                field("email", new FieldShape.Scalar(ScalarType.STRING)),
                field("relationshipId", new FieldShape.Scalar(ScalarType.STRING)),
                field("title", new FieldShape.Scalar(ScalarType.STRING)),
                field("description", new FieldShape.Scalar(ScalarType.STRING)),
                field("fileName", new FieldShape.Scalar(ScalarType.STRING)),
                field("custom_name", new FieldShape.Scalar(ScalarType.STRING)),
                field("planId", new FieldShape.Scalar(ScalarType.STRING)),
                field("stepId", new FieldShape.Scalar(ScalarType.STRING)),
                field("rowIndex", new FieldShape.Scalar(ScalarType.NUMBER)),
                field("count", new FieldShape.Scalar(ScalarType.NUMBER)),
                field("visible", new FieldShape.Scalar(ScalarType.BOOLEAN)),
                field("customFlag", new FieldShape.Scalar(ScalarType.BOOLEAN)),
                field("source", new FieldShape.TopLevelTypeSetRef("sourceTypes")),
                field("persistence", new FieldShape.TopLevelTypeSetRef("persistenceTypes")),
                field("stepPrototype", new FieldShape.TopLevelTypeSetRef("stepTypes")),
                field("actionPrototype", new FieldShape.TopLevelTypeSetRef("mutationActionTypes")),
                field("assertionPrototype", new FieldShape.TopLevelTypeSetRef("assertionTypes")),
                field("queryPrototype", new FieldShape.TopLevelTypeSetRef("inspectionQueryTypes")),
                field("plainBlock", new FieldShape.PlainTypeGroupRef("plainBlocks")),
                field("nested", new FieldShape.NestedTypeGroupRef("nestedValues")),
                field("recursive", new FieldShape.NestedTypeGroupRef("recursiveGroup")),
                field(
                    "recursiveUnion",
                    new FieldShape.NestedTypeGroupUnionRef(List.of("recursiveUnionGroup"))),
                new FieldEntry(
                    "enumChoice",
                    FieldRequirement.REQUIRED,
                    new FieldShape.Scalar(ScalarType.STRING),
                    List.of("ENUM_ALPHA", "ENUM_BETA")),
                field(
                    "values", new FieldShape.ListShape(new FieldShape.Scalar(ScalarType.STRING)))),
            List.of(
                new TargetSelectorEntry(
                    "cells", List.of("CELL_BY_ADDRESSES", "CELL_BY_ADDRESS", "CUSTOM_TARGET"))),
            Optional.empty());
    TypeEntry workbookTargetAction =
        new TypeEntry(
            "WORKBOOK_TARGET_ACTION",
            "Synthetic action preferring workbook-current targeting.",
            List.of(field("label", new FieldShape.Scalar(ScalarType.STRING))),
            List.of(
                new TargetSelectorEntry(
                    "workbook", List.of("CELL_BY_ADDRESS", "WORKBOOK_CURRENT"))),
            Optional.empty());
    return catalog(
        List.of(sourceText, sourceCustom),
        List.of(persistenceSaveAs),
        List.of(stepMutation),
        List.of(action, workbookTargetAction),
        List.of(assertionAlpha),
        List.of(queryAlpha),
        List.of(
            new NestedTypeGroup(
                "cellSelectorTypes",
                "type",
                List.of(workbookCurrent, cellByAddress, cellByAddresses, customTarget)),
            new NestedTypeGroup("nestedValues", "kind", List.of(nestedValue)),
            new NestedTypeGroup("recursiveGroup", "kind", List.of(recursive)),
            new NestedTypeGroup(
                "recursiveUnionGroup", "kind", List.of(unionRecursive(), unionRecursiveAlt()))),
        List.of(new PlainTypeGroup("plainBlocks", plainLabel)));
  }

  private static Catalog targetRuleCatalog() {
    TypeEntry workbookCurrent =
        new TypeEntry(
            "WORKBOOK_CURRENT",
            "Current workbook target.",
            List.of(field("name", new FieldShape.Scalar(ScalarType.STRING))));
    TypeEntry assertion =
        new TypeEntry(
            "TARGET_RULE_ASSERTION",
            "Assertion using a target rule.",
            List.of(field("label", new FieldShape.Scalar(ScalarType.STRING))),
            List.of(),
            Optional.of("Matches the nested target."));
    return catalog(
        List.of(),
        List.of(),
        List.of(),
        List.of(),
        List.of(assertion),
        List.of(),
        List.of(new NestedTypeGroup("workbookSelectorTypes", "type", List.of(workbookCurrent))),
        List.of());
  }

  private static Catalog brokenTargetCatalog() {
    TypeEntry action =
        new TypeEntry(
            "BROKEN_TARGET",
            "Broken target selector ids.",
            List.of(field("label", new FieldShape.Scalar(ScalarType.STRING))),
            List.of(new TargetSelectorEntry("cells", List.of("MISSING_TARGET"))),
            Optional.empty());
    return catalog(
        List.of(),
        List.of(),
        List.of(),
        List.of(action),
        List.of(),
        List.of(),
        List.of(),
        List.of());
  }

  private static Catalog missingWorkbookCurrentCatalog() {
    TypeEntry assertion =
        new TypeEntry(
            "NO_WORKBOOK_CURRENT",
            "Missing fallback workbook selector.",
            List.of(field("label", new FieldShape.Scalar(ScalarType.STRING))),
            List.of(),
            Optional.empty());
    return catalog(
        List.of(),
        List.of(),
        List.of(),
        List.of(),
        List.of(assertion),
        List.of(),
        List.of(),
        List.of());
  }

  private static Catalog missingNestedGroupCatalog() {
    TypeEntry action =
        new TypeEntry(
            "MISSING_NESTED_GROUP",
            "Missing nested group.",
            List.of(field("nested", new FieldShape.NestedTypeGroupRef("missingNested"))),
            List.of(new TargetSelectorEntry("workbook", List.of("WORKBOOK_CURRENT"))),
            Optional.empty());
    return catalog(
        List.of(),
        List.of(),
        List.of(),
        List.of(action),
        List.of(),
        List.of(),
        List.of(
            new NestedTypeGroup(
                "workbookSelectorTypes",
                "type",
                List.of(
                    new TypeEntry(
                        "WORKBOOK_CURRENT",
                        "Workbook target.",
                        List.of(field("name", new FieldShape.Scalar(ScalarType.STRING))))))),
        List.of());
  }

  private static Catalog missingPlainGroupCatalog() {
    TypeEntry action =
        new TypeEntry(
            "MISSING_PLAIN_GROUP",
            "Missing plain group.",
            List.of(field("plain", new FieldShape.PlainTypeGroupRef("missingPlain"))),
            List.of(new TargetSelectorEntry("workbook", List.of("WORKBOOK_CURRENT"))),
            Optional.empty());
    return catalog(
        List.of(),
        List.of(),
        List.of(),
        List.of(action),
        List.of(),
        List.of(),
        List.of(
            new NestedTypeGroup(
                "workbookSelectorTypes",
                "type",
                List.of(
                    new TypeEntry(
                        "WORKBOOK_CURRENT",
                        "Workbook target.",
                        List.of(field("name", new FieldShape.Scalar(ScalarType.STRING))))))),
        List.of());
  }

  private static Catalog bogusTopLevelTypeSetCatalog() {
    TypeEntry action =
        new TypeEntry(
            "BOGUS_TOP_LEVEL",
            "Unsupported top-level type set.",
            List.of(field("source", new FieldShape.TopLevelTypeSetRef("bogusTypes"))),
            List.of(new TargetSelectorEntry("workbook", List.of("WORKBOOK_CURRENT"))),
            Optional.empty());
    return catalog(
        List.of(),
        List.of(),
        List.of(),
        List.of(action),
        List.of(),
        List.of(),
        List.of(
            new NestedTypeGroup(
                "workbookSelectorTypes",
                "type",
                List.of(
                    new TypeEntry(
                        "WORKBOOK_CURRENT",
                        "Workbook target.",
                        List.of(field("name", new FieldShape.Scalar(ScalarType.STRING))))))),
        List.of());
  }

  private static Catalog emptyNestedGroupCatalog() {
    TypeEntry action =
        new TypeEntry(
            "EMPTY_NESTED_GROUP",
            "Nested group exists but is empty.",
            List.of(field("nested", new FieldShape.NestedTypeGroupRef("emptyGroup"))),
            List.of(new TargetSelectorEntry("workbook", List.of("WORKBOOK_CURRENT"))),
            Optional.empty());
    return catalog(
        List.of(),
        List.of(),
        List.of(),
        List.of(action),
        List.of(),
        List.of(),
        List.of(
            new NestedTypeGroup(
                "workbookSelectorTypes",
                "type",
                List.of(
                    new TypeEntry(
                        "WORKBOOK_CURRENT",
                        "Workbook target.",
                        List.of(field("name", new FieldShape.Scalar(ScalarType.STRING)))))),
            new NestedTypeGroup("emptyGroup", "type", List.of())),
        List.of());
  }

  private static Catalog missingTypeSpecificDefaultCatalog() {
    TypeEntry workbookCurrent =
        new TypeEntry(
            "WORKBOOK_CURRENT",
            "Current workbook target.",
            List.of(field("name", new FieldShape.Scalar(ScalarType.STRING))));
    TypeEntry assertion =
        new TypeEntry(
            "EXPECT_ANALYSIS_MAX_SEVERITY",
            "Assertion whose built-in template references a missing query type.",
            List.of(field("severity", new FieldShape.Scalar(ScalarType.STRING))),
            List.of(new TargetSelectorEntry("workbook", List.of("WORKBOOK_CURRENT"))),
            Optional.empty());
    return catalog(
        List.of(),
        List.of(),
        List.of(),
        List.of(),
        List.of(assertion),
        List.of(),
        List.of(new NestedTypeGroup("workbookSelectorTypes", "type", List.of(workbookCurrent))),
        List.of());
  }

  private static Catalog catalog(
      List<TypeEntry> sourceTypes,
      List<TypeEntry> persistenceTypes,
      List<TypeEntry> stepTypes,
      List<TypeEntry> mutationActionTypes,
      List<TypeEntry> assertionTypes,
      List<TypeEntry> inspectionQueryTypes,
      List<NestedTypeGroup> nestedTypes,
      List<PlainTypeGroup> plainTypes) {
    return new Catalog(
        GridGrindProtocolVersion.current(),
        "type",
        new TypeEntry("REQUEST", "Synthetic request.", List.of()),
        sourceTypes,
        persistenceTypes,
        stepTypes,
        mutationActionTypes,
        assertionTypes,
        inspectionQueryTypes,
        nestedTypes,
        plainTypes,
        List.of());
  }

  private static TypeEntry unionRecursive() {
    return new TypeEntry(
        "UNION_RECURSIVE",
        "Recursive union placeholder.",
        List.of(
            field("self", new FieldShape.NestedTypeGroupUnionRef(List.of("recursiveUnionGroup")))));
  }

  private static TypeEntry unionRecursiveAlt() {
    return new TypeEntry(
        "UNION_RECURSIVE_ALT",
        "Recursive union alternate placeholder.",
        List.of(
            field("self", new FieldShape.NestedTypeGroupUnionRef(List.of("recursiveUnionGroup")))));
  }

  private static FieldEntry field(String name, FieldShape shape) {
    return new FieldEntry(name, FieldRequirement.REQUIRED, shape, List.of());
  }
}
