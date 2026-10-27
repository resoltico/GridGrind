package dev.erst.gridgrind.json.codec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionSpec;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Dynamic filter domain values must remain distinct from the union discriminator. */
class DynamicAutofilterCodecTest {
  @Test
  void authoredDynamicFilterRoundTripsItsDomainType() {
    var mapper = GridGrindJsonMapperSupport.WIRE_JSON_MAPPER;
    var value =
        new AutofilterFilterCriterionSpec.Dynamic("today", Optional.of(45_200d), Optional.empty());
    var wire = mapper.writerFor(AutofilterFilterCriterionSpec.class).writeValueAsString(value);
    var tree = mapper.readTree(wire);
    assertEquals("DYNAMIC", tree.path("type").asText());
    assertEquals("today", tree.path("dynamicType").asText());
    assertEquals(value, mapper.readValue(wire, AutofilterFilterCriterionSpec.class));
  }

  @Test
  void observedDynamicFilterRoundTripsItsDomainType() {
    var mapper = GridGrindJsonMapperSupport.WIRE_JSON_MAPPER;
    var value =
        new AutofilterFilterCriterionReport.Dynamic(
            "today", Optional.of(45_200d), Optional.empty());
    var wire = mapper.writerFor(AutofilterFilterCriterionReport.class).writeValueAsString(value);
    var tree = mapper.readTree(wire);
    assertEquals("DYNAMIC", tree.path("type").asText());
    assertEquals("today", tree.path("dynamicType").asText());
    assertEquals(value, mapper.readValue(wire, AutofilterFilterCriterionReport.class));
  }

  @Test
  void schemaDrivenRequestIntakeBindsTheDynamicDomainField() {
    var plan =
        GridGrindJson.readRequest(
            """
        {"protocolVersion":1,"source":{"type":"NEW"},
         "persistence":{"type":"SAVE_AS","path":"out.xlsx","ifExists":"REPLACE"},
         "steps":[{"stepId":"filter","target":{"type":"RANGE_BY_RANGE","sheetName":"Data","range":"A1:A3"},
           "action":{"type":"SET_AUTOFILTER","criteria":[{"columnId":0,
             "criterion":{"type":"DYNAMIC","dynamicType":"today"}}]}}]}
        """);
    var step = assertInstanceOf(MutationStep.class, plan.steps().getFirst());
    var action = assertInstanceOf(StructuredMutationAction.SetAutofilter.class, step.action());
    var criterion =
        assertInstanceOf(
            AutofilterFilterCriterionSpec.Dynamic.class, action.criteria().getFirst().criterion());
    assertEquals("today", criterion.dynamicType());
  }
}
