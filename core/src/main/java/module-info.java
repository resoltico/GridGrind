module dev.erst.gridgrind.core {
  requires static com.fasterxml.jackson.annotation;
  requires static org.jspecify;

  exports dev.erst.gridgrind.core.authoring;
  exports dev.erst.gridgrind.core.catalog;
  exports dev.erst.gridgrind.core.catalog.gather;
  exports dev.erst.gridgrind.core.metadata;
  exports dev.erst.gridgrind.core.model.chart;
  exports dev.erst.gridgrind.core.model.customxml;
  exports dev.erst.gridgrind.core.model.drawing;
  exports dev.erst.gridgrind.core.model.formula;
  exports dev.erst.gridgrind.core.model.layout;
  exports dev.erst.gridgrind.core.model.pivot;
  exports dev.erst.gridgrind.core.model.print;
  exports dev.erst.gridgrind.core.model.protection;
  exports dev.erst.gridgrind.core.model.security;
  exports dev.erst.gridgrind.core.model.style;
  exports dev.erst.gridgrind.core.model.table;
  exports dev.erst.gridgrind.core.model.validation;
  exports dev.erst.gridgrind.core.model.value;
  exports dev.erst.gridgrind.core.plan;
  exports dev.erst.gridgrind.core.problem;
  exports dev.erst.gridgrind.core.result;
  exports dev.erst.gridgrind.core.semantics;
  exports dev.erst.gridgrind.core.source;
  exports dev.erst.gridgrind.core.validation;
  exports dev.erst.gridgrind.core.vocabulary;

  opens dev.erst.gridgrind.core.catalog to
      tools.jackson.databind;
  opens dev.erst.gridgrind.core.catalog.gather to
      tools.jackson.databind;
  opens dev.erst.gridgrind.core.model.chart to
      tools.jackson.databind;
  opens dev.erst.gridgrind.core.model.customxml to
      tools.jackson.databind;
  opens dev.erst.gridgrind.core.model.drawing to
      tools.jackson.databind;
  opens dev.erst.gridgrind.core.model.formula to
      tools.jackson.databind;
  opens dev.erst.gridgrind.core.model.layout to
      tools.jackson.databind;
  opens dev.erst.gridgrind.core.model.pivot to
      tools.jackson.databind;
  opens dev.erst.gridgrind.core.model.print to
      tools.jackson.databind;
  opens dev.erst.gridgrind.core.model.protection to
      tools.jackson.databind;
  opens dev.erst.gridgrind.core.model.security to
      tools.jackson.databind;
  opens dev.erst.gridgrind.core.model.style to
      tools.jackson.databind;
  opens dev.erst.gridgrind.core.model.table to
      tools.jackson.databind;
  opens dev.erst.gridgrind.core.model.validation to
      tools.jackson.databind;
  opens dev.erst.gridgrind.core.model.value to
      tools.jackson.databind;
  opens dev.erst.gridgrind.core.plan to
      tools.jackson.databind;
  opens dev.erst.gridgrind.core.problem to
      tools.jackson.databind;
  opens dev.erst.gridgrind.core.result to
      tools.jackson.databind;
  opens dev.erst.gridgrind.core.semantics to
      tools.jackson.databind;
  opens dev.erst.gridgrind.core.source to
      tools.jackson.databind;
  opens dev.erst.gridgrind.core.validation to
      tools.jackson.databind;
}
