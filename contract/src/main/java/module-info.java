module dev.erst.gridgrind.contract {
  requires transitive com.fasterxml.jackson.annotation;
  requires transitive dev.erst.gridgrind.excel.foundation;
  requires tools.jackson.databind;

  exports dev.erst.gridgrind.contract.assertion;
  exports dev.erst.gridgrind.contract.catalog;
  exports dev.erst.gridgrind.contract.action;
  exports dev.erst.gridgrind.contract.dto;
  exports dev.erst.gridgrind.contract.json;
  exports dev.erst.gridgrind.contract.query;
  exports dev.erst.gridgrind.contract.selector;
  exports dev.erst.gridgrind.contract.source;
  exports dev.erst.gridgrind.contract.step;

  opens dev.erst.gridgrind.contract.assertion to
      tools.jackson.databind;
  opens dev.erst.gridgrind.contract.catalog to
      tools.jackson.databind;
  opens dev.erst.gridgrind.contract.action to
      tools.jackson.databind;
  opens dev.erst.gridgrind.contract.dto to
      tools.jackson.databind;
  opens dev.erst.gridgrind.contract.query to
      tools.jackson.databind;
  opens dev.erst.gridgrind.contract.selector to
      tools.jackson.databind;
  opens dev.erst.gridgrind.contract.source to
      tools.jackson.databind;
  opens dev.erst.gridgrind.contract.step to
      tools.jackson.databind;
}
