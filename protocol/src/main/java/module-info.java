module dev.erst.gridgrind.protocol {
  requires dev.erst.gridgrind.engine;
  requires tools.jackson.databind;

  exports dev.erst.gridgrind.protocol.catalog;
  exports dev.erst.gridgrind.protocol.dto;
  exports dev.erst.gridgrind.protocol.exec;
  exports dev.erst.gridgrind.protocol.json;
  exports dev.erst.gridgrind.protocol.operation;
  exports dev.erst.gridgrind.protocol.read;

  opens dev.erst.gridgrind.protocol.catalog to
      tools.jackson.databind;
  opens dev.erst.gridgrind.protocol.dto to
      tools.jackson.databind;
  opens dev.erst.gridgrind.protocol.operation to
      tools.jackson.databind;
  opens dev.erst.gridgrind.protocol.read to
      tools.jackson.databind;
}
