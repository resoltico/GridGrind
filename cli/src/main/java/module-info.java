module dev.erst.gridgrind.cli {
  requires dev.erst.gridgrind.contract;
  requires dev.erst.gridgrind.engine;
  requires dev.erst.gridgrind.excel.foundation;
  requires static org.jspecify;
  requires tools.jackson.databind;

  exports dev.erst.gridgrind.cli;
  exports dev.erst.gridgrind.cli.discovery;

  opens dev.erst.gridgrind.cli.discovery to
      tools.jackson.databind;
}
