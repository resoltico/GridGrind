open module dev.erst.gridgrind.architecture.tests {
  requires com.tngtech.archunit;
  requires com.tngtech.archunit.junit6.api;
  requires dev.erst.gridgrind.authoring;
  requires dev.erst.gridgrind.cli;
  requires dev.erst.gridgrind.contract;
  requires dev.erst.gridgrind.engine;
  requires dev.erst.gridgrind.excel.foundation;

  exports dev.erst.gridgrind.architecture;
}
