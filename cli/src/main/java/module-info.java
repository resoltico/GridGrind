module dev.erst.gridgrind.cli {
  requires dev.erst.gridgrind.json;
  requires dev.erst.gridgrind.engine;
  requires static org.jspecify;
  requires tools.jackson.databind;
  requires org.apache.logging.log4j;
  requires org.apache.logging.log4j.core;

  exports dev.erst.gridgrind.cli;

  opens dev.erst.gridgrind.cli to
      tools.jackson.databind;
  opens dev.erst.gridgrind.cli.render to
      tools.jackson.databind;
}
