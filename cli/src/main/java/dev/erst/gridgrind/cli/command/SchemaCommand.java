package dev.erst.gridgrind.cli.command;

import dev.erst.gridgrind.cli.io.CliResponseWriter;
import dev.erst.gridgrind.json.codec.GridGrindJsonMapperSupport;
import dev.erst.gridgrind.json.schema.JsonSchemas;
import java.io.IOException;
import java.io.OutputStream;

/** Emits schemas generated from canonical request, grant and response models. */
final class SchemaCommand {
  private SchemaCommand() {}

  static int execute(
      CliCommand.Schema command,
      boolean pretty,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter writer)
      throws IOException {
    var mapper = GridGrindJsonMapperSupport.JSON_MAPPER;
    var schema = mapper.readTree(JsonSchemas.document(command.document()));
    byte[] bytes =
        (pretty ? mapper.writerWithDefaultPrettyPrinter() : mapper.writer())
            .writeValueAsBytes(schema);
    return writer.writePayload(command.responsePath(), stdout, stderr, bytes, 0);
  }
}
