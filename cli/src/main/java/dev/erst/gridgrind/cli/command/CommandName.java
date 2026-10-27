package dev.erst.gridgrind.cli.command;

import java.util.Arrays;
import java.util.Locale;
import java.util.Optional;

/** Concrete host command names shared by parsing and pre-execution diagnostics. */
enum CommandName {
  RUN,
  DOCTOR,
  CATALOG,
  SCHEMA,
  RECIPE,
  TEMPLATE,
  HELP,
  VERSION,
  LICENSE;

  String token() {
    return name().toLowerCase(Locale.ROOT);
  }

  static Optional<CommandName> forToken(String token) {
    return Arrays.stream(values()).filter(command -> command.token().equals(token)).findFirst();
  }

  static String primary(String[] args) {
    return args.length == 0
        ? HELP.token()
        : forToken(args[0]).map(CommandName::token).orElse("cli");
  }
}
