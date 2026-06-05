package dev.erst.gridgrind.cli;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.regex.Pattern;

/** Shared low-level wrapping and width helpers for CLI-owned text surfaces. */
final class GridGrindCliWrappingSupport {
  private static final int DEFAULT_HELP_TEXT_WIDTH = 88;
  private static final Pattern WHITESPACE_PATTERN = Pattern.compile("\\s+");

  private GridGrindCliWrappingSupport() {}

  static int helpTextWidth(String columns) {
    if (columns == null || columns.isBlank()) {
      return DEFAULT_HELP_TEXT_WIDTH;
    }
    try {
      return Math.max(72, Math.min(Integer.parseInt(columns.trim()), 120));
    } catch (NumberFormatException ignored) {
      return DEFAULT_HELP_TEXT_WIDTH;
    }
  }

  static String wrappedText(String text, String firstPrefix, String continuationPrefix, int width) {
    Objects.requireNonNull(text, "text must not be null");
    Objects.requireNonNull(firstPrefix, "firstPrefix must not be null");
    Objects.requireNonNull(continuationPrefix, "continuationPrefix must not be null");
    String normalizedText = text.trim();
    StringBuilder builder = new StringBuilder(firstPrefix);
    String currentPrefix = firstPrefix;
    int lineLength = currentPrefix.length();
    for (String token : wrappingTokens(normalizedText)) {
      int separatorWidth = lineLength == currentPrefix.length() ? 0 : 1;
      if (lineLength + separatorWidth + token.length() > width
          && lineLength > currentPrefix.length()) {
        builder.append('\n').append(continuationPrefix).append(token);
        currentPrefix = continuationPrefix;
        lineLength = continuationPrefix.length() + token.length();
      } else {
        if (separatorWidth == 1) {
          builder.append(' ');
          lineLength++;
        }
        builder.append(token);
        lineLength += token.length();
      }
    }
    return builder.toString();
  }

  static List<String> wrappingTokens(String text) {
    List<String> rawTokens =
        WHITESPACE_PATTERN.splitAsStream(text).filter(token -> !token.isBlank()).toList();
    if (rawTokens.isEmpty()) {
      return List.of();
    }
    List<String> tokens = new ArrayList<>();
    int index = 0;
    while (index < rawTokens.size()) {
      String token = rawTokens.get(index);
      if (isRedirectionToken(token) && index + 1 < rawTokens.size()) {
        tokens.add(token + " " + rawTokens.get(index + 1));
        index += 2;
      } else {
        tokens.add(token);
        index++;
      }
    }
    return tokens;
  }

  private static boolean isRedirectionToken(String token) {
    return switch (token) {
      case "<", ">", "1>", "2>", ">>", "1>>", "2>>" -> true;
      default -> false;
    };
  }
}
