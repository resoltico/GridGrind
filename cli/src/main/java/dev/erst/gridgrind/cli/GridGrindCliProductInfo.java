package dev.erst.gridgrind.cli;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.Objects;
import java.util.Optional;
import java.util.Properties;

/** Product metadata, help text, and license rendering for the CLI surface. */
final class GridGrindCliProductInfo {
  private GridGrindCliProductInfo() {}

  static String version() {
    return versionFrom(
        GridGrindCli.class.getPackage().getImplementationVersion(), GridGrindCli.class);
  }

  static String helpText(CliCommand.HelpTopic topic, String implementationVersion) {
    String version = versionFrom(implementationVersion, GridGrindCli.class);
    return GridGrindCliHelp.helpText(
        topic, version, description(), documentRef(version), containerImageRef(version));
  }

  static String productHeader(String version, String description) {
    return "GridGrind " + version + "\n" + description;
  }

  static String versionFrom(String implementationVersion) {
    return versionFrom(implementationVersion, GridGrindCli.class);
  }

  static String versionFrom(String implementationVersion, Class<?> anchor) {
    Objects.requireNonNull(anchor, "anchor must not be null");
    return versionFrom(implementationVersion, anchor.getResourceAsStream("/gridgrind.properties"));
  }

  static String versionFrom(String implementationVersion, InputStream stream) {
    if (implementationVersion != null && !implementationVersion.isBlank()) {
      return implementationVersion;
    }
    return readProperties(stream)
        .map(properties -> properties.getProperty("version", ""))
        .filter(version -> !version.isBlank())
        .orElse("unknown");
  }

  static String description() {
    return descriptionFrom(GridGrindCli.class);
  }

  static String descriptionFrom(Class<?> anchor) {
    Objects.requireNonNull(anchor, "anchor must not be null");
    return descriptionFrom(anchor.getResourceAsStream("/gridgrind.properties"));
  }

  static String descriptionFrom(InputStream stream) {
    return readProperties(stream)
        .map(properties -> properties.getProperty("description", ""))
        .filter(description -> !description.isBlank())
        .orElse("GridGrind");
  }

  static String licenseText(Class<?> anchor) {
    return licenseText(
        anchor.getResourceAsStream("/licenses/LICENSE"),
        anchor.getResourceAsStream("/licenses/NOTICE"),
        anchor.getResourceAsStream("/licenses/LICENSE-APACHE-2.0"),
        anchor.getResourceAsStream("/licenses/LICENSE-BSD-2-CLAUSE"),
        anchor.getResourceAsStream("/licenses/LICENSE-BSD-3-CLAUSE"),
        anchor.getResourceAsStream("/licenses/LICENSE-EDL-1.0"));
  }

  static String licenseText(
      InputStream own,
      InputStream notice,
      InputStream apache,
      InputStream bsd2,
      InputStream bsd3,
      InputStream edl) {
    String ownText = readLicenseStream(own);
    String thirdParty = buildThirdParty(notice, apache, bsd2, bsd3, edl);
    if (ownText.isEmpty() && thirdParty.isEmpty()) {
      return "License information not available in this distribution.\n";
    }
    StringBuilder result = new StringBuilder(ownText.length() + thirdParty.length() + 64);
    result.append(ownText);
    if (!thirdParty.isEmpty()) {
      if (!ownText.isEmpty()) {
        result.append("\n---\n\nThird-party notices and licenses:\n\n");
      }
      result.append(thirdParty);
    }
    String text = result.toString();
    return text.endsWith("\n") ? text : text + '\n';
  }

  static String requestTemplateText(GridGrindCli.RequestTemplateBytesSupplier supplier) {
    Objects.requireNonNull(supplier, "supplier must not be null");
    try {
      return new String(supplier.get(), StandardCharsets.UTF_8);
    } catch (IOException exception) {
      throw new IllegalStateException("Failed to render the built-in request template", exception);
    }
  }

  private static String containerImageRef(String version) {
    String tag = "unknown".equals(version) ? "latest" : version;
    return "ghcr.io/resoltico/gridgrind:" + tag;
  }

  private static String documentRef(String version) {
    String gitRef = "unknown".equals(version) ? "main" : "v" + version;
    return "https://github.com/resoltico/GridGrind/blob/" + gitRef;
  }

  private static String buildThirdParty(
      InputStream notice, InputStream apache, InputStream bsd2, InputStream bsd3, InputStream edl) {
    String noticeText = readLicenseStream(notice);
    String apacheText = readLicenseStream(apache);
    String bsd2Text = readLicenseStream(bsd2);
    String bsd3Text = readLicenseStream(bsd3);
    String edlText = readLicenseStream(edl);
    int capacity =
        noticeText.length()
            + apacheText.length()
            + bsd2Text.length()
            + bsd3Text.length()
            + edlText.length();
    StringBuilder result = new StringBuilder(capacity);
    String sep = "";
    if (!noticeText.isEmpty()) {
      result.append(noticeText);
      sep = "\n";
    }
    if (!apacheText.isEmpty()) {
      result.append(sep).append(apacheText);
      sep = "\n";
    }
    if (!bsd2Text.isEmpty()) {
      result.append(sep).append(bsd2Text);
      sep = "\n";
    }
    if (!bsd3Text.isEmpty()) {
      result.append(sep).append(bsd3Text);
      sep = "\n";
    }
    if (!edlText.isEmpty()) {
      result.append(sep).append(edlText);
    }
    return result.toString();
  }

  private static String readLicenseStream(InputStream stream) {
    if (stream == null) {
      return "";
    }
    try (stream) {
      return new String(stream.readAllBytes(), StandardCharsets.UTF_8);
    } catch (IOException ignored) {
      return "";
    }
  }

  private static Optional<Properties> readProperties(InputStream stream) {
    if (stream == null) {
      return Optional.empty();
    }
    try (stream) {
      Properties properties = new Properties();
      properties.load(stream);
      return Optional.of(properties);
    } catch (IOException ignored) {
      return Optional.empty();
    }
  }
}
