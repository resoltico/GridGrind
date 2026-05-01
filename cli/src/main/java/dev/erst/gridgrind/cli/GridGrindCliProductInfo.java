package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.catalog.GridGrindCliHelp;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.Objects;
import java.util.Properties;

/** Product metadata, help text, and license rendering for the CLI surface. */
final class GridGrindCliProductInfo {
  private GridGrindCliProductInfo() {}

  static String version() {
    return versionFrom(GridGrindCli.class.getPackage().getImplementationVersion());
  }

  static String helpText(String implementationVersion) {
    String version = versionFrom(implementationVersion);
    return GridGrindCliHelp.helpText(
        version, description(), documentRef(version), containerImageRef(version));
  }

  static String productHeader(String version, String description) {
    return "GridGrind " + version + "\n" + description;
  }

  static String versionFrom(String implementationVersion) {
    if (implementationVersion == null) {
      return "unknown";
    }
    return implementationVersion;
  }

  static String description() {
    return descriptionFrom(GridGrindCli.class);
  }

  static String descriptionFrom(Class<?> anchor) {
    return descriptionFrom(anchor.getResourceAsStream("/gridgrind.properties"));
  }

  static String descriptionFrom(InputStream stream) {
    if (stream == null) {
      return "GridGrind";
    }
    try (stream) {
      Properties properties = new Properties();
      properties.load(stream);
      String description = properties.getProperty("description", "");
      return description.isBlank() ? "GridGrind" : description;
    } catch (IOException exception) {
      return "GridGrind";
    }
  }

  static String licenseText(Class<?> anchor) {
    return licenseText(
        anchor.getResourceAsStream("/licenses/LICENSE"),
        anchor.getResourceAsStream("/licenses/NOTICE"),
        anchor.getResourceAsStream("/licenses/LICENSE-APACHE-2.0"),
        anchor.getResourceAsStream("/licenses/LICENSE-BSD-3-CLAUSE"));
  }

  static String licenseText(
      InputStream own, InputStream notice, InputStream apache, InputStream bsd) {
    String ownText = readLicenseStream(own);
    String thirdParty = buildThirdParty(notice, apache, bsd);
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

  private static String buildThirdParty(InputStream notice, InputStream apache, InputStream bsd) {
    String noticeText = readLicenseStream(notice);
    String apacheText = readLicenseStream(apache);
    String bsdText = readLicenseStream(bsd);
    int capacity = noticeText.length() + apacheText.length() + bsdText.length() + 2;
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
    if (!bsdText.isEmpty()) {
      result.append(sep).append(bsdText);
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
}
