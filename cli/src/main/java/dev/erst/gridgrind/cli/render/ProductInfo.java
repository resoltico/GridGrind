package dev.erst.gridgrind.cli.render;

import dev.erst.gridgrind.cli.command.CliCommand;
import dev.erst.gridgrind.cli.command.GridGrindCli;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.Objects;
import java.util.Optional;
import java.util.Properties;

/** Product metadata, help text, and license rendering for the CLI surface. */
public final class ProductInfo {
  private ProductInfo() {}

  /** Reads the packaged implementation version with a resource fallback. */
  public static String version() {
    return versionFrom(
        GridGrindCli.class.getPackage().getImplementationVersion(), GridGrindCli.class);
  }

  /** Renders one help topic using the packaged identity and reference URLs. */
  public static String helpText(CliCommand.HelpTopic topic, String implementationVersion) {
    String version = versionFrom(implementationVersion, GridGrindCli.class);
    return HelpRenderer.helpText(
        topic, version, description(), documentRef(version), containerImageRef(version));
  }

  /** Returns structured help with the shared parser and workflow surface. */
  public static CliHelpReport helpReport(CliCommand.HelpTopic topic) {
    Objects.requireNonNull(topic, "topic must not be null");
    String version = version();
    return new CliHelpReport(
        ProtocolVersion.CURRENT,
        topic.name(),
        version,
        description(),
        documentRef(version),
        containerImageRef(version),
        HelpSurface.CLI_SURFACE);
  }

  /** Returns the structured product version and documentation references. */
  public static CliVersionReport versionReport() {
    String version = version();
    return new CliVersionReport(
        ProtocolVersion.CURRENT,
        version,
        description(),
        documentRef(version),
        containerImageRef(version));
  }

  /** Returns the consolidated license surface with product identity. */
  public static CliLicenseReport licenseReport() {
    return new CliLicenseReport(
        ProtocolVersion.CURRENT, version(), licenseText(GridGrindCli.class));
  }

  /** Formats the two-line product identity header. */
  public static String productHeader(String version, String description) {
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

  /** Reads the packaged product description with a short fallback. */
  public static String description() {
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

  /** Reads and combines the packaged product and third-party legal texts. */
  public static String licenseText(Class<?> anchor) {
    return licenseText(
        anchor.getResourceAsStream("/licenses/LICENSE"),
        anchor.getResourceAsStream("/licenses/NOTICE"),
        anchor.getResourceAsStream("/licenses/LICENSE-APACHE-2.0"),
        anchor.getResourceAsStream("/licenses/LICENSE-BSD-2-CLAUSE"),
        anchor.getResourceAsStream("/licenses/LICENSE-BSD-3-CLAUSE"));
  }

  static String licenseText(
      InputStream own, InputStream notice, InputStream apache, InputStream bsd2, InputStream bsd3) {
    String ownText = readLicenseStream(own);
    String thirdParty = buildThirdParty(notice, apache, bsd2, bsd3);
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

  static String requestTemplateText(RequestTemplateBytesSupplier supplier) {
    Objects.requireNonNull(supplier, "supplier must not be null");
    try {
      return new String(supplier.get(), StandardCharsets.UTF_8);
    } catch (IOException exception) {
      throw new IllegalStateException("Failed to render the built-in request template", exception);
    }
  }

  static String containerImageRef(String version) {
    String tag = "unknown".equals(version) ? "latest" : version;
    return "ghcr.io/resoltico/gridgrind:" + tag;
  }

  static String documentRef(String version) {
    String gitRef = "unknown".equals(version) ? "main" : "v" + version;
    return "https://github.com/resoltico/GridGrind/blob/" + gitRef;
  }

  private static String buildThirdParty(
      InputStream notice, InputStream apache, InputStream bsd2, InputStream bsd3) {
    String noticeText = readLicenseStream(notice);
    String apacheText = readLicenseStream(apache);
    String bsd2Text = readLicenseStream(bsd2);
    String bsd3Text = readLicenseStream(bsd3);
    int capacity =
        noticeText.length() + apacheText.length() + bsd2Text.length() + bsd3Text.length();
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

  /** Supplies request-template bytes for help rendering. */
  @FunctionalInterface
  interface RequestTemplateBytesSupplier {
    /** Returns the UTF-8 bytes that should be embedded into the help text. */
    byte[] get() throws IOException;
  }
}
