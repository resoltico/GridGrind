package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import java.util.Objects;

/** Focused hyperlink builders for the fluent Java authoring surface. */
public final class Links {
  /** Hyperlink target wrapper for the focused authoring surface. */
  public sealed interface Target permits Url, Email, FileTarget, DocumentTarget {}

  /** Absolute URL hyperlink target. */
  public record Url(String target) implements Target {
    public Url {
      Objects.requireNonNull(target, "target must not be null");
    }
  }

  /** Email hyperlink target. */
  public record Email(String email) implements Target {
    public Email {
      Objects.requireNonNull(email, "email must not be null");
    }
  }

  /** Local or shared file hyperlink target. */
  public record FileTarget(String path) implements Target {
    public FileTarget {
      Objects.requireNonNull(path, "path must not be null");
    }
  }

  /** Internal workbook hyperlink target. */
  public record DocumentTarget(String target) implements Target {
    public DocumentTarget {
      Objects.requireNonNull(target, "target must not be null");
    }
  }

  private Links() {}

  /** Returns an absolute URL hyperlink target. */
  public static Target url(String target) {
    return new Url(target);
  }

  /** Returns an email hyperlink target. */
  public static Target email(String email) {
    return new Email(email);
  }

  /** Returns a file hyperlink target. */
  public static Target file(String path) {
    return new FileTarget(path);
  }

  /** Returns an internal workbook hyperlink target. */
  public static Target document(String target) {
    return new DocumentTarget(target);
  }

  static HyperlinkTarget toHyperlinkTarget(Target target) {
    return switch (Objects.requireNonNull(target, "target must not be null")) {
      case Url url -> new HyperlinkTarget.Url(url.target());
      case Email email -> new HyperlinkTarget.Email(email.email());
      case FileTarget fileTarget -> new HyperlinkTarget.File(fileTarget.path());
      case DocumentTarget documentTarget -> new HyperlinkTarget.Document(documentTarget.target());
    };
  }
}
