package dev.erst.gridgrind.buildlogic;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.util.Set;
import java.util.TreeSet;
import org.gradle.api.DefaultTask;
import org.gradle.api.GradleException;
import org.gradle.api.file.ConfigurableFileCollection;
import org.gradle.api.file.RegularFileProperty;
import org.gradle.api.provider.SetProperty;
import org.gradle.api.tasks.Classpath;
import org.gradle.api.tasks.Input;
import org.gradle.api.tasks.InputFile;
import org.gradle.api.tasks.PathSensitive;
import org.gradle.api.tasks.PathSensitivity;
import org.gradle.api.tasks.TaskAction;

/** Fails packaging when its external runtime artifacts or audited NOTICE markers drift. */
public abstract class VerifyRuntimeLegalInventoryTask extends DefaultTask {
  @Classpath
  public abstract ConfigurableFileCollection getRuntimeArtifacts();

  @Input
  public abstract SetProperty<String> getAuditedRuntimeArtifactNames();

  @Input
  public abstract SetProperty<String> getAuditedNoticeMarkers();

  @InputFile
  @PathSensitive(PathSensitivity.RELATIVE)
  public abstract RegularFileProperty getNoticeFile();

  @TaskAction
  void verifyInventory() throws IOException {
    Set<String> resolvedArtifactNames = new TreeSet<>();
    for (File artifact : getRuntimeArtifacts().getFiles()) {
      if (!resolvedArtifactNames.add(artifact.getName())) {
        throw new GradleException(
            "Runtime legal inventory contains duplicate artifact name: " + artifact.getName());
      }
    }

    Set<String> auditedArtifactNames = new TreeSet<>(getAuditedRuntimeArtifactNames().get());
    Set<String> unaudited = difference(resolvedArtifactNames, auditedArtifactNames);
    Set<String> noLongerResolved = difference(auditedArtifactNames, resolvedArtifactNames);
    if (!unaudited.isEmpty() || !noLongerResolved.isEmpty()) {
      throw new GradleException(inventoryDriftMessage(unaudited, noLongerResolved));
    }

    String noticeText = Files.readString(getNoticeFile().get().getAsFile().toPath());
    Set<String> missingMarkers = new TreeSet<>();
    for (String marker : getAuditedNoticeMarkers().get()) {
      if (!noticeText.contains(marker)) {
        missingMarkers.add(marker);
      }
    }
    if (!missingMarkers.isEmpty()) {
      throw new GradleException("NOTICE is missing audited runtime markers: " + missingMarkers);
    }
  }

  private static Set<String> difference(Set<String> left, Set<String> right) {
    Set<String> difference = new TreeSet<>(left);
    difference.removeAll(right);
    return difference;
  }

  private static String inventoryDriftMessage(
      Set<String> unaudited, Set<String> noLongerResolved) {
    StringBuilder message = new StringBuilder("Runtime legal inventory drifted.");
    if (!unaudited.isEmpty()) {
      message.append(" Unaudited: ").append(unaudited).append('.');
    }
    if (!noLongerResolved.isEmpty()) {
      message.append(" No longer resolved: ").append(noLongerResolved).append('.');
    }
    return message
        .append(" Audit the changed artifacts, then update NOTICE and the audited artifact set together.")
        .toString();
  }
}
