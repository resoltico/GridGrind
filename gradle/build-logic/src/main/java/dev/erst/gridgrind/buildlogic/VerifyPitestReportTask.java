package dev.erst.gridgrind.buildlogic;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Map;
import java.util.TreeMap;
import javax.xml.XMLConstants;
import javax.xml.parsers.DocumentBuilderFactory;
import org.gradle.api.DefaultTask;
import org.gradle.api.GradleException;
import org.gradle.api.file.RegularFileProperty;
import org.gradle.api.tasks.InputFile;
import org.gradle.api.tasks.OutputFile;
import org.gradle.api.tasks.PathSensitive;
import org.gradle.api.tasks.PathSensitivity;
import org.gradle.api.tasks.TaskAction;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

/** Rejects every PIT outcome other than a killed mutant, including timeout-derived detections. */
public abstract class VerifyPitestReportTask extends DefaultTask {
  @InputFile
  @PathSensitive(PathSensitivity.RELATIVE)
  public abstract RegularFileProperty getMutationReport();

  @OutputFile
  public abstract RegularFileProperty getVerificationReport();

  @TaskAction
  void verifyReport() throws Exception {
    NodeList mutations =
        secureDocumentBuilderFactory()
            .newDocumentBuilder()
            .parse(mutationReport().toFile())
            .getElementsByTagName("mutation");
    if (mutations.getLength() == 0) {
      throw new GradleException("PIT produced no mutation entries: " + mutationReport());
    }
    Map<String, Integer> statuses = new TreeMap<>();
    ArrayList<String> unexpected = new ArrayList<>();
    for (int index = 0; index < mutations.getLength(); index++) {
      Element mutation = (Element) mutations.item(index);
      String status = mutation.getAttribute("status");
      statuses.merge(status, 1, Integer::sum);
      if (!"KILLED".equals(status)) {
        unexpected.add(
            status
                + " "
                + childText(mutation, "mutatedClass")
                + "#"
                + childText(mutation, "mutatedMethod")
                + ":"
                + childText(mutation, "lineNumber"));
      }
    }
    writeVerificationReport(statuses);
    if (!unexpected.isEmpty()) {
      throw new GradleException(
          "PIT must kill every generated mutant; timeout, run-error, and uncovered outcomes are failures:\n - "
              + String.join("\n - ", unexpected.stream().limit(20).toList()));
    }
  }

  private Path mutationReport() {
    return getMutationReport().get().getAsFile().toPath();
  }

  private void writeVerificationReport(Map<String, Integer> statuses) throws IOException {
    Path reportPath = getVerificationReport().get().getAsFile().toPath();
    Files.createDirectories(reportPath.getParent());
    ArrayList<String> lines = new ArrayList<>();
    lines.add("status\tcount");
    statuses.forEach((status, count) -> lines.add(status + "\t" + count));
    Files.write(reportPath, lines);
  }

  private static String childText(Element parent, String childName) {
    return ((Element) parent.getElementsByTagName(childName).item(0)).getTextContent();
  }

  private static DocumentBuilderFactory secureDocumentBuilderFactory() throws Exception {
    DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
    factory.setFeature(XMLConstants.FEATURE_SECURE_PROCESSING, true);
    factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_DTD, "");
    factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_SCHEMA, "");
    factory.setExpandEntityReferences(false);
    return factory;
  }
}
