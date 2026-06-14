package dev.erst.gridgrind.buildlogic;

import java.io.BufferedWriter;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;

final class SourceShapeBudgetSupport {
  private SourceShapeBudgetSupport() {}

  static List<String> exceededMetrics(
      SourceShapeMetrics metrics, JavaSourceShapePolicy.Rule rule) {
    List<String> exceededMetrics = new ArrayList<>();
    addExceededMetric(exceededMetrics, "lines", metrics.lineCount(), rule.maxLines());
    addExceededMetric(exceededMetrics, "methods", metrics.methodCount(), rule.maxMethods());
    addExceededMetric(
        exceededMetrics,
        "publicMethods",
        metrics.publicMethodCount(),
        rule.maxPublicMethods());
    addExceededMetric(exceededMetrics, "imports", metrics.importCount(), rule.maxImports());
    addExceededMetric(exceededMetrics, "fields", metrics.fieldCount(), rule.maxFields());
    addExceededMetric(
        exceededMetrics, "nestedTypes", metrics.nestedTypeCount(), rule.maxNestedTypes());
    addExceededMetric(exceededMetrics, "switches", metrics.switchCount(), rule.maxSwitches());
    addExceededMetric(
        exceededMetrics, "maxSwitchArms", metrics.maxSwitchArms(), rule.maxSwitchArms());
    return List.copyOf(exceededMetrics);
  }

  static double riskScore(SourceShapeMetrics metrics, JavaSourceShapePolicy.Rule rule) {
    return ratio(metrics.lineCount(), rule.maxLines())
        + ratio(metrics.methodCount(), rule.maxMethods())
        + ratio(metrics.publicMethodCount(), rule.maxPublicMethods())
        + ratio(metrics.importCount(), rule.maxImports())
        + ratio(metrics.fieldCount(), rule.maxFields())
        + ratio(metrics.nestedTypeCount(), rule.maxNestedTypes())
        + ratio(metrics.switchCount(), rule.maxSwitches())
        + ratio(metrics.maxSwitchArms(), rule.maxSwitchArms());
  }

  static String reportRowTsv(
      String relativePath, JavaSourceShapePolicy.Rule rule, SourceShapeMetrics metrics) {
    return relativePath
        + '\t'
        + rule.kind()
        + '\t'
        + rule.role()
        + '\t'
        + metrics.lineCount()
        + '\t'
        + limit(rule.maxLines())
        + '\t'
        + metrics.methodCount()
        + '\t'
        + limit(rule.maxMethods())
        + '\t'
        + metrics.publicMethodCount()
        + '\t'
        + limit(rule.maxPublicMethods())
        + '\t'
        + metrics.importCount()
        + '\t'
        + limit(rule.maxImports())
        + '\t'
        + metrics.fieldCount()
        + '\t'
        + limit(rule.maxFields())
        + '\t'
        + metrics.nestedTypeCount()
        + '\t'
        + limit(rule.maxNestedTypes())
        + '\t'
        + metrics.switchCount()
        + '\t'
        + limit(rule.maxSwitches())
        + '\t'
        + metrics.maxSwitchArms()
        + '\t'
        + limit(rule.maxSwitchArms())
        + '\t'
        + metrics.topLevelTypeCount()
        + '\t'
        + String.format(Locale.ROOT, "%.2f", riskScore(metrics, rule))
        + '\t'
        + rule.owner()
        + '\t'
        + limit(rule.reviewExpiresOn())
        + '\t'
        + limit(rule.splitTrigger());
  }

  static void writeBudgetReport(Path reportPath, List<? extends ReportRowView> reportRows)
      throws IOException {
    Files.createDirectories(reportPath.getParent());
    List<? extends ReportRowView> sortedRows =
        reportRows.stream()
            .sorted(Comparator.comparingDouble(ReportRowView::riskScore).reversed())
            .toList();
    try (BufferedWriter writer = Files.newBufferedWriter(reportPath, StandardCharsets.UTF_8)) {
      writer.write(
          "path\tkind\trole\tlines\tmaxLines\tmethods\tmaxMethods\tpublicMethods"
              + "\tmaxPublicMethods\timports\tmaxImports\tfields\tmaxFields\tnestedTypes"
              + "\tmaxNestedTypes\tswitches\tmaxSwitches\tmaxSwitchArms"
              + "\tlimitMaxSwitchArms\ttopLevelTypes\trisk\towner"
              + "\treviewExpiresOn\tsplitTrigger");
      writer.newLine();
      for (ReportRowView row : sortedRows) {
        writer.write(row.toTsv());
        writer.newLine();
      }
    }
  }

  interface ReportRowView {
    double riskScore();

    String toTsv();
  }

  private static void addExceededMetric(
      List<String> exceededMetrics, String metricName, long actual, Integer limit) {
    if (limit != null && actual > limit) {
      exceededMetrics.add(metricName + "=" + actual + ">" + limit);
    }
  }

  private static double ratio(long actual, Integer limit) {
    return limit == null || limit == 0 ? 0.0d : (double) actual / (double) limit;
  }

  private static String limit(Object value) {
    return value == null ? "-" : value.toString();
  }
}
