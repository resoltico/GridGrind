package dev.erst.gridgrind.buildlogic;

import com.sun.source.tree.ClassTree;
import com.sun.source.tree.CompilationUnitTree;
import com.sun.source.tree.Tree;
import com.sun.source.tree.VariableTree;
import com.sun.source.util.JavacTask;
import com.sun.source.util.TreeScanner;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.ListIterator;
import java.util.List;
import java.util.Locale;
import java.util.regex.Pattern;
import javax.lang.model.element.Modifier;
import javax.tools.JavaCompiler;
import javax.tools.JavaFileObject;
import javax.tools.StandardJavaFileManager;
import javax.tools.ToolProvider;
import org.gradle.api.GradleException;

final class JavaForbiddenTopLevelShapeAnalyzer {
  private static final int MAX_RECORD_COMPONENTS = 12;
  private static final int MIN_TAGGED_UNION_STATE_SLOTS = 3;
  private static final int MIN_TAGGED_UNION_OPTIONAL_SLOTS = 2;
  private static final int MIN_WIDE_RECORD_OPTIONAL_SLOTS = 6;
  private static final Pattern DISCRIMINATOR_NAME_PATTERN =
      Pattern.compile("(?i)^(status|mode|variant|action|state)$");

  List<Violation> analyze(Path sourceFile, int javaRelease) throws IOException {
    JavaCompiler compiler = ToolProvider.getSystemJavaCompiler();
    if (compiler == null) {
      throw new GradleException(
          "Forbidden-shape verification requires a JDK compiler, but no system compiler was"
              + " available.");
    }

    try (StandardJavaFileManager fileManager =
        compiler.getStandardFileManager(null, Locale.ROOT, StandardCharsets.UTF_8)) {
      Iterable<? extends JavaFileObject> compilationUnits =
          fileManager.getJavaFileObjects(sourceFile.toFile());
      JavacTask javacTask =
          (JavacTask)
              compiler.getTask(
                  null,
                  fileManager,
                  null,
                  List.of("--release", Integer.toString(javaRelease), "-proc:none"),
                  null,
                  compilationUnits);
      CompilationUnitTree compilationUnit = javacTask.parse().iterator().next();
      TopLevelShapeScanner scanner = new TopLevelShapeScanner();
      scanner.scan(compilationUnit, null);
      return scanner.violations();
    } catch (RuntimeException exception) {
      throw new GradleException(
          "Failed to analyze forbidden tagged-union or god-record shapes for "
              + sourceFile.toAbsolutePath().normalize(),
          exception);
    }
  }

  record Violation(
      String typeKind,
      String typeName,
      int stateSlots,
      int optionalStateSlots,
      List<String> discriminatorSlots,
      String reason) {}

  private static final class TopLevelShapeScanner extends TreeScanner<Void, Void> {
    private final List<Violation> violations = new ArrayList<>();
    private final List<String> typeNameStack = new ArrayList<>();

    @Override
    public Void visitClass(ClassTree classTree, Void unused) {
      String typeName = classTree.getSimpleName().toString();
      boolean namedType = !typeName.isBlank();
      if (namedType) {
        typeNameStack.add(typeName);
        inspectType(classTree, qualifiedTypeName());
      }
      try {
        return super.visitClass(classTree, unused);
      } finally {
        if (namedType) {
          typeNameStack.removeLast();
        }
      }
    }

    List<Violation> violations() {
      return List.copyOf(violations);
    }

    private void inspectType(ClassTree classTree, String qualifiedTypeName) {
      if (classTree.getKind() == Tree.Kind.RECORD) {
        inspectRecord(classTree, qualifiedTypeName);
        return;
      }
      if (classTree.getKind() == Tree.Kind.CLASS) {
        inspectClass(classTree, qualifiedTypeName);
      }
    }

    private void inspectRecord(ClassTree classTree, String qualifiedTypeName) {
      int componentCount = 0;
      int optionalCount = 0;
      List<String> discriminatorSlots = new ArrayList<>();
      for (Tree member : classTree.getMembers()) {
        if (!(member instanceof VariableTree variableTree)) {
          continue;
        }
        if (variableTree.getModifiers().getFlags().contains(Modifier.STATIC)) {
          continue;
        }
        componentCount++;
        String componentName = variableTree.getName().toString();
        if (isOptionalOrNullable(variableTree)) {
          optionalCount++;
        }
        if (DISCRIMINATOR_NAME_PATTERN.matcher(componentName).matches()) {
          discriminatorSlots.add(componentName);
        }
      }
      if (!discriminatorSlots.isEmpty()
          && componentCount >= MIN_TAGGED_UNION_STATE_SLOTS
          && optionalCount >= MIN_TAGGED_UNION_OPTIONAL_SLOTS) {
        violations.add(
            new Violation(
                "record",
                qualifiedTypeName,
                componentCount,
                optionalCount,
                List.copyOf(discriminatorSlots),
                "Record mixes discriminator slots "
                    + discriminatorSlots
                    + " with "
                    + optionalCount
                    + " Optional-bearing state slots; encode the alternatives as sealed variants"
                    + " instead of a tagged union record."));
      }
      if (componentCount > MAX_RECORD_COMPONENTS
          && optionalCount >= MIN_WIDE_RECORD_OPTIONAL_SLOTS) {
        violations.add(
            new Violation(
                "record",
                qualifiedTypeName,
                componentCount,
                optionalCount,
                List.copyOf(discriminatorSlots),
                "Record combines "
                    + componentCount
                    + " state slots with "
                    + optionalCount
                    + " Optional-bearing components; split the state into sealed variants or"
                    + " smaller value records."));
      }
    }

    private void inspectClass(ClassTree classTree, String qualifiedTypeName) {
      int fieldCount = 0;
      int optionalCount = 0;
      List<String> discriminatorSlots = new ArrayList<>();
      for (Tree member : classTree.getMembers()) {
        if (!(member instanceof VariableTree variableTree)) {
          continue;
        }
        if (variableTree.getModifiers().getFlags().contains(Modifier.STATIC)) {
          continue;
        }
        fieldCount++;
        String fieldName = variableTree.getName().toString();
        if (isOptionalOrNullable(variableTree)) {
          optionalCount++;
        }
        if (DISCRIMINATOR_NAME_PATTERN.matcher(fieldName).matches()) {
          discriminatorSlots.add(fieldName);
        }
      }
      if (!discriminatorSlots.isEmpty()
          && fieldCount >= MIN_TAGGED_UNION_STATE_SLOTS
          && optionalCount >= MIN_TAGGED_UNION_OPTIONAL_SLOTS) {
        violations.add(
            new Violation(
                "class",
                qualifiedTypeName,
                fieldCount,
                optionalCount,
                List.copyOf(discriminatorSlots),
                "Class mixes discriminator slots "
                    + discriminatorSlots
                    + " with "
                    + optionalCount
                    + " Optional-bearing state slots; encode the alternatives as sealed variants"
                    + " instead of a tagged union class."));
      }
    }

    private boolean isOptionalOrNullable(VariableTree variableTree) {
      String typeText = variableTree.getType().toString();
      if (typeText.equals("Optional")
          || typeText.startsWith("Optional<")
          || typeText.endsWith(".Optional")
          || typeText.contains(".Optional<")) {
        return true;
      }
      return variableTree.getModifiers().getAnnotations().stream()
          .map(annotationTree -> annotationTree.getAnnotationType().toString())
          .anyMatch(
              annotationType ->
                  annotationType.equals("Nullable") || annotationType.endsWith(".Nullable"));
    }

    private String qualifiedTypeName() {
      StringBuilder qualifiedName = new StringBuilder();
      for (ListIterator<String> iterator = typeNameStack.listIterator(); iterator.hasNext(); ) {
        if (qualifiedName.length() > 0) {
          qualifiedName.append('.');
        }
        qualifiedName.append(iterator.next());
      }
      return qualifiedName.toString();
    }
  }
}
