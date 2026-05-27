package dev.erst.gridgrind.buildlogic;

import com.sun.source.tree.CaseTree;
import com.sun.source.tree.ClassTree;
import com.sun.source.tree.CompilationUnitTree;
import com.sun.source.tree.MethodTree;
import com.sun.source.tree.SwitchExpressionTree;
import com.sun.source.tree.SwitchTree;
import com.sun.source.tree.Tree;
import com.sun.source.tree.VariableTree;
import com.sun.source.util.JavacTask;
import com.sun.source.util.TreeScanner;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Locale;
import javax.lang.model.element.Modifier;
import javax.tools.JavaCompiler;
import javax.tools.JavaFileObject;
import javax.tools.StandardJavaFileManager;
import javax.tools.ToolProvider;
import org.gradle.api.GradleException;

public final class JavaSourceShapeAnalyzer {
  Metrics analyze(Path sourceFile, int javaRelease) throws IOException {
    long lineCount;
    try (var lines = Files.lines(sourceFile)) {
      lineCount = lines.count();
    }

    JavaCompiler compiler = ToolProvider.getSystemJavaCompiler();
    if (compiler == null) {
      throw new GradleException(
          "Java source-shape verification requires a JDK compiler, but no system compiler was"
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
      ShapeScanner scanner = new ShapeScanner();
      scanner.scan(compilationUnit, null);
      return new Metrics(
          lineCount,
          compilationUnit.getImports().size(),
          scanner.topLevelTypeCount,
          scanner.nestedTypeCount,
          scanner.methodCount,
          scanner.publicMethodCount,
          scanner.fieldCount,
          scanner.switchCount,
          scanner.maxSwitchArms);
    } catch (RuntimeException exception) {
      throw new GradleException(
          "Failed to analyze Java source shape for " + sourceFile.toAbsolutePath().normalize(),
          exception);
    }
  }

  record Metrics(
      long lineCount,
      int importCount,
      int topLevelTypeCount,
      int nestedTypeCount,
      int methodCount,
      int publicMethodCount,
      int fieldCount,
      int switchCount,
      int maxSwitchArms) {}

  private static final class ShapeScanner extends TreeScanner<Void, Void> {
    private int nestingDepth;
    private int topLevelTypeCount;
    private int nestedTypeCount;
    private int methodCount;
    private int publicMethodCount;
    private int fieldCount;
    private int switchCount;
    private int maxSwitchArms;

    @Override
    public Void visitClass(ClassTree classTree, Void unused) {
      if (nestingDepth == 0) {
        topLevelTypeCount++;
      } else {
        nestedTypeCount++;
      }
      nestingDepth++;
      try {
        for (Tree member : classTree.getMembers()) {
          if (member instanceof VariableTree) {
            fieldCount++;
          }
        }
        return super.visitClass(classTree, unused);
      } finally {
        nestingDepth--;
      }
    }

    @Override
    public Void visitMethod(MethodTree methodTree, Void unused) {
      methodCount++;
      if (!methodTree.getName().contentEquals("<init>")
          && methodTree.getModifiers().getFlags().contains(Modifier.PUBLIC)) {
        publicMethodCount++;
      }
      return super.visitMethod(methodTree, unused);
    }

    @Override
    public Void visitSwitch(SwitchTree switchTree, Void unused) {
      recordSwitch(switchTree.getCases());
      return super.visitSwitch(switchTree, unused);
    }

    @Override
    public Void visitSwitchExpression(SwitchExpressionTree switchExpressionTree, Void unused) {
      recordSwitch(switchExpressionTree.getCases());
      return super.visitSwitchExpression(switchExpressionTree, unused);
    }

    private void recordSwitch(List<? extends CaseTree> cases) {
      switchCount++;
      maxSwitchArms = Math.max(maxSwitchArms, cases.size());
    }
  }
}
