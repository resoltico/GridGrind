package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.net.URL;
import java.net.URLClassLoader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.zip.ZipFile;
import org.gradle.testfixtures.ProjectBuilder;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.objectweb.asm.ClassReader;
import org.objectweb.asm.ClassVisitor;
import org.objectweb.asm.ClassWriter;
import org.objectweb.asm.MethodVisitor;
import org.objectweb.asm.Opcodes;
import org.objectweb.asm.RecordComponentVisitor;

/** Independent modern-class controls against the pinned upstream and prepared visitors. */
class PrepareFuzzInstrumentationTest {
  @TempDir Path directory;

  @Test
  void sealedAndRecordMetadataSurviveActualVisitorInstrumentation() throws Exception {
    byte[] sealed = sealedFixture();
    var unsupported =
        assertThrows(
            UnsupportedOperationException.class,
            () ->
                new ClassReader(sealed)
                    .accept(
                        new janala.instrument.SnoopInstructionClassAdapter(
                            new ClassWriter(0), "fixture/Shape"),
                        0));
    assertTrue(unsupported.getMessage().contains("PermittedSubclasses requires ASM9"));
    var task = task();
    task.prepare();
    Path output = task.getAgentJar().get().getAsFile().toPath();
    try (var loader = new InstrumenterLoader(output)) {
      var adapter = loader.loadClass("janala.instrument.SnoopInstructionClassAdapter");
      ClassWriter sealedWriter = new ClassWriter(ClassWriter.COMPUTE_FRAMES);
      new ClassReader(sealed)
          .accept(
              (ClassVisitor)
                  adapter
                      .getConstructor(ClassVisitor.class, String.class)
                      .newInstance(sealedWriter, "fixture/Shape"),
              0);
      List<String> permitted = new ArrayList<>();
      new ClassReader(sealedWriter.toByteArray())
          .accept(
              new ClassVisitor(Opcodes.ASM9) {
                @Override
                public void visitPermittedSubclass(String name) {
                  permitted.add(name);
                }
              },
              0);
      assertEquals(List.of("fixture/Box"), permitted);

      ClassWriter recordWriter = new ClassWriter(ClassWriter.COMPUTE_FRAMES);
      new ClassReader(recordFixture())
          .accept(
              (ClassVisitor)
                  adapter
                      .getConstructor(ClassVisitor.class, String.class)
                      .newInstance(recordWriter, "fixture/Box"),
              0);
      List<String> components = new ArrayList<>();
      List<String> probes = new ArrayList<>();
      new ClassReader(recordWriter.toByteArray())
          .accept(
              new ClassVisitor(Opcodes.ASM9) {
                @Override
                public RecordComponentVisitor visitRecordComponent(
                    String name, String descriptor, String signature) {
                  components.add(name + ":" + descriptor);
                  return null;
                }

                @Override
                public MethodVisitor visitMethod(
                    int access,
                    String name,
                    String descriptor,
                    String signature,
                    String[] exceptions) {
                  return new MethodVisitor(Opcodes.ASM9) {
                    @Override
                    public void visitMethodInsn(
                        int opcode, String owner, String method, String descriptor, boolean itf) {
                      if (owner.startsWith("edu/berkeley/cs/jqf/instrument/tracing/"))
                        probes.add(method);
                    }
                  };
                }
              },
              0);
      assertEquals(List.of("value:I"), components);
      assertTrue(
          !probes.isEmpty(), "The record must receive coverage probes, not just remain readable");
    }
  }

  @Test
  void preparationIsDeterministicAndRetainsOtherUpstreamEntryBytesAndLegalNotices()
      throws Exception {
    var task = task();
    task.prepare();
    Path output = task.getAgentJar().get().getAsFile().toPath();
    byte[] first = Files.readAllBytes(output);
    task.prepare();
    assertArrayEquals(first, Files.readAllBytes(output));
    Set<String> changed =
        Set.of(
            "janala/instrument/SnoopInstructionClassAdapter.class",
            "janala/instrument/SnoopInstructionMethodAdapter.class");
    try (var upstream = new ZipFile(upstreamJar().toFile());
        var prepared = new ZipFile(output.toFile())) {
      for (var entry : upstream.stream().toList()) {
        byte[] old = upstream.getInputStream(entry).readAllBytes();
        byte[] actual = prepared.getInputStream(prepared.getEntry(entry.getName())).readAllBytes();
        if (!changed.contains(entry.getName())) assertArrayEquals(old, actual, entry.getName());
        else assertFalse(java.util.Arrays.equals(old, actual), entry.getName());
      }
      String notices =
          new String(
              prepared.getInputStream(prepared.getEntry("META-INF/LICENSE-JQF.txt")).readAllBytes(),
              java.nio.charset.StandardCharsets.UTF_8);
      assertTrue(notices.contains("The Regents of the University of California"));
      assertTrue(notices.contains("INRIA, France Telecom"));
      assertTrue(prepared.getEntry("META-INF/GRIDGRIND-INSTRUMENTATION.txt") != null);
    }
  }

  @Test
  void unknownArtifactIdentityIsRejectedBeforePublishingOutput() throws Exception {
    var task = task();
    Path different = directory.resolve("changed.jar");
    Files.copy(upstreamJar(), different);
    Files.write(different, new byte[] {0}, java.nio.file.StandardOpenOption.APPEND);
    task.getUpstreamJar().set(different.toFile());
    var error = assertThrows(IllegalStateException.class, task::prepare);
    assertTrue(error.getMessage().contains("Unreviewed JQF instrumentation artifact"));
    assertFalse(task.getAgentJar().get().getAsFile().exists());
  }

  private PrepareFuzzInstrumentation task() throws Exception {
    var project = ProjectBuilder.builder().withProjectDir(directory.toFile()).build();
    var task =
        project.getTasks().create("prepareInstrumentation", PrepareFuzzInstrumentation.class);
    task.getUpstreamJar().set(upstreamJar().toFile());
    task.getAgentJar().set(directory.resolve("prepared.jar").toFile());
    return task;
  }

  private static Path upstreamJar() throws Exception {
    return Path.of(
        janala.instrument.SnoopInstructionClassAdapter.class
            .getProtectionDomain()
            .getCodeSource()
            .getLocation()
            .toURI());
  }

  private static byte[] sealedFixture() {
    var writer = new ClassWriter(0);
    writer.visit(
        Opcodes.V26,
        Opcodes.ACC_PUBLIC | Opcodes.ACC_INTERFACE | Opcodes.ACC_ABSTRACT,
        "fixture/Shape",
        null,
        "java/lang/Object",
        null);
    writer.visitPermittedSubclass("fixture/Box");
    writer
        .visitMethod(Opcodes.ACC_PUBLIC | Opcodes.ACC_ABSTRACT, "value", "()I", null, null)
        .visitEnd();
    writer.visitEnd();
    return writer.toByteArray();
  }

  private static byte[] recordFixture() {
    var writer = new ClassWriter(0);
    writer.visit(
        Opcodes.V26,
        Opcodes.ACC_PUBLIC | Opcodes.ACC_FINAL | Opcodes.ACC_RECORD,
        "fixture/Box",
        null,
        "java/lang/Record",
        new String[] {"fixture/Shape"});
    writer.visitRecordComponent("value", "I", null).visitEnd();
    writer.visitField(Opcodes.ACC_PRIVATE | Opcodes.ACC_FINAL, "value", "I", null, null).visitEnd();
    MethodVisitor init = writer.visitMethod(Opcodes.ACC_PUBLIC, "<init>", "(I)V", null, null);
    init.visitCode();
    init.visitVarInsn(Opcodes.ALOAD, 0);
    init.visitMethodInsn(Opcodes.INVOKESPECIAL, "java/lang/Record", "<init>", "()V", false);
    init.visitVarInsn(Opcodes.ALOAD, 0);
    init.visitVarInsn(Opcodes.ILOAD, 1);
    init.visitFieldInsn(Opcodes.PUTFIELD, "fixture/Box", "value", "I");
    init.visitInsn(Opcodes.RETURN);
    init.visitMaxs(2, 2);
    init.visitEnd();
    MethodVisitor method = writer.visitMethod(Opcodes.ACC_PUBLIC, "value", "()I", null, null);
    method.visitCode();
    method.visitVarInsn(Opcodes.ALOAD, 0);
    method.visitFieldInsn(Opcodes.GETFIELD, "fixture/Box", "value", "I");
    method.visitInsn(Opcodes.IRETURN);
    method.visitMaxs(1, 1);
    method.visitEnd();
    writer.visitEnd();
    return writer.toByteArray();
  }

  private static final class InstrumenterLoader extends URLClassLoader {
    InstrumenterLoader(Path jar) throws Exception {
      super(new URL[] {jar.toUri().toURL()}, PrepareFuzzInstrumentationTest.class.getClassLoader());
    }

    @Override
    protected Class<?> loadClass(String name, boolean resolve) throws ClassNotFoundException {
      synchronized (getClassLoadingLock(name)) {
        if (!name.startsWith("janala.")) return super.loadClass(name, resolve);
        Class<?> loaded = findLoadedClass(name);
        if (loaded == null) loaded = findClass(name);
        if (resolve) resolveClass(loaded);
        return loaded;
      }
    }
  }
}
