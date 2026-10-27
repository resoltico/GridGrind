package dev.erst.gridgrind.buildlogic

import com.diffplug.spotless.FormatterFunc
import com.diffplug.spotless.FormatterStep
import com.google.googlejavaformat.java.Formatter
import com.google.googlejavaformat.java.ImportOrderer
import com.google.googlejavaformat.java.JavaFormatterOptions
import com.google.googlejavaformat.java.RemoveUnusedImports
import java.security.MessageDigest
import java.util.HexFormat

/** Formats Java through the Google formatter API as a reusable Spotless step. */
object GridGrindGoogleJavaFormatStep {
    @JvmStatic
    fun create(version: String): FormatterStep {
        // The state tracks this integration as well as the backend version for applicable reuse.
        val implementation = checkNotNull(javaClass.getResourceAsStream("GridGrindGoogleJavaFormatStep.class"))
            .use { HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(it.readAllBytes())) }
        return FormatterStep.create("google-java-format", "$version/$implementation") {
            val formatter = Formatter(JavaFormatterOptions.defaultOptions())
            FormatterFunc { source ->
                ImportOrderer.reorderImports(
                    RemoveUnusedImports.removeUnusedImports(formatter.formatSource(source)),
                    JavaFormatterOptions.Style.GOOGLE,
                )
            }
        }
    }
}
