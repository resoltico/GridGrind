package dev.erst.gridgrind.buildlogic

import org.gradle.api.attributes.Attribute

/** Named compiler repair views are never interchangeable directory artifacts. */
object NativeProviderVariants {
    val REPAIR_MODULE: Attribute<String> = Attribute.of("dev.erst.gridgrind.native-repair-module", String::class.java)
    const val POI_MODULE = "org.apache.poi.poi"
    const val OOXML_MODULE = "org.apache.poi.ooxml"
}
