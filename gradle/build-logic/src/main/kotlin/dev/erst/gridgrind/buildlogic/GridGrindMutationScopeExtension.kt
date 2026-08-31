package dev.erst.gridgrind.buildlogic

import org.gradle.api.model.ObjectFactory
import org.gradle.api.provider.Property
import org.gradle.api.provider.SetProperty
import javax.inject.Inject

/** Project-owned critical mutation scope and measured baseline thresholds. */
abstract class GridGrindMutationScopeExtension @Inject constructor(objects: ObjectFactory) {
    val targetClasses: SetProperty<String> = objects.setProperty(String::class.java)
    val targetTests: SetProperty<String> = objects.setProperty(String::class.java)
    val mutationThreshold: Property<Int> = objects.property(Int::class.java)
    val coverageThreshold: Property<Int> = objects.property(Int::class.java)
    val testStrengthThreshold: Property<Int> = objects.property(Int::class.java)
    val maxSurviving: Property<Int> = objects.property(Int::class.java)
}
