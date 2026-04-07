plugins {
    java
}

repositories {
    gradlePluginPortal()
}

dependencies {
    implementation(gradleApi())
}

java {
    toolchain {
        languageVersion = JavaLanguageVersion.of(25)
    }
}
