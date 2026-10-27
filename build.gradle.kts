plugins {
    id("gridgrind.root-conventions")
    id("gridgrind.fuzz-conventions")
}

for (suite in listOf("containerTest", "devcontainerTest")) {
    tasks.register(suite) {
        group = "verification"
        dependsOn(":cli:" + suite)
    }
}

tasks.register("verifyGeneratedDocs") {
    group = "verification"
    dependsOn(":cli:verifyGeneratedDocs")
}
tasks.named("check") {
    dependsOn("verifyGeneratedDocs")
}
