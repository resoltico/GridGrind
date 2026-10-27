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
