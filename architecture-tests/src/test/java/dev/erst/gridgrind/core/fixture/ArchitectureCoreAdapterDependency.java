package dev.erst.gridgrind.core.fixture;

/** Negative control: a core type must not depend on an application adapter. */
public record ArchitectureCoreAdapterDependency(
    dev.erst.gridgrind.engine.api.GridGrindRequestInputs inputs) {}
