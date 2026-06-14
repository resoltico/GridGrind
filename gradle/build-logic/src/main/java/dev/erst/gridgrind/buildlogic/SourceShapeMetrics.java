package dev.erst.gridgrind.buildlogic;

record SourceShapeMetrics(
    long lineCount,
    int importCount,
    int topLevelTypeCount,
    int nestedTypeCount,
    int methodCount,
    int publicMethodCount,
    int fieldCount,
    int switchCount,
    int maxSwitchArms) {}
