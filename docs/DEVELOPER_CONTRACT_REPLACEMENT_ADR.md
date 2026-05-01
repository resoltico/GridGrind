---
afad: "4.0"
version: "0.62.0"
domain: DEVELOPER
updated: "2026-05-01"
route:
  keywords: [gridgrind, adr, contract, executor, protocol, replacement, architecture, java, authoring-java]
  questions: ["what architecture decision locked the contract replacement", "why did gridgrind delete the protocol module", "what is contract replacement mode", "why is gridgrind split into contract and executor"]
---

# Architecture Decision Record: Contract Replacement Mode

**Status**: Accepted
**Decision date**: 2026-04-17
---

## Decision

GridGrind is in hard-break contract-replacement mode.

The old monolithic `protocol` module is deleted and replaced by the current six-module product
graph:

```text
dev.erst.gridgrind.authoring -> dev.erst.gridgrind.contract -> dev.erst.gridgrind.excel.foundation
dev.erst.gridgrind.cli -> dev.erst.gridgrind.executor -> dev.erst.gridgrind.contract -> dev.erst.gridgrind.excel.foundation
dev.erst.gridgrind.executor -> dev.erst.gridgrind.engine -> dev.erst.gridgrind.excel.foundation
```

The ownership boundaries are now:

- `excel-foundation`
  - shared POI-free Excel-domain foundation types, limits, and value objects
- `engine`
  - workbook domain behavior and POI-backed execution
- `contract`
  - canonical public contract model, metadata registry, and JSON codecs
- `executor`
  - the only bridge from contract to engine
- `authoring-java`
  - fluent Java authoring layer built strictly above the canonical contract and intentionally
    separate from execution
- `cli`
  - thin transport adapter and discovery surface

No new top-level request growth may reintroduce monolithic transport-plus-execution ownership.
Future redesign work must proceed through new accepted decision records or new blocking redesign
programs rooted in the post-replacement module graph. The later `excel-foundation` extraction is
part of that same post-replacement architecture: it preserves the no-`protocol` rule while making
`contract` a real boundary instead of a façade over engine-owned Excel types.

---

## Context

The former `protocol` module mixed three concerns that should not have shared one boundary:

- canonical public model ownership
- execution bridge ownership
- downstream transport/discovery support

That coupling was tolerable for the original surface area but it is the wrong base for the
next-generation Java-first system. Selector-centric targeting, first-class assertions, execution
journals, calculation policies, and an ergonomic Java authoring layer all require cleaner
boundaries.

The redesign therefore starts by deleting the monolith instead of layering more capability onto a
confused ownership model.

---

## Non-Negotiable Consequences

1. Java remains the only first-class implementation language.
2. Apache POI remains the only workbook execution engine.
3. The contract is Java-owned and transport-neutral.
4. JSON remains an encoding and interchange format, not the primary design center.
5. The CLI remains thin downstream from the core product graph.
6. No backward compatibility, migration layer, alias path, or compatibility bridge is allowed for
   the redesign.
7. Higher-level authoring work must build on the canonical `contract`; execution remains a
   separate concern layered through `executor`.

---

## Immediate Rules

- Do not add a new top-level Gradle product module that bypasses `executor`.
- Do not reintroduce `protocol` as a top-level module, JPMS module, or package root.
- Do not move execution logic back into `contract`.
- Do not put transport-specific behavior in `engine` or `executor`.
- Do not reintroduce separate `operations[]` and `reads[]` arrays; the canonical contract is the
  ordered `steps[]` workflow surface.

---

## Acceptance Evidence

Phase 0 and Phase 1 of the redesign are complete when all of the following are true:

- this ADR is accepted and linked from developer documentation
- `settings.gradle.kts` includes `excel-foundation`, `engine`, `contract`, `executor`,
  `authoring-java`, and `cli`, and no longer includes `protocol`
- the old top-level `protocol/` module is absent
- the CLI depends on `executor`, not directly on the deleted `protocol` module
- `authoring-java` depends on `contract`, not directly on `executor`, `engine`, or the deleted
  `protocol` module
- help and catalog discovery still render from the canonical contract registry
- relocated tests, parity suites, and quality gates remain green

---

## Follow-On Work

This ADR intentionally does not restate every detailed contract surface. Those details now live in
the implemented modules, their tests, and the surrounding developer/reference documentation for
the completed replacement system.
