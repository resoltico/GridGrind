---
afad: "5.0.1"
domain: CONFORMANCE
updated: "2026-09-09"
route:
  keywords: [gridgrind, conformance, grant, authority, evidence, publication, determinism, signature preview, filesystem, no-follow, docker]
  questions: ["what GridGrind guarantees are proven", "how does GridGrind establish publication state", "how does GridGrind constrain execution authority", "is signature preview reproducible", "which filesystems support secure GridGrind paths"]
---

# Conformance Record

This record names GridGrind guarantees that require evidence beyond ordinary unit tests and identifies the gate that supplies that evidence. A passing local gate never implies an unmeasured cross-host or cross-filesystem guarantee.

| Guarantee | Current evidence | Gate | Claim boundary |
|:----------|:-----------------|:-----|:---------------|
| V3 execution authority | Every executable Java, CLI, and bound doctor request receives an explicit host grant; missing or insufficient authority is rejected before workbook work. | `GridGrindEngineGrantTest`, `GridGrindCliExecutionGrantIntegrationTest`, `ExecutionGrantValidatorTest`, and `./check.sh`. | The grant constrains only the current execution. It does not prove that a host chose a sufficiently narrow grant. |
| Typed secret handling | Request documents carry references, while a host resolver supplies only granted secret material and result surfaces exclude it. | `CliFileSecretResolverTest`, request-redaction tests, and V3 execution regressions. | This proves GridGrind's request/result boundary; hosts remain responsible for their resolver and secret-store security. |
| Verified publication outcome | A persisted workbook is staged privately, structurally reopened, and then published through the bound destination. The result distinguishes `NOT_ATTEMPTED`, `NOT_PUBLISHED`, `PUBLISHED`, and `UNCERTAIN`. | `RequestPathPublicationTest`, `FullXssfPersistenceTest`, publication regressions, and Docker smoke. | `UNCERTAIN` means a destination may have changed and requires inspection; it is never safe to infer a retry. |
| Execution evidence | Results distinguish admitted operation effects and input identities from structural, computational, task-specific, preservation, and presentational claims. | `WorkbookExecutionEvidenceTest`, `ExecutionEvidenceTest`, and host-acceptance regressions. | A presentational claim remains `NOT_ASSESSED` unless an independent trusted renderer supplies evidence. |
| `SUMMARY` response determinism | Repeated execution serializes byte-identically with timing omitted. | `ExecutionJournalCoverageTest.summarySuccessResponsesSerializeDeterministicallyAcrossRepeatedRuns` and `./check.sh`. | Applies to identical requests on one supported runtime. |
| Request-owned path safety | Descriptor-relative, no-follow binding rejects unavailable capabilities and observed topology mutation before publication. | Request-path and preflight regressions plus Docker smoke. | Under a stable topology, writes remain beneath the execution root. Concurrent mutation in the documented residual window is not claimed safe. |
| Signature-line preview imagery | OOXML metadata and preview-image round trips are tested. | Engine drawing tests and Docker smoke signature-line authoring. | Cross-host pixel reproducibility is unproven until the same request is measured on at least two font configurations. Treat preview imagery as environment-sensitive. |
| No-follow identity capability | The supported local and Docker filesystems prove required handle operations; unsupported capability fails closed. | Path binding tests and Docker smoke. | Every supported OS/filesystem requires an explicit capability result before release qualification. Never fall back to path-string revalidation. |

## Release Qualification

Before declaring a new operating system or filesystem supported, capture the V3 grant, path-safety, and publication fixture results on that environment. Before claiming cross-host signature-preview byte reproducibility, retain two independent font-configuration captures. Until then, the claim boundaries above are the public contract.
