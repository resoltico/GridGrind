---
afad: "4.0"
version: "0.73.0"
domain: CONFORMANCE
updated: "2026-08-20"
route:
  keywords: [gridgrind, conformance, determinism, signature preview, filesystem, no-follow, docker]
  questions: ["what GridGrind guarantees are proven", "is signature preview reproducible", "which filesystems support secure GridGrind paths"]
---

# Conformance Record

This record names GridGrind guarantees that require evidence beyond ordinary unit tests and identifies the gate that supplies that evidence. A passing local gate never implies an unmeasured cross-host or cross-filesystem guarantee.

| Guarantee | Current evidence | Gate | Claim boundary |
|:----------|:-----------------|:-----|:---------------|
| `SUMMARY` response determinism | Repeated execution serializes byte-identically with timing omitted. | `ExecutionJournalCoverageTest.summarySuccessResponsesSerializeDeterministicallyAcrossRepeatedRuns` and `./check.sh` | Applies to identical requests on one supported runtime. |
| Request-owned path safety | Descriptor-relative, no-follow binding rejects unavailable capabilities and observed topology mutation before commit. | Request-path doctor/executor parity tests plus Docker smoke. | Under a stable topology, writes remain beneath the execution root. Concurrent mutation in the documented residual window is not claimed safe. |
| Signature-line preview imagery | OOXML metadata and preview-image round trips are tested. | Engine drawing tests and Docker smoke signature-line authoring. | Cross-host pixel reproducibility is unproven until the same request is measured on at least two font configurations. Treat preview imagery as environment-sensitive. |
| No-follow identity capability | The supported local and Docker filesystems prove required handle operations; unsupported capability fails closed. | Path binding tests and Docker smoke. | Every supported OS/filesystem requires an explicit capability result before release qualification. Never fall back to path-string revalidation. |

## Release Qualification

Before declaring a new operating system or filesystem supported, capture the path-safety fixture result on that environment. Before claiming cross-host signature-preview byte reproducibility, retain two independent font-configuration captures. Until then, the claim boundaries above are the public contract.
