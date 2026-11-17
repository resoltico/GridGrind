# Native provider repairs

This component owns reviewed, source-pinned native provider repairs. Its runtime artifact contains
resources under `META-INF/gridgrind/native-provider/`; it does not replace parent-classloader
upstream classes. [repairs.json](repairs.json) owns the upstream source, patch and helper identities.

The producer verifies source artifacts and each selected source before applying exact patches,
compiles with the declared project JDK/bytecode target, and derives every nested/new required class
and its provenance from those outputs. Engine compilation uses an explicit patched native view;
public runtime verification uses unmodified vendor classes plus the resource capsule.

Original upstream sources and headers retain Apache-2.0 terms. GridGrind modifications and the
helper use AGPL-3.0-only; the capsule retains both license texts, mapped sources and provenance.
See [project licensing](../PROJECT-NOTICE.md) and [third-party notices](../NOTICE).
