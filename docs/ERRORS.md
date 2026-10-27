<!-- Generated from core metadata; do not edit. -->
# Problems and recovery

## INVALID_ARGUMENTS

Invalid host arguments

Category: ARGUMENTS. Recovery: CHANGE_REQUEST.

Fix the host arguments and rerun the command.

## INVALID_JSON

Invalid JSON payload

Category: REQUEST. Recovery: CHANGE_REQUEST.

Send syntactically valid JSON.

## INVALID_ENCODING

Invalid request encoding

Category: REQUEST. Recovery: CHANGE_REQUEST.

Encode the request document as UTF-8.

## INVALID_REQUEST_SHAPE

Invalid request shape

Category: REQUEST. Recovery: CHANGE_REQUEST.

Send a payload whose fields and discriminator IDs match the GridGrind protocol.

## INVALID_REQUEST

Invalid request

Category: REQUEST. Recovery: CHANGE_REQUEST.

Fix the request data and retry the workflow.

## NUMBER_NOT_REPRESENTABLE

Number cannot be represented exactly

Category: REQUEST. Recovery: CHANGE_REQUEST.

Store identifiers or precision-sensitive values as TEXT, or use a value exactly representable as an IEEE-754 number.

## PATH_ESCAPES_ROOT

Path escapes execution root

Category: REQUEST. Recovery: CHANGE_REQUEST.

Use a request-owned path inside the execution root; absolute paths are allowed only when they resolve inside that root.

## ASSERTION_FAILED

Assertion failed

Category: ASSERTION. Recovery: CHANGE_REQUEST.

Inspect the observed workbook facts, then adjust the plan expectations or authored mutations and retry.

## PRESERVATION_FAILED

Preservation requirement failed

Category: ASSERTION. Recovery: CHANGE_REQUEST.

Inspect the protected facts, then adjust the plan mutations or the host-required preservation selection and retry.

## INVALID_CELL_ADDRESS

Invalid cell address

Category: REQUEST. Recovery: CHANGE_REQUEST.

Use a valid A1-style address such as A1 or BC12.

## INVALID_RANGE_ADDRESS

Invalid range address

Category: REQUEST. Recovery: CHANGE_REQUEST.

Use a valid A1-style range such as A1:C3 or a single address like B2.

## INVALID_FORMULA

Invalid formula

Category: FORMULA. Recovery: CHANGE_REQUEST.

Fix the formula syntax or workbook references, then retry.

## INVALID_FORMULA_TEXT

Invalid opaque formula text

Category: FORMULA. Recovery: CHANGE_REQUEST.

Provide nonempty XML 1.0 formula character data without a leading =.

## UNSUPPORTED_FORMULA_CONSTRUCT

Unsupported formula construct

Category: FORMULA. Recovery: CHANGE_REQUEST.

Use a formula construct that Apache POI can parse on the authored request path, then retry.

## CIRCULAR_FORMULA_REFERENCE

Circular formula reference

Category: FORMULA. Recovery: CHANGE_REQUEST.

Break the circular formula dependency or use a non-evaluating calculation strategy.

## MISSING_EXTERNAL_WORKBOOK

Missing external workbook

Category: FORMULA. Recovery: CHANGE_REQUEST.

Bind the referenced external workbook or switch to cached-value evaluation.

## UNREGISTERED_USER_DEFINED_FUNCTION

Unregistered user-defined function

Category: FORMULA. Recovery: CHANGE_REQUEST.

Register the required UDF in formulaEnvironment or skip server-side evaluation.

## UNSUPPORTED_FORMULA

Unsupported formula

Category: FORMULA. Recovery: CHANGE_REQUEST.

Use a formula supported by Apache POI or skip server-side formula evaluation.

## WORKBOOK_NOT_FOUND

Workbook not found

Category: RESOURCE. Recovery: CHANGE_REQUEST.

Create the workbook first or provide an existing workbook path.

## INPUT_SOURCE_NOT_FOUND

Input source not found

Category: RESOURCE. Recovery: CHANGE_REQUEST.

Provide an existing authored input file or correct its path. Relative request-owned paths resolve from the explicit execution root.

## INPUT_SOURCE_UNAVAILABLE

Input source unavailable

Category: REQUEST. Recovery: CHANGE_REQUEST.

Bind the required STANDARD_INPUT content or change the request to use INLINE or FILE sources.

## INPUT_SOURCE_IO_ERROR

Input source I/O failure

Category: IO. Recovery: CHECK_ENVIRONMENT.

Check authored input paths, permissions, file locks, and transport bindings before retrying.

## SHEET_NOT_FOUND

Sheet not found

Category: RESOURCE. Recovery: CHANGE_REQUEST.

Create the sheet first or correct the sheet name.

## NAMED_RANGE_NOT_FOUND

Named range not found

Category: RESOURCE. Recovery: CHANGE_REQUEST.

Create the named range first or correct the requested scope and name.

## CELL_NOT_FOUND

Cell not found

Category: RESOURCE. Recovery: CHANGE_REQUEST.

Write the cell first or adjust the analysis target.

## WORKBOOK_NOT_OPENABLE

Workbook package is not openable

Category: REQUEST. Recovery: CHANGE_REQUEST.

Provide a valid `.xlsx` OOXML workbook package.

## SOURCE_PATH_IS_DIRECTORY

Workbook source path is a directory

Category: REQUEST. Recovery: CHANGE_REQUEST.

Provide a workbook file path instead of a directory.

## OUTPUT_PATH_IS_DIRECTORY

Workbook output path is a directory

Category: REQUEST. Recovery: CHANGE_REQUEST.

Provide a workbook file path instead of a directory.

## OUTPUT_PATH_ALREADY_EXISTS

Workbook output path already exists

Category: REQUEST. Recovery: CHANGE_REQUEST.

Choose a new output path, use ifExists=REPLACE, or remove the conflicting file.

## WORKBOOK_PASSWORD_REQUIRED

Workbook password required

Category: SECURITY. Recovery: CHANGE_REQUEST.

Provide source.security.password for the encrypted workbook and retry.

## INVALID_WORKBOOK_PASSWORD

Invalid workbook password

Category: SECURITY. Recovery: CHANGE_REQUEST.

Supply the correct source.security.password for the encrypted workbook and retry.

## ENCRYPTION_SOURCE_NOT_ENCRYPTED

Source workbook is not encrypted

Category: SECURITY. Recovery: CHANGE_REQUEST.

Use persistence.security.encryption.type=NONE or ENCRYPT for a plaintext source workbook.

## ENCRYPTION_SOURCE_NOT_PRESERVABLE

Source encryption cannot be preserved

Category: SECURITY. Recovery: CHANGE_REQUEST.

Use persistence.security.encryption.type=NONE or ENCRYPT when the source encryption envelope is not compatible with GridGrind's AGILE write contract.

## INVALID_SIGNING_CONFIGURATION

Invalid signing configuration

Category: SECURITY. Recovery: CHANGE_REQUEST.

Fix persistence.security.signature keystore, alias, password, or digest settings and retry.

## WORKBOOK_SECURITY_ERROR

Workbook security failure

Category: SECURITY. Recovery: CHECK_ENVIRONMENT.

Check the secure workbook package, cryptographic material, and runtime environment before retrying.

## UNSAFE_PATH_ACCESS

Unsafe path access

Category: SECURITY. Recovery: CHECK_ENVIRONMENT.

Use a filesystem that supports GridGrind's no-follow path binding and retry after any observed path topology change is resolved.

## AUTHORITY_DENIED

Execution authority denied

Category: SECURITY. Recovery: CHANGE_REQUEST.

Request a host grant that explicitly permits the required resource, operation, target, or publication effect.

## PUBLICATION_UNCERTAIN

Workbook publication state is uncertain

Category: IO. Recovery: ESCALATE.

Inspect the destination workbook and its recorded publication evidence before any retry; do not rerun the plan blindly.

## IO_ERROR

I/O failure

Category: IO. Recovery: CHECK_ENVIRONMENT.

Check file paths, permissions, file locks, and disk state before retrying.

## INTERNAL_ERROR

Internal GridGrind failure

Category: INTERNAL. Recovery: ESCALATE.

Capture the problem details and escalate; this indicates an unexpected runtime failure.
