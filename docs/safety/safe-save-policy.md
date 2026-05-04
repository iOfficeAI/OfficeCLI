# Safe Save Policy

OfficeCLI HWP and HWPX mutation must preserve the source file unless a
transactional save path proves the edited file can be reopened and validated.

## Current policy

The stable policy is output-first:

```text
input.hwp -> output.hwp
input.hwpx -> output.hwpx
```

Commands that mutate HWP or HWPX content must require an explicit output path
until the safe-save transaction contract is implemented for that operation.
They must not overwrite the input path by default.

## Transaction contract

A safe-save transaction must follow these gates before any in-place replace:

1. write the edited document to a temporary file in the same directory as the
   final target;
2. fsync the temporary file where the platform exposes a supported flush;
3. reopen the temporary output with the same provider;
4. reopen with an alternate provider when one is available;
5. validate the expected semantic delta, such as text, field, or table-cell
   changes;
6. for HWPX, validate ZIP readability, XML well-formedness, manifest references,
   header references, and BinData references;
7. render SVG when the provider supports it and compare the edited render with
   the expected visual tolerance;
8. write a backup before replacing an existing source file;
9. write a transaction manifest containing checks, warnings, backup path, and
   verification evidence;
10. atomically replace only after every required gate passes.

If any required check fails, the original input file must remain unchanged and
the command must return a structured validation error.

## Interface schemas

The policy is described by:

```text
schemas/interfaces/save-policy.v1.schema.json
schemas/interfaces/save-transaction.v1.schema.json
```

`save-policy` describes which checks an operation requires. `save-transaction`
describes the result returned after an output-first or in-place save attempt.

## HWP and HWPX boundaries

Binary HWP has the strictest policy because corruption is difficult to inspect
manually. HWPX is ZIP/XML and easier to inspect, but it still needs package and
reference validation before in-place writes are allowed.

Agent-facing help should keep these claims separate:

```text
documented shape
handler-supported operation
readback-verified output
safe in-place mutation
```

Only the last category may overwrite an input file, and only with backup and
transaction evidence.
