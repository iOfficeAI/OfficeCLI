---
title: "Using Batch Safely"
summary: "How to prepare, run, verify, and recover OfficeCLI batch operations without command-shape or flush mistakes."
topics: [guides, batch]
sources:
  - id: batch-command
    type: file
    path: src/officecli/CommandBuilder.Batch.cs
  - id: batch-types
    type: file
    path: src/officecli/BatchTypes.cs
  - id: readme
    type: file
    path: README.md
---

Use batch when many OfficeCLI operations must run against one document in one pass. A safe batch is a JSON array of small command objects, not a list of shell commands: each item names a bare verb such as `add`, `set`, `remove`, `move`, `swap`, `get`, or `query`, and puts arguments in sibling fields such as `path`, `parent`, `type`, `props`, `to`, and `path2` [@batch-command] [@batch-types]. The successful outcome is a replay that reports every item result, leaves the document readable by later commands, and does not hide malformed fields or partial failures.

## Write Items In Protocol Shape

Start from the [batch item protocol](../concepts/batch-item-protocol), not from a command-line transcript. Do not put `add /slide[1] --type shape` inside `command`; write `"command": "add"`, `"parent": "/slide[1]"`, `"type": "shape"`, and a `props` object instead [@batch-command]. `props` can also be an array of `key=value` strings, because the batch deserializer accepts both the object form and the CLI-style array form [@batch-types].

Use only known item fields. The CLI pre-validates every object before deserialization and rejects unknown fields with a message that names the invalid fields and lists valid ones [@batch-command]. This is the first thing to check when a batch fails before opening the document.

## Choose One Input Source

Pass the JSON array with `--commands`, `--input <file>`, `--input -`, or stdin. Do not combine `--commands` and `--input`; the command rejects that combination [@batch-command]. If stdin is redirected while `--commands` or a real `--input` file is also present, OfficeCLI warns that stdin will be ignored unless `OFFICECLI_BATCH_ALLOW_STDIN_REDIRECT=1` is set [@batch-command].

If piping from another OfficeCLI command, keep the normal JSON envelope. Batch unwraps a root object whose `data` property is an array, so output shaped like `{"success":true,"data":[...]}` can feed back into `batch` without an extra transform [@batch-command].

## Pick Failure Behavior

The default batch behavior is continue-on-error: failed rows are recorded and later rows still run [@batch-command]. Use `--stop-on-error` when a later item would be unsafe after an earlier failure, such as applying formatting to a node that should have been created by the previous item [@batch-command].

Read the result as a per-item report. Each failed `BatchResult` includes the item index, the error, and the original item, so the caller can inspect or retry only the failed operation [@batch-types]. In JSON mode the outer batch succeeds only when all rows succeed; a single failed row flips the envelope success and exits non-zero even if other rows completed [@batch-command].

## Run Against The Right Lifetime

A non-resident batch opens the file once, runs all items, and saves once when the handler is disposed [@batch-command]. That is the safest default for a standalone replay because it avoids repeated open/save cycles.

If a resident process already owns the file, the CLI sends the whole batch to that resident as one `batch` request [@batch-command]. The resident applies changes in memory; disk bytes are made current by `save`, `close`, idle autosave, or `OFFICECLI_RESIDENT_FLUSH=each` [@batch-command]. See [batch execution](../architecture/cli/batch-execution) for the shared replay path.

## Verify And Recover

After a mutating batch, verify through OfficeCLI first. Resident reads see the live in-memory document even before a disk flush, while a separate program that opens the file directly may still see old bytes until `save`, `close`, or idle autosave runs [@batch-command].

When a batch partially fails, keep the output. The index and original item on each failed row are the recovery map [@batch-types]. Fix the item shape, remove successful rows that should not run twice, and rerun either the corrected subset or the full batch with idempotent operations. The README's quick-start flow also treats `close` as the final flush step after edits, which is a useful habit before handing the file to another tool [@readme].
