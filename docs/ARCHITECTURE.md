# Warranty Application Archive Architecture

## Boundaries

The application uses independent root entry scripts, a `flows/` orchestration
layer, and reusable capabilities under `modules/`. The main business flows are:

1. Application-material intake.
2. Worker-list intake.
3. Approved-PDF intake.

All subworkflows operate on the same JSON repository and recognition cache.
The summary HTML is an output adapter and is never read as business data.
Legacy Excel is read only during the one-time migration.

`_input` is the only user-facing drop zone. The input router moves Word and
image files to the application-material workflow. It recognizes each PDF once,
routes approval forms to approved-PDF intake, and marks other PDFs as
application material so they cannot be consumed by the approval workflow.
`_inbox` is an internal processing area.

Unmatched approval PDFs use a separate staging repository,
`待人工审核匹配PDF.json`. Saving the local HTML review page validates the staged
IDs, hashes and dataset revision, immediately applies every non-pending
decision, and then refreshes the authoritative dataset, formal summary HTML,
and review artifacts. The standalone apply command remains as a compatibility
path for decisions saved by older review pages. Human delete decisions are
recoverable moves to `_trash`, not filesystem deletion.

## Data authority

`质保作业申请数据.json` is the authoritative dataset. It contains:

- schema and dataset revisions;
- application cases and parsed business fields;
- required and missing material roles;
- file paths, original names, hashes and provenance;
- approved-PDF relationships;
- OCR cache entries;
- run summaries and auditable file changes.

The filesystem contains the binary evidence. JSON links each logical record to
that evidence with a relative path and SHA-256.

## Transaction model

Legacy migration is intentionally plan-first:

1. Read the flat directory and legacy workbook.
2. Build a complete JSON migration plan without changing files.
3. Compare the primary directory to a user-provided backup by path, size and
   SHA-256.
4. Execute explicit single-file move/copy operations.
5. Verify the target hash after every operation.
6. Persist JSON atomically.
7. Render summary HTML from JSON.
8. Validate JSON, filesystem evidence and HTML tab structure.

Incremental workflows also record their file operations. Duplicate evidence is
moved to quarantine rather than deleted.

## Package layout

- `src/warranty_application_archive/modules/`: parsing, naming, storage,
  recognition, rendering, validation, and filesystem adapters.
- `src/warranty_application_archive/flows/`: archive intake, migration,
  approval review, and local review-page orchestration.
- root `run_archive.py`, `migrate_archive.py`, and
  `serve_archive_review.py`: the only operator-facing Python entries.
- `logging_config.py`: the only logging-handler configuration point; rolling
  log files are written to `logs/`.
- `config_loader.py` and `context.py`: centralized configuration, path
  expansion, and shared entry/flow context.

## Compatibility layer

`modules/legacy.py` temporarily supplies the mature Word parser, llama.cpp
client and PDF matching primitives. The former unified subcommand entry has
been removed; internal operations remain in `flows/` and are reached only
through the three documented root interfaces. Future refactoring can split the
legacy module further without changing the JSON schema.
