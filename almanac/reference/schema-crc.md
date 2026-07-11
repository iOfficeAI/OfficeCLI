---
title: "Schema CRC"
summary: "Lookup reference for the OfficeCLI schema CRC flag and the exact embedded help-schema bytes it fingerprints."
topics: [reference, cli, schemas]
sources:
  - id: program
    type: file
    path: src/officecli/Program.cs
  - id: schema-crc
    type: file
    path: src/officecli/Help/SchemaCrc.cs
---

Schema CRC is the eight-character lowercase hexadecimal CRC32 value printed by `officecli --output-schema-crc`. It fingerprints the embedded `schemas/help/**` resource tree by hashing each canonical resource name and the raw bytes of that resource in stable sorted order, so downstream automation can detect help-schema surface drift across binary upgrades [@program] [@schema-crc]. It is part of the help-schema contract described in [Help Schemas](../concepts/help-schemas) and loaded through [Help Schema Loader](../architecture/cli/help-schema-loader).

## Command

`--output-schema-crc` is handled in `Program.cs` before normal root command construction [@program]. When it is the only argument, the process prints `SchemaCrc.Compute()` and exits successfully [@program].

## What Is Hashed

`SchemaCrc.Compute()` scans the current assembly manifest resource names, normalizes each name by replacing backslashes with forward slashes and lowercasing it, and keeps only names that start with `schemas/help/` [@schema-crc]. It sorts those entries by canonical name, then feeds the canonical UTF-8 name followed by the resource's raw bytes into the CRC accumulator [@schema-crc].

## Algorithm

The implementation builds the standard reflected CRC32 table with polynomial `0xEDB88320`, starts from `0xFFFFFFFF`, appends bytes by table lookup, XORs the final value with `0xFFFFFFFF`, and formats the result as `x8` lowercase hexadecimal [@schema-crc]. Resource streams are read in chunks of 81,920 bytes [@schema-crc].

## Stability Boundary

The source comment defines the CRC as a fingerprint of the embedded help-schema tree only. It covers schema files and their canonical resource names, but it does not cover serialization behavior implemented in code, JSON field order produced by renderers, command parsing behavior, handlers, or wiki prose [@schema-crc].
