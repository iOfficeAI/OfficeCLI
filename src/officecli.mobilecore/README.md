# OfficeCLI MobileCore

An experimental in-process facade for Android/iOS hosts. It references the
existing OfficeCLI assembly but exposes only document-local operations:

- open an OOXML file from app-private storage;
- render its HTML preview;
- read outline/get/query data;
- set, add, and remove document elements;
- save back to the same file.

It intentionally excludes CLI parsing, MCP stdio, named pipes, watch servers,
plugins, installers, browsers, subprocesses, raw package writes, and arbitrary
output paths.

The AI layer should deserialize function calls into `MobileOfficeCommand` and
invoke `Execute`; it should never receive a shell or CLI command-string tool.
