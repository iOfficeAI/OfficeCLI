# Agents Project

This small project helps import the provided `agents.zip` into an `agents/` directory.

Usage

- Copy the provided `agents.zip` into this folder (`agents-project/`).
- From this folder run either:

```
sh import_agents.sh
```

or, if you prefer npm:

```
npm run import
```

Result

- The archive will be extracted into the `agents/` directory.

Notes

- The script uses the system `unzip` utility. On macOS you can install it with `brew install unzip` if missing.
- If you want automatic processing or programmatic access, I can add a Node.js importer that uses `unzipper` or `adm-zip`.
