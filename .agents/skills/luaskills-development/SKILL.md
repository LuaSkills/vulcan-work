---
name: luaskills-development
description: Edit, debug, verify, or package the Vulcan Work LuaSkills package.
---

# Vulcan Work development

1. Keep the package version in `skill.yaml`. Update it only when the user requests a version change. The debugger tag is pinned separately in `.luaskills-version`.
2. Read the entry in `skill.yaml`, its schema, the Lua wrapper, the Python handler, and the matching help topic before editing. Keep the tool and its documented examples consistent.
3. Use the locked managed Python dependencies in `dependencies.yaml` and `python/requirements.lock`. Keep the published ZIP's third-party licenses and notices aligned with the dependencies and bundled source.
4. Write English first and Chinese second in code comments and docstrings. Explain the purpose of functions, their inputs and outputs, and non-obvious choices.
5. Run the official debugger and the verification workflow in [the debugging guide](references/ai-debugging.md). Confirm the resulting ZIP has one `vulcan-work/` root with the manifest, schemas, runtime, help, Python sources, and licenses.
6. Publish only on explicit request. The release scripts rerun verification and use the exact tag corresponding to `skill.yaml`.
