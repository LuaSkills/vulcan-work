# Vulcan Work debugging

The wrappers download the pinned official LuaSkills debugger, verify its checksum, and prepare the local managed Python runtime. The debugger stages the source in its ignored `.luaskills-debug/` workspace.

Run verification from the repository root. On Windows, set `VULCAN_WORK_DEBUGGER` to the full path of an installed matching debugger to exercise that executable with the repository's isolated runtime.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify_skill.ps1
```

```bash
bash scripts/verify_skill.sh
```

Verification loads and enumerates the entries, runs real calls with synthetic Office fixtures, and validates and packages the skill. It writes one standard ZIP and its SHA-256 check file to `dist/`.

For a single invocation, save an argument file with an absolute path:

```json
{"path":"D:/work/report.pdf","start_page":1,"end_page":1}
```

```powershell
.\scripts\debug_skill.ps1 -Command call -Tool vulcan-work-pdf-extract-text -ArgsFile .\args.json -Output json
```

```bash
bash scripts/debug_skill.sh call --tool vulcan-work-pdf-extract-text --args-file ./args.json --output json
```

Use the JSON envelope to verify the actual tool result. The `inspect` and `list-tools` commands report the loaded manifest and registered names. For help and examples about office formats, see `help/`.
