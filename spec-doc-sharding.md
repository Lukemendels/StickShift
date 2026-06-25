---
type: Spec
title: Document Sharding — long-concept splitting for cheaper writes
description: Convention plus two thin runtime changes that let DHSChat break a long concept into a folder of section files (so an edit rewrites one small file, not the whole document), with a deterministic linter check that flags long concepts for sharding. Keeps retrieval scoping sharp.
status: spec
okf_version: "0.1"
---

# Document Sharding

## 0. Model note (DHSChat = GPT-5.1)

The assistant on the other end of the loop is **GPT-5.1**. Two consequences for this spec:

- The skill and the system-prompt text are written as **explicit imperative steps with a
  literal output contract** ("emit exactly one envelope, nothing else"). GPT-5.1 follows literal
  output contracts well but will otherwise pad with preamble, which breaks clean copy of the
  envelope.
- The skill `description` is **keyword-rich** (shard / split / break up / long / section by
  section) so GPT-5.1 selects it from the manifest when you ask.

The reliable trigger in practice is **you asking** ("shard builds/X"); the skill is the
reference GPT-5.1 loads to know how to do it well. The linter (section 5) is the mechanical
backstop that tells *you* when to ask.

## 1. The problem (the friction you actually hit)

`StickShiftWriteApply` replaces **whole files**. Every `### FILE:` block overwrites the entire
target (`ApplyWriteEnvelopeText` writes `fileContents(i)` verbatim; there is no patch, append,
or section-merge path). So when DHSChat edits one heading of a long concept, it must re-emit the
*entire* concept inside `<VBA_WRITE>`. On a 1,200-1,800 line file that is slow to generate (the
tokens/sec drag), bloats the envelope, and raises the odds of a truncation or stray-character
regression.

The fix is structural, not mechanical: **stop letting any single concept get long.** Shard a
long concept into a folder of section files. An edit to one section then re-emits only that
section (tens of lines); the rest of the document is untouched.

## 2. What already exists (so we do not rebuild it)

1. **Granular retrieval.** `BuildContextBundle` in `mode: bundle` pulls *only* the seeds you
   name plus their graph neighborhood at the depth you set. It does **not** "pull the whole
   folder." Scoping is already as sharp as you want.

2. **Folder index = a free hub.** `StickShiftIndexGenerator` writes an `index.md` in *every*
   directory, listing that directory's concepts as **same-folder relative links**
   (`* [Section One](01-overview.md) - ...`). So a sharded folder's `index.md` is already a table
   of contents whose links the bundler can traverse.

Together these give the whole retrieval story for a sharded document with zero new bundler code:

- **Whole document:** seed the folder's `index.md` at `depth: 1`, outbound. The hub
  (`index.md`, layer `map`) plus every section (layer `selected`) return in one hop. Verified
  against `ResolveLinkToRel` / `ExtractLinksScoped`.
- **One section:** seed that section file at `depth: 0`.
- **Edit one section:** one small `### FILE:` block; nothing else is rewritten.

So this is a **convention + a skill + two small runtime changes** (a writer hardening and a
linter check). Deliberately small; the YAGNI-correct shape (see `_meta/okf-roadmap.md`).

## 3. The sharding convention

A sharded document is an ordinary folder under any concept directory. Layout:

```
builds/
  autoreviewer.md            <- STUB (preserves the old path + inbound links)
  autoreviewer/              <- the shard folder
    index.md                 <- AUTO-GENERATED hub (do NOT hand-write; reserved)
    00-overview.md           <- orientation: what the doc is, how the sections fit
    01-architecture.md
    02-write-path.md
    03-open-decisions.md
```

Rules:

- **One section per file.** Each section is a normal concept (frontmatter + body). Its
  `description` is what shows in the hub, so write it as a real one-liner.
- **Zero-padded numeric prefixes** (`00-`, `01-`, ...). A shard folder carries no `status`, so
  the generator lists it **flat alphabetical** (the `skills/`-style path); the numeric prefix is
  what gives reading order.
- **`index.md` is machine-owned - never hand-write it.** The generator overwrites it, and
  `StickShiftWriteApply` skips it on writes. The hub is free.
- **The original path becomes a thin stub, not a deletion.** `builds/autoreviewer.md` stays as a
  short concept whose body points to the shard
  (`See [the full document](/builds/autoreviewer/index.md).`). This preserves its identity and
  keeps inbound links (`/builds/autoreviewer.md` from other files) resolving. **Never delete
  the original.**
- **Cross-links use `/`-rooted paths** in bodies, like every other concept.
- **Depth note:** pulling the stub at `depth: 1` reaches the hub (one layer); the sections are
  two layers from the stub. So the canonical "load it all" seed is the **folder's `index.md`**,
  not the stub. The stub is for link-preservation and human discovery.

### Discovery path

`mode: index` returns the root `index.md`, which lists `builds/` under `# Subdirectories`.
DHSChat pulls `builds/index.md`, sees `autoreviewer/` under *its* Subdirectories, and seeds
`builds/autoreviewer/index.md`. This is the existing navigator pattern - no new mechanism.

## 4. Runtime change A: recursive folder creation (writer)

`StickShiftWriteApply` creates parent directories **one level deep**:

```vba
parentDir = fso.GetParentFolderName(absPath)
If Not fso.FolderExists(parentDir) Then fso.CreateFolder parentDir
```

`FileSystemObject.CreateFolder` cannot create a chain. Writing `builds/autoreviewer/00-overview.md`
when `builds/autoreviewer/` does not yet exist works today (one new level under existing
`builds/`). But `area/topic/00.md` with two missing levels **raises an unhandled runtime error
and halts the apply**. Normal sharding is one level deep, so you will not hit this; the fix is
insurance for the nested case.

Replace the single `CreateFolder` with a recursive helper:

```vba
' Create every missing folder in the chain down to dirPath (absolute, under m_BundleRoot).
Private Sub EnsureFolderTree(ByVal dirPath As String)
    If dirPath = "" Then Exit Sub
    If fso.FolderExists(dirPath) Then Exit Sub
    Dim parent As String
    parent = fso.GetParentFolderName(dirPath)
    If parent <> "" And Not fso.FolderExists(parent) Then EnsureFolderTree parent
    On Error Resume Next
    fso.CreateFolder dirPath
    On Error GoTo 0
End Sub
```

Call site (replace the two lines above):

```vba
parentDir = fso.GetParentFolderName(absPath)
If Not fso.FolderExists(parentDir) Then EnsureFolderTree parentDir
```

ASCII only, no `ChrW` (CI guards). Recursion terminates at the first existing ancestor (the
bundle root always exists - checked before the write loop), so it cannot run away.

### Why no patch/append write mode

A section-patch / append mode would also cut re-emission, but it is a much larger change (a
second envelope grammar, merge semantics, a new failure surface) and weakens the "every write
restates a whole, schema-valid file" invariant. Sharding gets the same win by making the whole
file small. **Defer patch-mode to the roadmap;** revisit only if a single *section* routinely
gets long.

## 5. Runtime change B: a linter check that flags long concepts

Detection is judgment, and GPT-5.1 can only judge length for a file it is holding - it cannot
scan the bundle. So make detection **mechanical and operator-facing**, in the same idiom as the
stall detector (which floats stale `working` builds to the top so you act on them): the linter
measures every build concept deterministically and warns when one is long enough to be worth
sharding. You then trigger the shard.

**Scope: `builds/`, top level, non-recursive** - exactly the existing linter scope. This is
deliberate: `StickShiftLint` only scans `builds/` and its required-field checks
(`status`/`effort`/`impact`) would falsely error on skills or foundation files. `builds/` is
also the primary sharding target. Two consequences, both acceptable for v1: section files inside
a shard subfolder are not scanned (re-sharding a long section is an edge case the operator
notices while editing), and a sharded build's **short stub stops tripping the warning
automatically** once the body moves into the folder.

Change `StickShiftLint.bas` as follows.

Add a configurable threshold near the existing `STALE_DAYS` const (extension seam, same style as
`GROUP_BY_CANDIDATES` in the generator):

```vba
Private Const SHARD_WARN_LINES As Long = 400   ' long-concept advisory threshold
```

Inside `ScanBundle`, declare a collector alongside the existing arrays:

```vba
Dim longArr(0 To 999) As Variant
Dim longCount As Long: longCount = 0
```

Inside the existing per-file loop (where `content = ReadUtf8(f.path)` is already read), after the
required-field checks, add:

```vba
' Long-concept advisory: file is large enough to be worth sharding.
Dim lc As Long: lc = ConceptLineCount(content)
If lc >= SHARD_WARN_LINES Then
    longArr(longCount) = Array(lc, f.Name)
    longCount = longCount + 1
End If
```

In the warnings region (after the stall findings are added, near the active-to-archived block),
add the findings longest-first:

```vba
If longCount > 1 Then SortByLineCountDesc longArr, longCount
Dim li As Long
For li = 0 To longCount - 1
    AddFinding findings, "warning", CStr(longArr(li)(1)), _
        "long concept: " & CStr(longArr(li)(0)) & " lines (>= " & SHARD_WARN_LINES & _
        "); consider sharding (see skills/doc-sharding.md)"
Next li
```

Add two private helpers:

```vba
Private Function ConceptLineCount(ByVal content As String) As Long
    Dim n As String
    n = Replace(Replace(content, vbCrLf, vbLf), vbCr, vbLf)
    ConceptLineCount = UBound(Split(n, vbLf)) + 1
End Function

Private Sub SortByLineCountDesc(ByRef arr() As Variant, ByVal n As Long)
    Dim i As Long, j As Long, tmp As Variant
    For i = 1 To n - 1
        tmp = arr(i): j = i - 1
        Do While j >= 0
            If CLng(arr(j)(0)) >= CLng(tmp(0)) Then Exit Do
            arr(j + 1) = arr(j): j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i
End Sub
```

ASCII only, no `ChrW`. The `WriteReport` sheet already colours `warning` rows amber - no report
change needed. Findings still land under the existing "StickShift Lint Report" sheet.

## 6. The skill: `skills/doc-sharding.md`

Delivered as a standalone file (`doc-sharding.md`) to drop straight into `skills/`. Full content
(identical to the standalone file, reproduced here so this spec is self-contained for the build):

```markdown
---
okf_version: "0.1"
type: Skill
title: Document Sharding
description: Splits one long concept into a folder of smaller section files so future edits rewrite a single section instead of the whole document, and keeps retrieval scoped to what is needed. Use this skill when the operator asks to shard, split, or break up a document, when a concept has grown long, or when the operator keeps editing the same document section by section.
tags: [skill, authoring, sharding]
---

# document-sharding

## Purpose

Convert one long concept file into a folder of smaller section files. After sharding:

- an edit touches one section (tens of lines) instead of forcing a whole-file rewrite,
- the operator can pull one section or the whole document on demand,
- the original path keeps working, so nothing that linked to it breaks.

This is a structural refactor, not new content. You are reorganizing what already exists.

## When to use this skill

Primary trigger: the operator asks to shard, split, or break up a document, or names a file
and says it is getting long.

Also appropriate when, in the course of other work, you are editing the same long document
section by section and re-emitting the whole file each time is wasteful.

You cannot see the bundle and you cannot measure a file you are not holding. So do not wait for
a size signal you cannot get: act when the operator asks, or when a document you are already
holding is clearly long (roughly 400+ lines).

Do not shard short documents. Do not shard the foundation profile or the schema unless asked -
small whole files are easier to reason over and are pulled every session.

## Before you start

You need the target file's full current contents. If you are not holding them, retrieve the
file first (one <CONTEXT_REQUEST>, depth: 0), then run this procedure. You cannot shard a
file you have not read - you would lose its body.

## How to think through the split

1. Read the whole document. Find its natural seams - the top-level # / ## headings are
   usually the section boundaries. Each becomes one file.
2. Right-size the pieces. Aim for sections that are individually short enough to rewrite
   cheaply. Merge small adjacent sections into one file rather than making many tiny files;
   split a single oversized section into two if one heading dominates.
3. Decide the folder name: the original slug with .md removed.
   builds/autoreviewer.md -> folder builds/autoreviewer/.
4. Name sections with zero-padded numeric prefixes in reading order: 00-overview.md,
   01-<slug>.md, 02-<slug>.md, ... The prefix sets the order they list in (a shard folder
   has no status, so it lists flat alphabetical; the numbers are what give it sequence).

## What to write

Produce these files, all in one envelope:

1. 00-overview.md - written first. Two to four sentences on what the document is and how the
   sections fit together, then a short list of /-rooted links to each section. This is the
   human and model entry point into the shard.
2. One file per section - each a normal concept: frontmatter (type carried from the original,
   a title, and a real one-line description, since the description is what shows in the
   auto-generated folder listing) followed by that section's body. Where sections reference each
   other, link with /-rooted paths.
3. The original file, rewritten as a thin STUB - keep its frontmatter, replace the body with a
   single pointer line: See [the full document](/<folder>/index.md). Do not delete the
   original and do not empty its frontmatter: other concepts may link to its path, and the stub
   keeps those links resolving.

Do not write any index.md. The folder's index.md is generated automatically after the write
and lists the sections for you; writing one yourself does nothing (it is overwritten, and the
writer skips it as a reserved file).

## Output

Emit exactly one <VBA_WRITE> envelope containing the stub, 00-overview.md, and every section
file, and nothing else - no preamble, no summary, no prose after the envelope. New folders are
created automatically by the writer from the file paths; you do not issue any separate "create
folder" command - there is none. A folder comes into existence the moment a file is written into
it.

## Retrieving a sharded document afterward

- Whole document, one hop: mode: bundle, include: [<folder>/index.md], depth: 1. The
  folder index lists the sections as links, so the index and all sections come back together.
- One section: request that section file at depth: 0.
- One edit later: re-emit only the section file you changed. That is the whole point - the rest
  of the document stays untouched.

## Example

Operator: "builds/autoreviewer.md is huge, shard it." You retrieve it if you are not already
holding it. It has headings: Overview, Architecture, Write path, Open decisions. You emit one
envelope writing: builds/autoreviewer.md (stub pointing to /builds/autoreviewer/index.md),
builds/autoreviewer/00-overview.md, 01-architecture.md, 02-write-path.md,
03-open-decisions.md. After apply, builds/autoreviewer/index.md appears on its own. Next time
the operator wants the whole thing, you seed builds/autoreviewer/index.md at depth: 1.

## Edge cases

- A single section is itself very long: prefer re-splitting that section's headings more finely
  first; shard recursively (a sub-folder) only if it is genuinely a document of its own.
- Uneven headings: group small adjacent sections; do not create one-line files.
- The operator only wants part of it sharded: shard just the named sections; leave the rest in
  the stub's body instead of an empty pointer.
```

## 7. System-prompt additions (`okf-context-assistant.md`)

These are **additive blocks** for your current (edited, sidecar-free) instructions file - merge
them in by meaning, not by line number. Optional but recommended, so GPT-5.1 knows the
convention before it loads the skill. The skill remains the primary mechanism.

**In the RETRIEVE section, add a rule:**

> - A sharded document is a folder. To load it whole, seed the folder's `index.md` at `depth: 1`
>   (the hub lists its sections as links, so all sections return in one hop). To load one
>   section, seed that section file at `depth: 0`.

**In the WRITE section, add a rule:**

> - Prefer sharding over an ever-growing file. If a concept has grown long (roughly 400+ lines)
>   and you are editing it repeatedly, propose splitting it into a folder of section files (the
>   Document Sharding skill gives the exact layout) instead of re-emitting the whole file each
>   time. A new folder is created automatically when a `### FILE:` path names one - there is no
>   separate create-folder step.

## 8. Non-goals (per the roadmap discipline)

- **No `mode: folder` bundler addition.** Whole-folder pull already works via the folder's
  `index.md` at depth 1. (Roadmap candidate.)
- **No patch/append write mode.** See 4. (Roadmap candidate.)
- **No recursive / all-directory linting.** The long-concept check stays `builds/`-only, like
  the rest of the linter. Extend coverage only if non-build concepts start needing shards.
- **No auto-sharding.** The split is a deliberate, human-gated action triggered by the skill or
  by you, never something the writer does on its own.

## 9. Tests / CI

- **Linter twin (`tests/test_okf_lint.py`) - update the model and add vectors.** Add a
  `shard_warn_lines: int = 400` parameter to `lint_bundle`; in its per-file phase, compute the
  line count and append `Finding("warning", f.name, "long concept: N lines (>= T); consider
  sharding")` when `line_count >= shard_warn_lines`. Line count = normalize CRLF, then
  `content.count("\n") + 1` (matches the VBA `UBound(Split(...)) + 1`). Add tests:
  - over threshold -> one `warning` finding on that file (use a small `shard_warn_lines` so the
    fixture file stays small);
  - under threshold -> no long-concept finding;
  - inclusive boundary -> a file of exactly T lines is flagged, the same file with threshold
    T+1 is not.
  Existing assertions are unaffected: the default 400 leaves the tiny clean/broken fixtures
  untouched, so `test_clean_bundle_produces_no_findings` and the broken-bundle expectations
  still hold.
- **Writer change (recursive folder creation) is not twin-testable** in the current pure-logic
  harness (the parse/compute twin never touches the filesystem). Verify by **manual test**:
  apply a `<VBA_WRITE>` whose `### FILE:` path is two-levels-new (`a/b/c.md`); confirm both
  folders are created, the file lands, and `log.md` updates.
- Run `pytest tests/ -q` (all existing tests must still pass). ASCII guard + ChrW guard must
  pass on the edited `StickShiftWriteApply.bas` and `StickShiftLint.bas`.

## 10. Build prompt for Claude Code

> Read `_meta/spec-doc-sharding.md` in this repo and implement it. Scope:
>
> 1. `builds/StickShiftWriteApply.bas`: add the private `EnsureFolderTree` helper from section 4
>    and replace the single-level parent-folder creation at its call site with a call to it. Do
>    not change the `<VBA_WRITE>` grammar, the `index.md`/`log.md` write guard, the logging, or
>    the auto-reindex call. `ApplyWriteEnvelopeText` / `ApplyStickShiftWrite` public contract
>    unchanged. Pure ASCII, no `ChrW`.
> 2. `builds/StickShiftLint.bas`: add the `SHARD_WARN_LINES` const, the long-concept collection
>    inside the existing `builds/` per-file loop, the longest-first findings in the warnings
>    region, and the `ConceptLineCount` / `SortByLineCountDesc` helpers, exactly as in
>    section 5. Keep the linter `builds/`-only and non-recursive. Do not touch the other checks
>    or `WriteReport`. Pure ASCII, no `ChrW`.
> 3. `tests/test_okf_lint.py`: extend the `lint_bundle` twin with a `shard_warn_lines=400`
>    parameter and the long-concept finding, and add the three tests in section 9 (over
>    threshold, under threshold, inclusive boundary). All existing tests must still pass
>    unchanged.
> 4. Create `skills/doc-sharding.md` with the exact content in section 6 (it is also provided as
>    a standalone file). New knowledge-base file at the repo root under `skills/`; do not embed
>    it in any module.
> 5. Apply the two `okf-context-assistant.md` inserts in section 7 by meaning (the file has been
>    edited locally; merge additively, do not assume specific surrounding text or rule numbers).
>
> Constraints: `pytest tests/ -q` passes with no changes to existing tests/golden vectors beyond
> the additive linter cases. ASCII + ChrW guards pass on both edited `.bas` files. Do not add a
> patch/append write mode, a `mode: folder` bundler path, or recursive linting - all explicitly
> out of scope (section 8). Summarize the diff and give the exact Excel manual test for
> two-level-deep folder creation.
