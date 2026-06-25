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
file first (one `<CONTEXT_REQUEST>`, `depth: 0`), then run this procedure. You cannot shard a
file you have not read - you would lose its body.

## How to think through the split

1. Read the whole document. Find its natural seams - the top-level `#` / `##` headings are
   usually the section boundaries. Each becomes one file.
2. Right-size the pieces. Aim for sections that are individually short enough to rewrite
   cheaply. Merge small adjacent sections into one file rather than making many tiny files;
   split a single oversized section into two if one heading dominates.
3. Decide the folder name: the original slug with `.md` removed.
   `builds/autoreviewer.md` -> folder `builds/autoreviewer/`.
4. Name sections with zero-padded numeric prefixes in reading order: `00-overview.md`,
   `01-<slug>.md`, `02-<slug>.md`, ... The prefix sets the order they list in (a shard folder
   has no `status`, so it lists flat alphabetical; the numbers are what give it sequence).

## What to write

Produce these files, all in one envelope:

1. `00-overview.md` - written first. Two to four sentences on what the document is and how the
   sections fit together, then a short list of `/`-rooted links to each section. This is the
   human and model entry point into the shard.
2. One file per section - each a normal concept: frontmatter (`type` carried from the original,
   a `title`, and a real one-line `description`, since the description is what shows in the
   auto-generated folder listing) followed by that section's body. Where sections reference each
   other, link with `/`-rooted paths.
3. The original file, rewritten as a thin STUB - keep its frontmatter, replace the body with a
   single pointer line: `See [the full document](/<folder>/index.md).` Do not delete the
   original and do not empty its frontmatter: other concepts may link to its path, and the stub
   keeps those links resolving.

Do not write any `index.md`. The folder's `index.md` is generated automatically after the write
and lists the sections for you; writing one yourself does nothing (it is overwritten, and the
writer skips it as a reserved file).

## Output

Emit exactly one `<VBA_WRITE>` envelope containing the stub, `00-overview.md`, and every section
file, and nothing else - no preamble, no summary, no prose after the envelope. New folders are
created automatically by the writer from the file paths; you do not issue any separate "create
folder" command - there is none. A folder comes into existence the moment a file is written into
it.

## Retrieving a sharded document afterward

- Whole document, one hop: `mode: bundle`, `include: [<folder>/index.md]`, `depth: 1`. The
  folder index lists the sections as links, so the index and all sections come back together.
- One section: request that section file at `depth: 0`.
- One edit later: re-emit only the section file you changed. That is the whole point - the rest
  of the document stays untouched.

## Example

Operator: "builds/autoreviewer.md is huge, shard it." You retrieve it if you are not already
holding it. It has headings: Overview, Architecture, Write path, Open decisions. You emit one
envelope writing: `builds/autoreviewer.md` (stub pointing to `/builds/autoreviewer/index.md`),
`builds/autoreviewer/00-overview.md`, `01-architecture.md`, `02-write-path.md`,
`03-open-decisions.md`. After apply, `builds/autoreviewer/index.md` appears on its own. Next time
the operator wants the whole thing, you seed `builds/autoreviewer/index.md` at `depth: 1`.

## Edge cases

- A single section is itself very long: prefer re-splitting that section's headings more finely
  first; shard recursively (a sub-folder) only if it is genuinely a document of its own.
- Uneven headings: group small adjacent sections; do not create one-line files.
- The operator only wants part of it sharded: shard just the named sections; leave the rest in
  the stub's body instead of an empty pointer.
