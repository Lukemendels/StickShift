---
type: Schema
title: Build Portfolio — Concept Schema
description: Frontmatter and body conventions every build entry in this bundle follows.
timestamp: 2026-06-17T00:00:00Z
---

# Purpose

This bundle is a portfolio tracker: one Markdown file per build (or idea), so
the whole set of things on your plate lives in OneDrive as structured,
queryable files instead of in your head. It is small by design — the entire
index fits in a single DHSChat paste, so you reason over all of it at once and
do **not** need the navigator/retrieval machinery. That stays reserved for
large knowledge bases.

---

# The `Build` concept

Every build is one `.md` file. Filename = a short slug (e.g.
`briefing-builder.md`). The path is the build's identity.

## Frontmatter

```
---
type: Build                       # REQUIRED (OKF)
title: Briefing Builder           # short display name
description: One sentence — what it is, in plain terms.
status: working                   # see status enum below
effort: S                         # effort REMAINING to next ship: S | M | L
impact: high                      # value if shipped: low | med | high
domain: TSA                       # which arena: TSA | FocusFlow | Personal
timestamp: 2026-06-17T00:00:00Z   # last meaningful change
last_touched: 2026-06-17          # date you last actually moved this build
tags: [vba, dhschat]              # optional, cross-cutting themes
---
```

`last_touched` is the stall detector. It exists because the dangerous state for
you isn't a build that's *stopped* — it's one that's nominally `working` but
silently hasn't moved in weeks while sitting at the top of the queue. The index
generator sorts `working` items by `last_touched`, so anything stale floats up
for a deliberate keep-or-`park` decision instead of rotting unnoticed. This turns
a known pattern — start, drift, never formally stop — into something the system
catches instead of something that catches you.

Two precision notes, because they change the answers you get:

- `effort` is **remaining** effort to the next shippable point, not total work
  sunk so far. A build you've poured months into can still be effort `S` if it's
  one bugfix from done. Prioritization cares about what's left, not what's spent.
- `impact` is value **if shipped**, independent of effort. Keeping the two axes
  separate is the whole point — it's what lets "low effort × high impact" surface
  as a category instead of getting blended into a vague "priority" score.

## Status enum (the lifecycle)

| status        | meaning                                                        |
| ------------- | -------------------------------------------------------------- |
| `idea`        | concept only — nothing built                                   |
| `spec`        | spec/design drafted, no working code                           |
| `boilerplate` | scaffold/skeleton exists, not yet functional                   |
| `working`     | functional but with known bugs / not yet stable ("done w/ bugs")|
| `production`  | live, in the hands of real users                               |
| `parked`      | deliberately set aside — reachable, not dead                   |
| `archived`    | killed or superseded (dimmed-but-reachable; never deleted)     |

`parked` and `archived` are the off-ramps — the soft-delete you wanted, so
abandoned threads leave your active view without being destroyed.

## Body sections

```
# What it is
Two or three sentences. Plain language.

# Next action
The single next concrete step. One line. This is externalized working memory,
not a nicety: because slower processing speed makes a context switch *lose* your
place rather than just slow the reload, this line is the state you'd otherwise
have to rebuild from scratch. Write it down before you switch, and the switch
costs minutes instead of erasing where you were. If you can't name the next
action, the build isn't really moving.

# Dependencies
* [Index Generator](/builds/index-generator.md) — needs this before it can run.

# Notes
Anything else worth keeping.
```

The `# Dependencies` links are the engine. They make a build's *blockers*
explicit, and they turn the portfolio into a lightweight graph.

---

# Prioritization queries (paste the index into DHSChat, then ask)

These are the questions the schema is built to answer:

- **Foundational layer (build first):** "Which builds are linked to in the most
  other builds' Dependencies?" High in-degree = many things wait on it = do it
  first. This answers *foundational* mechanically, not by gut.
- **Finish these:** "List builds where status is `working`, impact is `high`, and
  all Dependencies are themselves `production`." Closest to shippable, highest
  payoff, unblocked.
- **Quick wins:** "List builds where effort is `S`, impact is `high`, and no
  Dependency is below `working`." Cheap, valuable, unblocked.
- **Work-in-progress count:** "How many builds are `boilerplate` or `working`?"
  That's your open WIP. If it's high, the instruction to yourself is finish
  before starting — the portfolio makes over-extension visible.

---

# Worked example — `builds/briefing-builder.md`

```
---
type: Build
title: Briefing Builder
description: VBA pipeline that assembles briefing packets from clipboard input.
status: production
effort: S
impact: high
domain: TSA
timestamp: 2026-06-17T00:00:00Z
tags: [vba, dhschat]
---

# What it is
Clipboard-as-primary-ingest tool that builds briefing packets with a
friendly-fail UX. In production with users; 23/23 twin tests passing.

# Next action
Monitor for edge-case failures; no active development.

# Dependencies
(none — self-contained)

# Notes
Reference case for "shipped, in use" — useful as a proof-point in pitches.
```

---

# The OKF system tracking itself (dogfood snapshot)

Each row below is its own file under `builds/`. Statuses reflect where these
stood as of our build session — set the real values yourself; treat these as a
starting draft, not ground truth.

| Build                  | status      | effort | impact | depends on                  |
| ---------------------- | ----------- | ------ | ------ | --------------------------- |
| OKF Build Portfolio    | working     | S      | high   | Index Generator             |
| Index Generator (VBA)  | working     | S      | high   | —                           |
| Retrieval Macro (VBA)  | working     | M      | med    | Index Generator             |
| Navigator Prompt       | working     | S      | med    | Index Generator             |
| Scribe (writer) Prompt | spec        | M      | med    | Write Macro                 |
| Write Macro (VBA)      | idea        | M      | high   | —                           |
| Maintainer Prompt      | idea        | L      | low    | Linter, Write Macro         |
| Linter (VBA)           | idea        | M      | med    | —                           |

Reading this the way the queries would: **Index Generator** is the foundational
layer (three builds depend on it, and it's already `working` — so it's done and
unblocking). **OKF Build Portfolio** itself is the quick win to finish this week
(effort `S`, impact `high`, its one dependency is satisfied). The Maintainer is
correctly last — low impact for now, two unmet dependencies. The graph is
already telling you the sequence.

---

# How to stand it up (Phase 1 — this week)

1. In OneDrive, make a folder (e.g. `build-portfolio/`) with a `builds/`
   subfolder and a `_meta/` subfolder. Drop this file in `_meta/`.
2. Brain-dump every build to DHSChat; ask it to return one schema-conformant
   `.md` file per build. Save them into `builds/`.
3. Run the Index Generator VBA against the folder to produce `index.md`.
4. Paste `index.md` into DHSChat and run the prioritization queries above.
5. Stop. That's the shipped tool. Phase 2/3 only if you outgrow it.

---

# Operating discipline (the rule that makes it work)

The schema is the instrument; this is the rule. Without it, the portfolio is
just a tidier list of things you're still not finishing.

**Cap `working` at one.** Nothing leaves `spec` until the current `working` item
reaches `production`. Not "try to focus" — a hard count of one. Ideas pile up for
free at `idea`; the one scarce resource is finished-execution attention, and
exactly one thing occupies it at a time.

**Capture removes the real reason you switch.** You don't pick up a new project
because the current one is boring — you pick it up because you're afraid of losing
the idea. So the fix isn't willpower, it's capture. The moment an idea is a file
here at `status: idea`, parking it costs nothing: it's on disk, reachable, safe.
Write it down, then go back to the one `working` item. Capture is what makes
single-tasking feel like sequencing instead of abandonment.

**Why the cap speeds things up (Little's Law):** cycle time = WIP ÷ throughput.
The number of things in progress divides into how fast any one finishes. Eight
open builds means each runs roughly eight times slower. Cutting WIP doesn't starve
the portfolio — it gives each item the undivided server until done, so more ships
per quarter, sooner. Doing everything at once is the configuration that
*minimizes* throughput.

**Price in your optimism.** Everything takes longer than the version in your head.
Running one item at a time keeps that optimism from compounding silently across
parallel threads that are all secretly behind.

**Blocked vs. bored — the only legitimate exception.** WIP-of-one is the target,
not a suicide pact. If the active item genuinely stalls on something *outside you*
— a review you're waiting on, data you don't have, a model-access window — pulling
the next queued item is correct. The test: is the work waiting on *the world*, or
waiting on *you to push through discomfort*? Waiting on the world → pull the next
card. Waiting on you → that's the leak the cap exists to plug. And a switch is
never free even when justified: you pay a reload cost rebuilding the project's
mental state. Price that in before you make the jump.

**Two altitudes — don't confuse them.** Your evaluation recommends switching tasks
when you feel overwhelmed or bored. That is *task-level* advice and it's correct:
inside the one active build, if you seize up on a piece, move to a different piece
of *the same build* to keep momentum. The WIP cap is *project-level*: never promote
a second build to `working`. So you have full permission to shuffle micro-tasks
within the active build, and no permission to open a second front. Different layers,
both true — naming the distinction here so future-you doesn't read them as a
contradiction and abandon one of them.

**Run a stall sweep.** Once a week, look at the top of the `working`-sorted-by-
`last_touched` list. Anything that hasn't moved is asking you a question: is this
still the one thing, or has it quietly stalled? If it's stalled, `park` it
honestly rather than letting it sit and block the slot. The point of the cap is a
*live* single item, not a stale one squatting in the only chair.
