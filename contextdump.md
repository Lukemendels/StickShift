# OKF Seed Packet — TSA Work Portfolio

This packet is the initial content for your OKF knowledge base. It is organized into the
three bundle directories (`_foundation/`, `skills/`, `builds/`) and delimited file-by-file so
you can port each concept straight into the bundle.

**Scope:** TSA work only. Personal ventures, career-advancement specifics (band/promotion), and
internal-politics characterizations are deliberately excluded — this is a professional
knowledge base your assistant reasons over inside DHSChat. Where a fact needs your confirmation
or detail, you'll see `[CONFIRM]` or `[FILL]`.

---

## How to port this into DHSChat

1. **Set the bundle root.** In the dashboard, click **Set Bundle Root** and pick your OneDrive
   bundle folder. Create the three subfolders: `_foundation/`, `skills/`, `builds/`.
2. **Create each file.** For every `FILE:` block below, create a file at that path and paste
   the content between the delimiters (the YAML frontmatter and the body — not the delimiter
   line itself).
3. **Do NOT create `index.md` files.** They are generated. After the concept files exist, click
   **Regenerate Indexes** — it writes `index.md` in the root and in every subfolder, including
   the `skills/index.md` manifest that DHSChat reads in Hop 1.
4. **Lint.** Click **Run Linter.** Expect one finding (see the WIP note below); otherwise it
   should be clean — every dependency link in this packet resolves to a file in the packet.
5. **Start a session.** With nothing on the clipboard, click **Build Context Bundle** → it
   defaults to `index` mode and writes `OKF-context.md`. Attach that file to DHSChat.

**WIP note (expected lint finding):** this packet has **two builds at `status: working`** —
`okf-system` and `air-cargo-detail` — reflecting your real dual focus. The linter's
work-in-progress check will flag this on first run. That is the linter working as designed.
Decide which is your single WIP and set the other to `parked`, or accept two if that's your
reality. I set it honestly rather than tidying it away.

---

## File manifest

**`_foundation/`** (always loaded in Hop 1 — keep tight and stable)
- `00-operating-profile.md` — role, the approved stack, how you work, key collaborators
- `10-operating-principles.md` — AHA, constraints-as-governance, the build philosophy
- `20-glossary.md` — pinned terms

**`skills/`** (discovered via manifest in Hop 1; bodies pulled on demand)
- `document-review.md` · `house-voice-drafting.md` · `persona-training.md`
- `briefing-assembly.md` · `narrative-synthesis.md` · `capability-conops.md`
- `org-mapping.md` · `tool-handoff.md`

**`builds/`** (the portfolio graph)
- `okf-system.md` · `autoreviewer.md` · `briefingbuilder.md` · `auto-modeler.md`
- `ingestion-pipeline.md` · `dhschat-memory.md` · `legibility-pipeline.md`
- `operation-paperclip.md` · `ai-look-book.md` · `air-cargo-detail.md`

═══════════════════════════════════════════════════════════════════════════════
FILE: _foundation/00-operating-profile.md
═══════════════════════════════════════════════════════════════════════════════
---
type: Foundation
title: Operating Profile
description: Who the operator is, the approved stack the assistant works within, how the operator works, and the key collaborators relevant to the work.
---

# Role
A TSA economist in the Strategy directorate [CONFIRM exact office name/acronym] and a DHS
Generative AI Ambassador. The work spans two lanes: economic regulatory analysis (Regulatory
Impact Analyses and related products) and building governed AI tooling that the agency can run
on the tools it already has.

# The approved stack (the runtime)
- **DHSChat (GPT-5.1)** — the reasoning engine. Approved, but isolated: no native tool-calling,
  no file access, no API.
- **Excel VBA** — the deterministic actuation layer. Macros are the tools.
- **OneDrive / SharePoint** — storage for knowledge and generated outputs.
- **Power Automate** — event-driven I/O at the edges (e.g. email ingestion).
- **GitHub** — code version control; pulled into the environment by download (no local git client).
- Environment is GCC-High. No approved Python/R runtime [CONFIRM]; Office Scripts is the
  intended in-Excel compute migration target. The working shorthand for the architecture is the
  "paperclip" pattern: clipboard-as-human-bus, no LLM API, no Python orchestration.

# How the operator works
- Build governed capability from the substrate the institution already has. Constraints
  constitute governance rather than impeding it.
- On details and rollouts: **listen → build → transfer.** Map the real work first, build a
  prototype, then hand the tool to an internal owner.
- **WIP-of-one** discipline: one active build at a time; stalled work surfaces for a keep/park
  decision rather than drifting.
- Demo-first; receipts over pitches.

# Interaction preferences (for the assistant)
- Lead with the answer; be terse. Examples help. Offer frameworks.
- Externalize state and reasoning so it can be inspected.
- Push back honestly — flag risks and weak points; do not flatter or rubber-stamp.
- Flag when a concept is foundational enough to deserve separate study.

# Key collaborators (roles only — relevant to the work)
- **Katie** — direct supervisor / section chief.
- **Jim** — branch chief economist; the editorial authority whose review standard the
  document-review tooling encodes.
- **Heather Royce** — sponsor of the Air Cargo detail (the vehicle that puts AI work formally
  in the operator's duties).
- **Paul** — Enterprise Support; co-lead of Operation Paperclip.
- **Genevieve** — IT / AI program management; the route by which Paperclip reaches the CAIO.
- **Pete Griffith** — peer in the Strategy directorate; advocate for approved Python/R access.
- **ACD** — collaborator on the legibility (work-record) pipeline.

═══════════════════════════════════════════════════════════════════════════════
FILE: _foundation/10-operating-principles.md
═══════════════════════════════════════════════════════════════════════════════
---
type: Foundation
title: Operating Principles
description: The governing methodology behind the operator's AI work — the human-in-the-loop pattern, constraints-as-governance, the analytical lens, and the build philosophy.
---

# AHA — the governing loop
**AI surfaces → Human decides → AI executes → Human verifies.** A human is in every operational
loop by design. In a trust-based federal environment, the manual checkpoint is the governance
feature, not a limitation. No action reaches a system of record, an external party, or a final
decision without human approval.

# Constraints constitute governance
Build on approved tools only; zero added procurement; plain-text, auditable artifacts; no
autonomous external action. The thing that disqualifies a system as "fully autonomous" is the
same thing that makes it deployable on a locked-down stack.

# Analytical lens — institutional economics
Design AI around how the organization *actually* works — its real workflows and informal
structure — not around the org chart. Treat agent design as institutional design: name the
principal-agent problems and design for them.

# Model-selection discipline
Match the model and the effort to the task. Minimize cost per unit of judgment. Reserve scarce,
expensive capacity (and scarce human attention) for the steps that need it.

# Friction allocation
Autonomy is earned by paying the alignment cost up front. The human holds the
non-parallelizable judgment; the machine takes the parallelizable, accelerable work.

# Build philosophy
- Build durable **freight trains, not Formula One cars** — robust over flashy.
- **Multiplication over personal productivity**; **pull, not push**.
- **Build it, prove it, hand over the keys, move on.**
- Demo-first. A live ten-minute run beats any one-pager.

═══════════════════════════════════════════════════════════════════════════════
FILE: _foundation/20-glossary.md
═══════════════════════════════════════════════════════════════════════════════
---
type: Foundation
title: Glossary
description: Pinned terms used across this knowledge base.
---

# Terms
- **DHSChat** — the approved enterprise AI chat (GPT-5.1 wrapper); no API/tools/file access.
- **OKF** — Open Knowledge Format; the open Google Cloud standard this knowledge base conforms to.
- **Bundle** — the OKF knowledge base: a tree of markdown concept files (`_foundation/`, `skills/`, `builds/`).
- **Concept** — one markdown file with YAML frontmatter (always a `type`) and a body.
- **Build / Skill / Foundation** — the three concept types in this bundle.
- **Hop** — one manual clipboard round-trip between the model and the VBA tools.
- **CONTEXT_REQUEST** — the structured block DHSChat emits to pull context; the VBA "tool call."
- **VBA_WRITE** — the envelope DHSChat emits to propose a change; applied behind a human gate.
- **Persona** — a trained reviewer/author voice (named by role and document type, not by person).
- **Hot / cold** — the co-thinker (interactive) vs. serializer (mechanical transcription) split.
- **AR_ anchor** — a stable bookmark id used to apply tracked changes at a verified location.
- **RIA** — Regulatory Impact Analysis. **BLUF** — Bottom Line Up Front.
- **Paperclip pattern** — clipboard-as-human-bus architecture: no LLM API, no Python orchestration.
- **GCC-High** — the government cloud environment constraint.

═══════════════════════════════════════════════════════════════════════════════
FILE: skills/document-review.md
═══════════════════════════════════════════════════════════════════════════════
---
type: Skill
title: Document Review (persona-based)
description: Use when reviewing a drafted document against a trained reviewer's editorial standard — not for first-draft generation. Produces anchored, individually rejectable tracked-change edits with rationale. The review-mode procedure behind AutoReviewer / PreFlight.
tags: [review, editing, autoreviewer]
---

# When to use
A draft needs a critical editorial pass in a known reviewer's voice (e.g. the chief
economist's standard) before it goes up the chain — to make cycle two into cycle one.

# Procedure
1. Work from a copy; never touch the original. The last mile is always a human in Word.
2. Read the exported document text with its bookmark index (every paragraph/cell/footnote
   carries a stable `AR_` anchor).
3. Apply the reviewer persona's editorial heuristics (structure first, then line edits).
4. Emit edits as JSONL keyed to anchors — each as `replace_text`, `delete_element`, or
   `add_comment_only`, with a one-line rationale.
5. Anchor every edit to a verified location or refuse it — never guess a position.
6. VBA applies the JSONL as tracked changes authored "AutoReviewer," individually rejectable.

# Inputs / Outputs
In: the draft + the reviewer persona. Out: anchored tracked-change edits + rationale + a run log
(what changed, what was skipped and why).

# Notes
[FILL: the specific editorial rules of the target persona live in that persona's profile, not here.]

═══════════════════════════════════════════════════════════════════════════════
FILE: skills/house-voice-drafting.md
═══════════════════════════════════════════════════════════════════════════════
---
type: Skill
title: House-Voice Drafting
description: Use when translating AI output or rough notes into a finished draft already in the TSA econ house voice and RIA form, so the human reviews instead of wordsmithing. The Drafter / Prose Writer procedure.
tags: [drafting, writing, ria, autoreviewer]
---

# When to use
You have content (AI-generated or notes/data) and want it rendered in final house voice and
house form — not a raw draft you'll have to rewrite.

# Procedure
1. Determine mode. **Translate**: an existing draft is the source; surface every dropped claim
   as a ratification item (a Claim Ledger). **Compose**: notes/data are the complete universe.
2. Draft in the target authorial persona's voice, with structural elements over loose prose.
3. **Never write computed numbers.** Numbers come from VBA via calc placeholders bound to
   spreadsheet cells; the model writes the prose and the bindings, not the arithmetic.
4. Produce annotated intermediate prose with inline binding tags, so a separate mechanical pass
   transcribes it into final form without inference.
5. Flag gaps rather than inventing content.

# Inputs / Outputs
In: source material + the authorial persona + (for numbers) the data cells. Out: a finished
draft in house form + a ledger of any claims dropped or flagged.

# Notes
Format target is the TSA RIA standard [FILL: specific RIA formatting rules].

═══════════════════════════════════════════════════════════════════════════════
FILE: skills/persona-training.md
═══════════════════════════════════════════════════════════════════════════════
---
type: Skill
title: Persona Training (corpus → reviewer/author voice)
description: Use when building a new reviewer or author persona from a corpus of redlined documents or finished exemplars. A rare, offline map-reduce procedure that yields a reusable persona. Distinct from routine review/drafting.
tags: [training, persona, map-reduce, autoreviewer]
---

# When to use
You want to capture a specific person's editorial or authorial style as a reusable, portable
persona (named by role and document type, not by the person).

# Procedure
1. **Map (automated, per document).** For each redlined doc: pick the target author; keep only
   that author's revisions/comments; accept prior non-target revisions as the baseline; stamp
   anchors; extract structured records into a corpus file.
2. **Reduce (human-in-the-loop, multi-pass through DHSChat).** Pass 1: cluster revisions into
   pattern categories. Pass 2: extract a heuristic per category. Pass 3: synthesize the style
   profile.
3. Assemble the profile into a ready-to-use persona (style rules + the JSON output contract),
   and register it (persona name, corpus, profile, document count, last updated).

# Inputs / Outputs
In: N redlined/exemplar documents. Out: a registered, reusable persona.

# Notes
Keep the stable frame separate from the per-persona profile — only the profile changes per
persona.

═══════════════════════════════════════════════════════════════════════════════
FILE: skills/briefing-assembly.md
═══════════════════════════════════════════════════════════════════════════════
---
type: Skill
title: Briefing Assembly (deck generation)
description: Use when turning source material into a leadership briefing or slide deck. Produces strictly-schema'd JSONL slide data for the BriefingBuilder VBA pipeline, which renders it into a formatted PowerPoint. The output is parsed by a literal string matcher — formatting rules are absolute.
tags: [briefing, slides, powerpoint, briefingbuilder]
---

# When to use
A leadership briefing or deck is needed from existing source material, destined for the
BriefingBuilder workbook.

# Procedure
1. Confirm the meta: purpose, audience, length.
2. Generate slide content as JSON Lines, one object per slide, using the exact slide-type
   schema. The parser matches the literal pattern `"Key":"` with zero whitespace — a single
   stray space silently drops the field and the build fails with no error.
3. No preamble, no code fences, no smart quotes; every line begins with `{`.
4. The operator copies the JSONL to the clipboard; VBA parses it to a slide table and builds the
   deck from the TSA template; friendly-fail logging captures any issue.

# Inputs / Outputs
In: source material + meta (purpose/audience/length). Out: schema-clean JSONL → a formatted deck.

# Notes
The exact slide-type keys are defined by the BriefingBuilder parser contract; follow it verbatim.

═══════════════════════════════════════════════════════════════════════════════
FILE: skills/narrative-synthesis.md
═══════════════════════════════════════════════════════════════════════════════
---
type: Skill
title: Narrative Synthesis (bulk qualitative compression)
description: Use when compressing a large volume of qualitative narratives or free-text records into a synthesized analytic artifact (themes, categories, BLUF). The method behind the Look Book's anchor result (16k narratives, six months of manual work compressed to two days).
tags: [nlp, synthesis, analysis, qualitative]
---

# When to use
You have many free-text records (narratives, comments, survey text) and need structured
findings out of them quickly.

# Procedure
1. Chunk the source into LLM-sized units (see the ingestion pipeline) and carry stable record ids.
2. Map: classify/extract per chunk against a defined coding scheme.
3. Reduce: aggregate into themes/categories with counts; preserve traceability to source ids.
4. Produce a BLUF synthesis with the supporting structure beneath it.

# Inputs / Outputs
In: the corpus + a coding scheme. Out: themes/categories with counts + a BLUF, traceable to source.

# Notes
Keep the coding scheme explicit and versioned so the synthesis is reproducible and defensible.

═══════════════════════════════════════════════════════════════════════════════
FILE: skills/capability-conops.md
═══════════════════════════════════════════════════════════════════════════════
---
type: Skill
title: Capability CONOPS (leadership-facing)
description: Use when writing a Concept of Operations or capability document for a non-technical leadership audience. Translates technical capability into mission value, operation, governance, and a responsible adoption path. Operational, not a sales pitch.
tags: [conops, writing, leadership, enablement]
---

# When to use
A capability needs to be explained to executives — what it does, why it matters, how it
operates, how it is governed, and what adoption looks like.

# Procedure
1. Lead with mission value; translate every technical term into operational language.
2. Cover: problem → capability → how it works (the human/AI division of labor) → governance and
   auditability → security posture → benefits → honest limitations → maturity → recommendation.
3. Foreground the human-in-the-loop governance and the open-standard / approved-stack /
   zero-procurement story — the risk-and-longevity points executives buy.
4. Be honest about the ceiling; honesty reads as credibility in a trust-based room.
5. Leave classification markings and authorship for the operator to set per policy.

# Inputs / Outputs
In: the capability + the audience. Out: a structured, executive-readable CONOPS.

# Notes
Architecture-first ordering suits a technical room; value-first ordering suits leadership — pick
the order to the room.

═══════════════════════════════════════════════════════════════════════════════
FILE: skills/org-mapping.md
═══════════════════════════════════════════════════════════════════════════════
---
type: Skill
title: Org Mapping via Dashboard
description: Use in the early phase of a detail or new assignment to surface a unit's real workflow and informal structure by building a dashboard people actually want. The listen-phase method (Month 1 of listen-build-transfer).
tags: [detail, dashboard, mapping, methodology]
---

# When to use
You're entering a new unit and need to understand how the work actually flows before building
anything for it.

# Procedure
1. Pick a small, genuinely useful artifact the unit wants (a status dashboard, a tracker).
2. Build it from the data and processes they already use; let the act of building it reveal who
   owns what, where work queues, and how decisions actually move.
3. Treat the dashboard as the listening instrument, not the deliverable — the map of the real
   workflow is the output.

# Inputs / Outputs
In: access to the unit's data and people. Out: a working dashboard + an accurate model of the
unit's real workflow and ownership.

# Notes
Keep observations factual and method-focused; do not record sensitive internal-politics reads
into the knowledge base.

═══════════════════════════════════════════════════════════════════════════════
