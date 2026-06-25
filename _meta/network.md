---
okf_version: "0.1"
type: Spec
title: Text-based CRM (Network) Specification
description: Locked schema and conventions for storing network and contact information in OKF, so tools can parse and visualize it reliably.
owner: Luke Mendelsohn
component: TSA
---

# Purpose

This file defines the **fixed, tool-facing specification** for your text-based CRM inside the
bundle. HTML tools (e.g., a network viewer) should treat this as the canonical schema and layout.

The goals:

- One consistent place for people and interaction history.
- A simple, markdown-only format that DHSChat and HTML tools can both read and write.
- Stable field names and section headings so parsers can be deterministic.

# Directory layout

All CRM data lives under the top-level `/network` folder.

- `/network/index.md`
  - Simple listing of the spec file and the `contacts/` folder.
- `/network/network.md`
  - This specification file (the one you are reading).
- `/network/contacts/`
  - One file per person.
  - Filenames are slugged, lower-case, words separated by hyphens (no spaces), for example:
    - `katie-lastname.md`
    - `jim-lastname.md`
    - `heather-royce.md`

Tools should:

- Treat `/network/contacts/` as the authoritative source of contact records.
- Never write or modify `network.md` except via deliberate spec changes.

# Concept types

Within `/network`:

- `network/network.md`  
  - `type: Spec` — this file.
- Each contact file under `/network/contacts/`  
  - `type: Contact` — one person per file.

No other `type` values in the `/network` tree should be assumed by tools unless the spec is
explicitly updated.

# Contact file schema

Each contact file in `/network/contacts/` must have:

## Frontmatter (YAML)

Required fields:

- `type: Contact`  
  - Fixed literal; used by tools to identify contact records.
- `name`  
  - Full name as displayed (e.g., "Katie Lastname").
- `role`  
  - Primary role or title relevant to you (e.g., "Section Chief").
- `component`  
  - Organizational component (e.g., "TSA").
- `email`  
  - Work email address. If unknown, set to an empty string `""`.
- `last_contact`  
  - Date of the most recent interaction, ISO format `YYYY-MM-DD`.  
  - If unknown, omit the field entirely (do not use null).
- `relationship`  
  - Short descriptor of how they relate to you (e.g., `supervisor`, `peer`, `sponsor`).
- `tags`  
  - YAML list of short, machine-friendly tags, for example:
    - `tags: [econ, leadership]`

Optional fields (tools should preserve but not require):

- `office`  
  - Office or division (e.g., "Strategy Policy and Engagement (SPE)").
- `phone`  
  - Work phone number.
- `location`  
  - Physical location or primary duty station.
- Any additional fields the operator adds in frontmatter.  
  - Tools must preserve unrecognized fields without modification.

Example frontmatter:

```yaml
---
type: Contact
name: Katie Lastname
role: Section Chief
component: TSA
office: Strategy Policy and Engagement (SPE)
email: katie.lastname@tsa.dhs.gov
last_contact: 2026-06-25
relationship: supervisor
tags: [econ, leadership]
---
```

## Body (markdown)

Contact body structure is standardized into three headings:

1. `# Summary`  
   - One or a few paragraphs describing:
     - Who this person is to you.
     - What they care about.
     - How they fit into your work (Econ, Air Cargo AI, broader TSA context).

2. `# Notes`  
   - Bulleted or paragraph notes containing:
     - Preferences.
     - Key facts or reminders.
     - Anything that does not fit neatly into the interaction log.

3. `# Interaction log`  
   - Chronological list of interactions, newest at the top or bottom (your choice; tools should
     not assume sort order unless they enforce it).
   - Each entry is a markdown bullet with the following pattern:

     ```markdown
     - YYYY-MM-DD: Short description of the interaction.
     ```

   - Example:

     ```markdown
     # Interaction log
     - 2026-06-25: Met to discuss the 24-hour airport shutdown cost estimate repack tasker.
     - 2026-06-20: Quick Teams chat about upcoming econ requests for Q3.
     ```

Tools should:

- Recognize and parse these headings literally:
  - `# Summary`
  - `# Notes`
  - `# Interaction log`
- Treat any additional headings or sections as freeform content and preserve them on write.

# Conventions and parsing rules

For HTML tools and other automation:

1. **Identity and lookup**
   - Use the file path `/network/contacts/<slug>.md` as the stable contact id.
   - Use `frontmatter.name` for display.
   - Example id: `network/contacts/heather-royce.md`.

2. **Dates**
   - All dates in frontmatter (`last_contact`) and interaction log entries use ISO format:
     - `YYYY-MM-DD`
   - Interaction log bullet format has:
     - Hyphen bullet (`-`)
     - Space
     - Date
     - Colon
     - Space
     - Description

3. **Tags**
   - Tags are short, lowercase or kebab-case strings.
   - Tags can be used for filtering (e.g., `econ`, `air-cargo`, `leadership`, `ai-champion`).

4. **Unknown or extra fields**
   - Tools must not delete or rewrite unknown frontmatter fields.
   - When updating a contact, tools should:
     - Read existing frontmatter.
     - Modify only the fields they intend to change.
     - Preserve all others as-is.

5. **File creation**
   - New contacts:
     - Must be created under `/network/contacts/`.
     - Must conform to the frontmatter schema above (all required fields present).
     - May have empty `Summary`, `Notes`, and `Interaction log` sections initially.

6. **Index behavior**
   - `/network/index.md` is a simple listing, not a source of truth.
   - Tools should derive the contact list by scanning `/network/contacts/` for `type: Contact`.

# Example contact file (full)

```markdown
---
type: Contact
name: Heather Royce
role: Air Cargo Policy sponsor
component: TSA
office: Air Cargo Policy
email: heather.royce@tsa.dhs.gov
last_contact: 2026-06-20
relationship: sponsor
tags: [air-cargo, leadership, ai]
---

# Summary
Sponsor of the Air Cargo AI detail and a key stakeholder for AI enablement work in Air Cargo
Policy. Primary audience for demo-first capability that fits within TSA's approved stack and
governance.

# Notes
- Interested in practical, deployable tools rather than theoretical discussions.
- Values clear governance and sustainability for any AI capability.
- Likely to be a champion for successful tools that show value quickly.

# Interaction log
- 2026-06-20: Briefed initial concept for Paperclip-style AI tooling supporting Air Cargo Policy.
- 2026-06-10: Detail kickoff conversation; aligned on goals and constraints.
```

# Spec change discipline

To keep tools stable:

- Treat this file (`/network/network.md`) as the locked spec.
- Any changes to:
  - Directory layout
  - Frontmatter schema
  - Required headings
  - Bullet formats
  must be made explicitly here and reflected in tool code.

Tools may assume:

- All contacts live under `/network/contacts/`.
- All contact frontmatter includes the required fields defined above.
- All contact bodies contain the three canonical sections:
  - `# Summary`
  - `# Notes`
  - `# Interaction log`
