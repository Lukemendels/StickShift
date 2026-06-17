# Build Portfolio — Usage Loop

Four steps. Do them in order. Stop after step 4.

---

## Step 1 — Populate `builds/`

Brain-dump every project to DHSChat. Ask it to return one schema-conformant
`.md` file per project. Copy each file into `builds/` with a short slug as the
filename (e.g. `briefing-builder.md`).

One file = one build. That's the whole data model.

## Step 2 — Run the Index Generator

Open `OKFIndexGenerator.bas` in Excel's VBA editor (Alt+F11 → Import File).

Before first run:
1. Set `BUNDLE_ROOT` in the constants block to your actual OneDrive path
   (e.g. `C:\Users\YourName\OneDrive\build-portfolio\`).
2. Verify Tools → References includes **Microsoft ActiveX Data Objects 2.x**.
3. Test against a 3-file subfolder before running on the full bundle.
4. Run `GenerateOKFIndexes`. Confirm `index.md` appears at the bundle root
   and reads as a status board with `working` items at the top, sorted oldest
   `last_touched` first.

## Step 3 — Paste the index into the Portfolio Strategist

Open the Portfolio Strategist DHSChat Assistant (bookmark the URL after first
setup). Paste the contents of `index.md` as your first message. Then ask:
"What should I work on?"

The Assistant holds the reasoning frame. The index is the data. Fresh paste
each session = always current, no retrieval needed.

## Step 4 — Act

Do the one thing the Strategist recommends. When it's done, update the build
file (`status`, `last_touched`, `# Next action`), re-run the generator, and
repeat from Step 3.

---

That's the shipped tool. Do not add to it until you've outgrown it.
