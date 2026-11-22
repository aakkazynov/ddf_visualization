# Single prompt: DDF Funders Network (with international filter)

Copy everything between the lines into Claude.

---

You are helping me build a data visualization project for a Data Visualization class. Do the work in three phases and report back briefly after each phase before moving on.

## Phase 1: Research and dataset

1. Search the Doris Duke Foundation (DDF) website and the wider web for every grant program where DDF pooled or co-funded money with other funders (example: Humanity AI). Include both formal pooled funds and looser co-funding partnerships, and label each as "pooled" or "co-funded".
2. Group each program under DDF's **official** program-area names as they appear on dorisduke.org (e.g., Arts, Child Well-being, Environment, Medical Research, Building Bridges, plus "Cross-program" if needed). Use each program's official full name, and each funder's official name (check each one).
3. For each program, record: official name, DDF program area, year announced or first reported, total amount and what that amount means (pledge, goal, first round, or cumulative total), list of funders, manager or intermediary, source URLs, and a confidence level (high/medium/low).
4. **Geography (new):** for each program, record where the money goes:
   - `scope`: "US", "International", or "Both"
   - `countries`: the list of recipient countries (ISO country names), and the region for each (Africa, Asia, Europe, Latin America & Caribbean, Middle East, North America, Oceania). Include the US when US grantees receive funds.
   - Include DDF's international and global programs (for example global health or research work abroad), not only US ones. If a program funds work abroad but the countries are not public, set `countries` to "Not disclosed" and say so. Do not guess countries.
5. For each funder, record: official name, type (private foundation, corporate, public charity, government, etc.), headquarters country, endowment or total assets (FY23 AUM or the latest tax filing from 2023–2024, with the source), website URL, and the programs they joined.
6. Save everything in an Excel workbook `DDF_Pooled_Fund_Programs.xlsx` with tabs: Programs, Funders, Program–Funder links, Countries (one row per program–country pair), Excluded items, and Leads to follow up. Also export the same data as `data.json` for the visualization.
7. Tell me which rows are medium/low confidence and which facts you could not verify.

## Phase 2: The visualization (single self-contained `index.html`, D3.js v7)

**Layout and encoding**
- Network graph. Each dot is a funder. DDF is the central node.
- Force-directed (spring) layout, with funders **clustered by program** so that co-funders of the same program sit together. Dots can be dragged. Scroll to zoom, plus +/− zoom buttons.
- Edges: DDF→funder links have width set by the program's funding amount. Funders that co-fund the same program are also linked to each other; these non-DDF links are much fainter and thinner, in gray. Line width = combined size of the programs the two funders share. Use dotted lines when no amount was found. In the "About this data" text, explain that the lines show shared membership, not a direct relationship.
- Dot size = funder endowment (square-root scale; largest ≈ 24px radius, smallest ≈ 4px). Funders with no endowment figure get a fixed size of about 6px, drawn pale with a thin outline. The legend notes "pale = endowment not found".
- Dot color = DDF program area. A funder in more than one program area is drawn as a pie split across those colors.
- **Tap targets:** add invisible hit areas of at least 16px radius on phones and 10px on desktop.

**Color scheme (replace the original palette)**
Dark background (#14161a), white-ish text (90% white). Program-area colors from a saturated red/blue/green/orange family, each checked for contrast on the dark background:
- Arts: red #E53935
- Medical Research: blue #1E88E5
- Environment: green #43A047
- Child Well-being: orange #FB8C00
- Building Bridges: purple #8E24AA
- Cross-program: teal #00ACC1
- Extra areas if needed: crimson #C62828, sky blue #64B5F6, lime #9CCC65, amber #FFB300
- DDF node: white with a thin gold ring (#FFD54F).
- Highlight ring on selected funders: bright yellow #FFEB3B (not teal, so it stays distinct from the program colors).
- Non-DDF links: gray (#6b6f76) at low opacity; when highlighted, a slightly lighter gray, never white.
- International-only indicators (see below) use the same palette; don't introduce new hues.

**Filters and controls (top bar)**
- **Program filter:** a dropdown of the 13 (or however many) programs, using official names. Choosing one highlights its funders (yellow ring, pale dots fill in more strongly) and opens that program's card on the right. "All" clears the highlight and closes the card.
- **Program-area filter:** using the official DDF area names.
- **Year filter:** a range slider with two handles (first and last year shown) and a ▶ Play button that moves the end handle forward year by year. During playback, show the current year in the top-left corner of the graph: 50px, light gray, Helvetica.
- **International filter (new):**
  - A "Scope" toggle: All / US only / International / Both.
  - A searchable multi-select "Countries" list, grouped by region, built only from countries in the data, with a count of programs next to each.
  - Filtering hides funders and links that don't belong to a matching program. All filters combine (AND). If nothing matches, show "No programs match these filters."
  - Funders tied to at least one international program get a small globe marker (or a dashed outer ring) so international reach is visible without filtering.
  - Optional: a small world map (D3 geoNaturalEarth1 + world-atlas TopoJSON) under the graph, shading recipient countries by number of programs, using a light-to-dark blue scale. Clicking a country applies the country filter, and selecting a program outlines its countries.
- **Reset** button: clears all filters, highlights, and open cards.
- In the graph's top-right corner, add two lightweight text buttons: "N funders" and "N programs". Each opens a list panel on the right. Sort funders by number of programs, then A–Z. Clicking an item opens its detail card, with a "Back to list" button.

**Detail cards (right-side panel)**
- **Funder card:** name; endowment and source year, then "· Website ↗" (opens in a new tab); on the next line "**X** programs with DDF"; HQ country; a list of program names, each clickable to open the program view (with "Back to [funder]").
- **Program card:** official name, program area (color swatch), year, amount and what it means, scope, **recipient countries** (as chips; clicking a chip filters by that country), manager, list of funders (clickable, with "Back to program"), and source links.
- **Edge tooltip:** on hover over DDF→funder links, show the program(s) and amount. Turn off hover on non-DDF links; clicking them still opens a small card.

**Typography and style**
- Minimalist. Title "DDF Funders Network", centered, about 2.2× the body size. Subtitle: "Each dot represents a funder. Click a funder or a line for details. Scroll to zoom."
- Program legend on its own line under the controls, with the scope/globe legend next to it.
- Funder labels: 90% white. Funders with 2+ programs in regular weight; funders with 1 program in a lighter, thinner weight. On hover, make the label only slightly bolder (no fuzzy glow). On hover/click, highlight the node and its neighbors, and fade everything else.
- Add a hint that dots are clickable: "click for more info ›" as a small underlined link to the right of the funder name in the tooltip.
- A collapsible "About this data" section under the legend, closed on load, explaining the sources, the research date, the encodings (size, color, lines, dotted lines, pale dots, globe marker), how countries were found, and the caveats.
- Responsive: works on a laptop and a phone.

## Phase 3: Check and deliver
1. Test that every program in the dropdown highlights its funders. Test every Scope option and several countries. Test Reset. Check that no program is lost when combining filters.
2. List every funder website link you did not actually open, and every country you could not confirm.
3. Give me the files (`index.html`, `data.json`, the workbook) and tell me how to deploy them to Vercel or GitHub Pages.
4. Add a page `how-claude-built-this.html`, linked under "About this data", that shows this prompt and a short summary of what you did in each phase.

Don't invent data. Wherever a figure is missing, show it as missing in the data and the visualization, and tell me.

---
