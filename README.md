# pokemon-go-iv-calculator
While there exist many IV calculators already, I prefer to keep track of my Pokemon in a spreadsheet environment, because of the **following benefits on top of existing IV calculators**:

* **No need to re-calculate IVs again and again when comparing new catches with older Pokemon** (Alternatively, no need to waste additional time in renaming your Pokemon into a cipher to store IV information)
* **Faster inputs in a spreadsheet environment** versus a web-based form - can batch input Pokemon Names & CPs from the box screen first, followed by HP & Stardust information later for efficiency
* **Analysis-ready data capture:** All inputs and computed IV ranges are stored in a flat file structure (1 Pokemon, 1 row) to enable easy analysis across your current box of Pokemon (e.g. sort by CP, #, easily find the max possible IV across 100 Pidgeys, note which Pokemon are for keeps vs. evolve, keep count of evolves for Lucky Egg, etc.)

**Instructions for Use:**

* Blue shaded cells are inputs
* Appraisal inputs are based on A, B, C, D grading instead of the team-specific phrase outputs
* (Optional) Input exact Pokemon level if you know it (you can infer this by placing your thumb near the dot, and move to other same Stardust Pokemon, which are typically adjacent if sorted by CP, and check whether the dot is in a different location - from this you can infer which of the two integer levels it is, assuming the Pokemon non-Powered)
* (Optional) Pokemon are assumed to be non-Powered by default, to reduce inputs, but there is a 0/1 switch for this
* (Optional) Store remarks on the right-hand side to keep notes about why I am keeping that Pokemon (e.g. bench vs. evolve)

**Plans**
* Gen 4 update
