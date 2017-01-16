# IV Calculator Algorithm
This is the documentation of the algorithm used to determine IV ranges executed in the MS Excel spreadsheet

User Inputs:
*	CP
*	HP
*	Stardust
*	Wild (Optional – Reduces level possibilities from 4 to 2)
*	Known Level (Optional – Narrows to 1 level possibility, known by hatch or arc angle)
*	Appraisal (Optional – Reduces possibilities from 4,096 to ≤360)
  *	Range of Sum of IVs
    *	A (e.g. “Simply amazes”): 37-45
    *	B (e.g. “Strong”): 30-36
    *	C (e.g. “Decent”): 23-29
    *	D (e.g. “May not be great in battle”): 0-22
  *	Attributes (HP, Atk, Def) with highest IV
  *	Range of Best IV
    *	A (e.g. “WOW!”): 15
    *	B (e.g. “Excellent”): 13-14
    *	C (e.g. “Get the job done”): 8-12
	  * D (e.g. “Don’t point to greatness”): 0-7

Definitions:
* A, D, S refer to Base Atk, Def, and Stamina (aka HP)
* x, y, z refer to the IV value of Atk, Def, and Stamina, bounded from 0 to 15
* C refers to the level multiplier
* L refers to the Pokemon level

Known Formulas and Facts:
*	CP = max(10, floor( 1/10 * C^2 * ( (A+x)^2 * (D+y) * (S+z) )^0.5 ) )
*	HP = max(10, floor( C * (S+z) ) )
*	C is a function of L
* Stardust implies one of 4 possible values of L, and can be further reduced if any of the following are known:
  *	Wild: 2 possible values (L1 and L3 where L1, L2, L3, L4 are in ascending order); also L≤30
  *	Hatched: L = min(20, current trainer level)
  *	Arc Angle (based on screenshot or relative comparison)

Process:
*	Determine possible L values (up to 4) based on Stardust value, Wild, Known fields
*	For each L (up to 4 possible values based on Stardust)
  *	Determine possible values of z by reversing the HP formula:
	  * z ≥ if(HP > 10, max(0, ceiling( C^(-1)*HP - S) ), 0) [Infeasible if >15]
	  * z ≤ min(15, floor( C^(-1)*(HP+1) - S ) ) [Infeasible if <0]
	* Determine possible values of (A+x)^2 * (D+y) * (S+z) by reversing the CP formula:
	  * (A+x)^2 * (D+y) * (S+z) ≥ if(CP > 10, 100 * C^(-4) * CP^2, A^2 * D * S)	[Infeasible if > (A+15)^2 * (D+15) * (S+15)]
	  * (A+x)^2 * (D+y) * (S+z) < 100 * C^(-4) * (CP+1)^2 [Infeasible if < A^2 * D * S]
	* Identify possible IV scenarios (out of 4,096) based on:
	  * HP-inferred z values
	  * CP-inferred (A+x)^2 * (D+y) * (S+z) values
	  * Appraisal results
	* If feasible, report minimum and maximum Sum of IVs and individual x, y & z ranges
