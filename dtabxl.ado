*! Title:       dtabxl.ado   
*! Version:     1.1 published August 2, 2023
*! Author:      Zachary King 
*! Email:       zacharyjking90@gmail.com
*! Description: Tabulate univariate descriptive statistics in Excel

program def dtabxl

	* Ensure Stata runs dtabxl using version 17 syntax
	
	version 17
	
	* Define syntax

	syntax varlist(numeric) [if] [in] using/ [, stats(namelist) ///
	sheetname(string) tablename(string)                        ///
	extrarows(numlist integer max=1 >0 <11)                   ///
	extracols(numlist integer max=1 >0 <11)                     ///
	roundto(numlist integer max=1 >=0 <=26)                    ///
	bifurcate(varlist numeric min=1 max=1)                    ///
	extrabicols(numlist integer max=1 >0 <11)                   ///
	sig(numlist max=1 >0 <1)                                   ///
	3stars(numlist sort min=3 max=3 >0 <1)                    ///
	TESTMEAN TESTMEDIAN NOSTARS BOLD ITALIC NOZEROS             ///
	SWITCH REPLACE SLEFT SRIGHT]
	
	* Set default statistics if not specified
	
	if "`stats'" == "" local stats n mean sd p25 median p75
	
	* Validate statistics
	
	tempname invalid_stat
	local `invalid_stat' = 0
	
	foreach stat in `stats' {
		
		capture valid_stat , `stat' // program defined below
		
		if _rc di in err `"unknown statistic: `stat'"'
		if _rc local `invalid_stat' = 1
		
	}
	
	if ``invalid_stat'' exit 198
	
	* Set default sheet name if not specified
	
	if "`sheetname'" == "" local sheetname = "Descriptives"
	
	* Validate length of sheet name not too long
	
	if length("`sheetname'") >= 32 {
		di as error "sheet name too long; must be less than 32 characters"
		exit 198
	}
	
	* Set default table name if not specified
	
	if "`tablename'" == "" local tablename = "Descriptive statistics"
	
	* If replace option not specified, verify current sheet does not exist
	
	if "`replace'" == "" {
		
		preserve
		cap: import excel "`using'", sheet("`sheetname'") clear
		
		if _rc {
			restore
		}
		
		else {
			
			restore
			
			di as error "worksheet {bf:`sheetname'} already exists, specify {bf:replace} option to overwrite it"
			exit 601
			
		}
		
	}
	
	* Set default rounding if not specified
	
	if "`roundto'" == "" local roundto = 2
	
	* Validate bifurcate variable is binary 1/0 variable if specified
	
	if "`bifurcate'" != "" {
		
		qui: levelsof `bifurcate'
		
		if `r(r)' != 2 {
			di as error "{bf:bifurcate} option can only be used with 1/0 dummy variable"
			exit 198
		}
		
		foreach level in `r(levels)' {
			
			if `level' != 1 & `level' != 0 {
				di as error "{bf:bifurcate} option can only be used with 1/0 dummy variable"
				exit 198
			}
			
		}
		
	}
	
	* Set default extra rows if not specified
	
	if "`extrarows'" == "" local extrarows = 0
	
	* Set default extra columns if not specified
	
	if "`extracols'" == "" local extracols = 0
	
	* Only allow bioptions if bifurcate option is specified
	
	tempname bioption_check
	
	local `bioption_check' = 0
	
	if "`bifurcate'" == "" {
		
		if "`switch'" != "" {
			di as error "{bf:switch} option only allowed with {bf:bifurcate} option"
			local `bioption_check' = ``bioption_check'' + 1
		}
		
		if "`extrabicols'" != "" {
			di as error "{bf:extrabicols} option only allowed with {bf:bifurcate} option"
			local `bioption_check' = ``bioption_check'' + 1
		}
		
		if "`testmean'" != "" {
			di as error "{bf:testmean} option only allowed with {bf:bifurcate} option"
			local `bioption_check' = ``bioption_check'' + 1
		}
		
		if "`testmedian'" != "" {
			di as error "{bf:testmedian} option only allowed with {bf:bifurcate} option"
			local `bioption_check' = ``bioption_check'' + 1
		}
		
		if ``bioption_check'' > 0 {
			exit 198
		}
		
	}
	
	* Set default extra bifurcate columns if not specified
	
	if "`extrabicols'" == "" local extrabicols = 0
	
	* If testmean specified, verify mean in stats
	
	if "`testmean'" != "" {
		
		tempname mean_check
		local `mean_check' = 0
		
		foreach stat in `stats' {
			if "`stat'" == "mean" local `mean_check' = 1
		}
		
		if ``mean_check'' == 0 {
			di as error "{bf:testmean} option not allowed if {bf:stats} option used to exclude {bf:mean} statistic from table"
			exit 198
		}
		
	}
	
	* If testmedian specified, verify median or p50 in stats
	
	if "`testmedian'" != "" {
		
		tempname median_check
		local `median_check' = 0
		
		foreach stat in `stats' {
			if "`stat'" == "median" local `median_check' = 1
			if "`stat'" == "p50" local `median_check' = 1
		}
		
		if ``median_check'' == 0 {
			di as error "{bf:testmedian} option not allowed if {bf:stats} option used to exclude {bf:median/p50} statistic from table"
			exit 198
		}
		
	}
	
	* Only allow sigoptions if testmean or testmedian option specified
	
	tempname sigoption_check
	
	local `sigoption_check' = 0
	
	if "`testmean'" == "" & "`testmedian'" == "" {
		
		if "`bold'" != "" {
			di as error "{bf:bold} option only allowed if {bf:testmean} and/or {bf:testmedian} option specified"
			local `sigoption_check' = ``sigoption_check'' + 1
		}
		
		if "`italic'" != "" {
			di as error "{bf:italic} option only allowed if {bf:testmean} and/or {bf:testmedian} option specified"
			local `sigoption_check' = ``sigoption_check'' + 1
		}
		
		if "`nostars'" != "" {
			di as error "{bf:nostars} option only allowed if {bf:testmean} and/or {bf:testmedian} option specified"
			local `sigoption_check' = ``sigoption_check'' + 1
		}
		
		if "`sleft'" != "" {
			di as error "{bf:sleft} option only allowed if {bf:testmean} and/or {bf:testmedian} option specified"
			local `sigoption_check' = ``sigoption_check'' + 1
		}
		
		if "`sright'" != "" {
			di as error "{bf:sright} option only allowed if {bf:testmean} and/or {bf:testmedian} option specified"
			local `sigoption_check' = ``sigoption_check'' + 1
		}
		
		if "`sig'" != "" {
			di as error "{bf:sig} option only allowed if {bf:testmean} and/or {bf:testmedian} option specified"
			local `sigoption_check' = ``sigoption_check'' + 1
		}
		
		if "`3stars'" != "" {
			di as error "{bf:3stars} option only allowed if {bf:testmean} and/or {bf:testmedian} option specified"
			local `sigoption_check' = ``sigoption_check'' + 1
		}
		
		if ``sigoption_check'' > 0 {
			exit 198
		}
		
	}
	
	* Ensure sig and 3stars options not used in combination
	
	if "`sig'" != "" & "`3stars'" != "" {
		di as error "only one of {bf:sig} and {bf:3stars} options is allowed"
		exit 198
	}
	
	* Ensure bold, italic, and nostars options not used with 3stars option
	
	if "`3stars'" != "" & ("`bold'" != "" | "`italic'" != "" | "`nostars'" != "") {
		di as error "{bf:bold}, {bf:italic}, and {bf:nostars} options not allowed with {bf:3stars} option"
		exit 198
	}
	
	* Set default significance level if not specified
	
	if "`sig'" == "" local sig = 0.05

	* Save significance levels if 3stars specified
	
	if "`3stars'" != "" {
		
		tempname snum s1 s2 s3
		local `snum' = 3
		
		foreach level of numlist `3stars' {
			local `s``snum''' = `level'
			local `snum' = ``snum'' - 1
		}
		
	}
	
	* Set whether significance is indicated on left and/or right
	
	tempname sig_on_left sig_on_right
	
	if "`sleft'" != "" local `sig_on_left' = 1
	else local `sig_on_left' = 0
	
	if "`sright'" != "" local `sig_on_right' = 1
	else local `sig_on_right' = 0
	
	* Set zeros to missing if nozeros is specified, and calculate and save descriptive statistics
	
	if "`bifurcate'" != "" {
		
		preserve
		if "`switch'" != "" qui: keep if `bifurcate' == 0
		else qui: keep if `bifurcate' == 1
		
		if "`nozeros'" != "" {
			foreach v of varlist `varlist' {
				qui: replace `v' = . if `v' == 0
			}
		}
		
		qui: tabstat `varlist' `if' `in', save s(me n su ma mi r sd v cv sem sk k p1 p5 p10 p90 p95 p99 iqr q)
		matrix define stats = r(StatTotal)
		
		restore
		
		preserve
		if "`switch'" != "" qui: keep if `bifurcate' == 1
		else qui: keep if `bifurcate' == 0
		
		if "`nozeros'" != "" {
			foreach v of varlist `varlist' {
				qui: replace `v' = . if `v' == 0
			}
		}
		
		qui: tabstat `varlist' `if' `in', save s(me n su ma mi r sd v cv sem sk k p1 p5 p10 p90 p95 p99 iqr q)
		matrix define stats = (stats,r(StatTotal))
		
		restore
		
		preserve
	
		if "`nozeros'" != "" {
			foreach v of varlist `varlist' {
				qui: replace `v' = . if `v' == 0
			}
		}
		
	}
	
	else {
		
		preserve
	
		if "`nozeros'" != "" {
			foreach v of varlist `varlist' {
				qui: replace `v' = . if `v' == 0
			}
		}
		
		qui: tabstat `varlist' `if' `in', save s(me n su ma mi r sd v cv sem sk k p1 p5 p10 p90 p95 p99 iqr q)
		matrix define stats = r(StatTotal)
		
	}
	
	* Index rows of each statistic in the "stats" matrix
	
	index_matrix_rows // program defined below
	
	* Open Excel file
	
	qui: putexcel set "`using'", open modify sh("`sheetname'", replace)
	
	* Write table name to cell A1
	
	qui: putexcel A1 = "`tablename'"
	
	* Tokenize A, B, C, ... , AA, AB, AC, ... , ZZ to loop over Excel columns
	
	tempname cell_letters
	
	forvalues i = 0/26 {
		if `i' == 0 {
			forvalues j = 1/26 {
				local `cell_letters' = "``cell_letters'' " + char(`j' + 64)
			}
		}
		else {
			forvalues j = 1/26 {
				local `cell_letters' = "``cell_letters'' " + char(`i' + 64) + char(`j' + 64)
			}
		}
	}
	
	tokenize "``cell_letters'' `varlist'"
	
	* Write variable names to Excel
	
	tempname r
	if "`bifurcate'" != "" local `r' = 4
	else local `r' = 3

	foreach v of varlist `varlist' {
		qui: putexcel A``r'' = "`v'"
		local `r' = ``r'' + 1 + `extrarows'
	}
	
	* Write correlation table note to Excel
	
	tempname signote
	
	if "`testmean'" == "" & "`testmedian'" == "" {
		local `signote' = ""
	}
	
	else if "`nostars'" != "" {
		if "`bold'" == "" & "`italic'" == "" local `signote' = ""
		else if "`bold'" != "" & "`italic'" != "" local `signote' = "Bold italics indicate significant at p-value < 0`sig' level (two-tailed)"
		else if "`bold'" != "" local `signote' = "Bold indicates significant at p-value < 0`sig' level (two-tailed)"
		else local `signote' = "Italics indicate significant at p-value < 0`sig' level (two-tailed)"
	}
	
	else {
		if "`bold'" == "" & "`italic'" == "" local `signote' = "* indicates significant at p-value < 0`sig' level (two-tailed)"
		else if "`bold'" != "" & "`italic'" != "" local `signote' = "Bold italics with * indicates significant at p-value < 0`sig' level (two-tailed)"
		else if "`bold'" != "" local `signote' = "Bold with * indicates significant at p-value < 0`sig' level (two-tailed)"
		else local `signote' = "Italics with * indicates significant at p-value < 0`sig' level (two-tailed)"
	}
	
	if "`3stars'" != "" local `signote' = "*** (**, *) indicates significant at p-value < 0``s3'' (0``s2'', 0``s1'') level (two-tailed)"   
	
	qui: putexcel A``r'' = "``signote''"
	
	* Write statistic names to Excel
	
	tempname c
	local `c' = 2
	
	if "`bifurcate'" != "" local `r' = 3
	else local `r' = 2

	foreach stat in `stats' {
		qui: putexcel ```c'''``r'' = "`stat'"
		local `c' = ``c'' + 1 + `extracols'
	}
	
	if "`bifurcate'" != "" {
		
		local `c' = ``c'' + `extrabicols'
		
		foreach stat in `stats' {
			qui: putexcel ```c'''``r'' = "`stat'"
			local `c' = ``c'' + 1 + `extracols'
		}
		
	}
	
	if "`testmean'" != "" & ``sig_on_left'' == 0 & ``sig_on_right'' == 0 {
		local `c' = ``c'' + `extrabicols'
		qui: putexcel ```c'''``r'' = "mean diff"
		local `c' = ``c'' + 1 + `extracols'
	}
	
	if "`testmedian'" != "" & ``sig_on_left'' == 0 & ``sig_on_right'' == 0 {
		local `c' = ``c'' + `extrabicols'
		qui: putexcel ```c'''``r'' = "median diff"
	}
	
	* Write subsample labels to Excel if bifurcate
	
	tempname nvars c2 nstats
	local `nstats': list sizeof local(stats)
	
	if "`bifurcate'" != "" {
		
		local `c' = 2
		
		if "`switch'" != "" qui: putexcel ```c'''2 = "`bifurcate' = 0"
		else qui: putexcel ```c'''2 = "`bifurcate' = 1"
		
		local `c2' = ``nstats'' + ``c'' - 1 + `extracols'*(``nstats''-1)
		
		qui: putexcel ```c'''2:```c2'''2, overwritefmt merge hcenter
		
		local `c' = ``c2'' + 1 + `extrabicols' + `extracols'
		
		if "`switch'" != "" qui: putexcel ```c'''2 = "`bifurcate' = 1"
		else qui: putexcel ```c'''2 = "`bifurcate' = 0"
		
		local `c2' = ``nstats'' + ``c'' - 1 + `extracols'*(``nstats''-1)
		
		qui: putexcel ```c'''2:```c2'''2, overwritefmt merge hcenter
		
	}
	
	* Write descriptive statistics to Excel
	
	tempname dval digits int_length colsofstats nvars varindex skipsig
	
	if "`bifurcate'" != "" local `r' = 4
	else local `r' = 3
	
	local `colsofstats' = colsof(stats)
	
	local `nvars': list sizeof local(varlist)
	
	forvalues i = 1/``colsofstats'' {
		
		if `i' == ``nvars'' + 1 local `r' = 4
		
		if `i' > ``nvars'' local `c' = 2 + ``nstats'' + `extracols'*``nstats'' + `extrabicols'
		else local `c' = 2
		
		if `i' > ``nvars'' local `varindex' = `i' - ``nvars'' + 702
		else local `varindex' = `i' + 702
		
		if `i' <= ``nvars'' {
			if ``sig_on_left'' == 1 local `skipsig' = 0
			else local `skipsig' = 1
		}
		
		else {
			if ``sig_on_right'' == 1 local `skipsig' = 0
			else local `skipsig' = 1
		}
		
		foreach stat in `stats' {
			
			local `int_length' = length(strofreal(int(stats[e(`stat'),`i'])))
			
			if "`stat'" == "n" {
				local `digits' = ``int_length'' + ceil(``int_length''/3) - 1
				local `dval' : di %-``digits''.0fc stats[e(`stat'),`i']
			}
			
			else {
				local `digits' = `roundto' + ``int_length'' + ceil(``int_length''/3)
				local `dval' : di %-``digits''.`roundto'fc stats[e(`stat'),`i']
			}
			
			if "`stat'" == "mean" & "`testmean'" != "" & ``skipsig'' == 0 {
				
				qui: ttest ```varindex''' `if' `in' , by(`bifurcate')
				
				if "`3stars'" == "" {
					
					if "`nostars'" != "" qui: putexcel ```c'''``r'' = "``dval''"
					
					else {
						add_stars, value("``dval''") p(`r(p)') levelone(`sig')
						qui: putexcel ```c'''``r'' = "`s(dval)'"
					} 
					
					if r(p) < `sig' & ("`bold'" != "" | "`italic'" != "") qui: putexcel ```c'''``r'', overwritefmt `bold' `italic'
					
				}
				
				else if "`3stars'" != "" {
					
					add_stars, value("``dval''") p(`r(p)') levelone(``s1'') leveltwo(``s2'') levelthree(``s3'')
					qui: putexcel ```c'''``r'' = "`s(dval)'"
					
				}
				
			}
			
			else if ("`stat'" == "median" | "`stat'" == "p50") & "`testmedian'" != "" & ``skipsig'' == 0 {
				
				qui: median ```varindex''' `if' `in' , by(`bifurcate') exact
				
				if "`3stars'" == "" {
					
					if "`nostars'" != "" qui: putexcel ```c'''``r'' = "``dval''"
					
					else {
						add_stars, value("``dval''") p(`r(p_exact)') levelone(`sig')
						qui: putexcel ```c'''``r'' = "`s(dval)'"
					} 
					
					if r(p_exact) < `sig' & ("`bold'" != "" | "`italic'" != "") qui: putexcel ```c'''``r'', overwritefmt `bold' `italic'
					
				}
				
				else if "`3stars'" != "" {
					
					add_stars, value("``dval''") p(`r(p_exact)') levelone(``s1'') leveltwo(``s2'') levelthree(``s3'')
					qui: putexcel ```c'''``r'' = "`s(dval)'"
					
				}
				
			}
			
			else {
				
				qui: putexcel ```c'''``r'' = "``dval''"
				
			}
			
			local `c' = ``c'' + 1 + `extracols'
			
		}
		
		local `r' = ``r'' + 1 + `extrarows'
	}
	
	* Write mean differences to Excel
	
	if "`testmean'" != "" & ``sig_on_left'' == 0 & ``sig_on_right'' == 0 {
		
		local `c' = ``c'' + `extrabicols'
		local `r' = 4
		
		foreach v of varlist `varlist' {
			
			qui: ttest `v' `if' `in' , by(`bifurcate')
			
			tempname diff
			
			if "`switch'" == "" local `diff' = r(mu_1) - r(mu_2)
			else local `diff' = r(mu_2) - r(mu_1)
			
			local `int_length' = length(strofreal(int(``diff'')))
			
			local `digits' = `roundto' + ``int_length'' + ceil(``int_length''/3)
			local `dval' : di %-``digits''.`roundto'fc ``diff''
			
			if "`3stars'" == "" {
					
				if "`nostars'" != "" qui: putexcel ```c'''``r'' = "``dval''"
					
				else {
					add_stars, value("``dval''") p(`r(p)') levelone(`sig')
					qui: putexcel ```c'''``r'' = "`s(dval)'"
				} 
					
				if r(p) < `sig' & ("`bold'" != "" | "`italic'" != "") qui: putexcel ```c'''``r'', overwritefmt `bold' `italic'
					
			}
				
			else if "`3stars'" != "" {
					
				add_stars, value("``dval''") p(`r(p)') levelone(``s1'') leveltwo(``s2'') levelthree(``s3'')
				qui: putexcel ```c'''``r'' = "`s(dval)'"
					
			}
			
			local `r' = ``r'' + 1 + `extrarows'
			
		}
		
		local `c' = ``c'' + 1 + `extracols'
		
	}
	
	* Write median differences to Excel
	
	if "`testmedian'" != "" & ``sig_on_left'' == 0 & ``sig_on_right'' == 0 {
		
		local `c' = ``c'' + `extrabicols'
		local `r' = 4
		
		tempname vind
		local `vind' = 1
		
		foreach v of varlist `varlist' {
			
			qui: median `v' `if' `in' , by(`bifurcate') exact
			
			tempname diff
			local `diff' = stats[e(median),``vind''+``nvars''] - stats[e(median),``vind'']
			
			local `int_length' = length(strofreal(int(``diff'')))
			
			local `digits' = `roundto' + ``int_length'' + ceil(``int_length''/3)
			local `dval' : di %-``digits''.`roundto'fc ``diff''
			
			if "`3stars'" == "" {
					
				if "`nostars'" != "" qui: putexcel ```c'''``r'' = "``dval''"
					
				else {
					add_stars, value("``dval''") p(`r(p_exact)') levelone(`sig')
					qui: putexcel ```c'''``r'' = "`s(dval)'"
				} 
					
				if r(p_exact) < `sig' & ("`bold'" != "" | "`italic'" != "") qui: putexcel ```c'''``r'', overwritefmt `bold' `italic'
					
			}
				
			else if "`3stars'" != "" {
					
				add_stars, value("``dval''") p(`r(p_exact)') levelone(``s1'') leveltwo(``s2'') levelthree(``s3'')
				qui: putexcel ```c'''``r'' = "`s(dval)'"
					
			}
			
			local `r' = ``r'' + 1 + `extrarows'
			local `vind' = ``vind'' + 1
			
		}
		
		local `c' = ``c'' + 1 + `extracols'
		
	}
	
	* Close Excel
	
	qui: putexcel close
	
	* Display link to open Excel file
	
	display_clickable using "`using'"
	
end

program def valid_stat
	
	syntax [, MEAN N SUM MAX MIN RANGE SD VAR CV SEMEAN SKEW KURT P1 P5 P10 P25 P50 P75 P90 P95 P99 IQR MEDIAN]
	
end

program def index_matrix_rows, eclass
	
	ereturn clear
	ereturn scalar mean   =  1
	ereturn scalar n      =  2
	ereturn scalar sum    =  3
	ereturn scalar max    =  4
	ereturn scalar min    =  5
	ereturn scalar range  =  6
	ereturn scalar sd     =  7
	ereturn scalar var    =  8
	ereturn scalar cv     =  9
	ereturn scalar semean = 10
	ereturn scalar skew   = 11
	ereturn scalar kurt   = 12
	ereturn scalar p1     = 13
	ereturn scalar p5     = 14
	ereturn scalar p10    = 15
	ereturn scalar p25    = 20
	ereturn scalar p50    = 21
	ereturn scalar p75    = 22
	ereturn scalar p90    = 16
	ereturn scalar p95    = 17
	ereturn scalar p99    = 18
	ereturn scalar iqr    = 19
	ereturn scalar median = 21

end

program def add_stars, sclass 
	
	syntax , value(string) p(numlist max=1) levelone(numlist max=1) [leveltwo(numlist max=1) levelthree(numlist max=1)]
	
	if "`leveltwo'" == "" & "`levelthree'" == "" {
		
		if `p' < `levelone' sreturn local dval "`value'*"
		else sreturn local dval "`value'"
		
	}
	
	else {
		
		if `p' < `levelthree' sreturn local dval "`value'***"
		else if `p' < `leveltwo' sreturn local dval "`value'**"
		else if `p' < `levelone' sreturn local dval "`value'*"
		else sreturn local dval "`value'"
		
	}
	
end

program def display_clickable

	syntax using/
	
	mata {
		path_and_file = st_local("using")
		path = ""
		file = ""
		pathsplit(path_and_file, path, file)
		st_local("path", path) 
		st_local("file", file) 
	}
	
	if "`path'" == "" local path = "`c(pwd)'"
	if strrpos("`file'", ".xlsx") == 0 & strrpos("`file'", ".xls") == 0 local file = "`file'.xlsx"
	
	mata {
		file_w_ext_and_path = pathjoin(st_local("path"),st_local("file"))
		st_local("file_w_ext_and_path", file_w_ext_and_path) 
	}
	
	di as text `"{browse "`file_w_ext_and_path'":click here to open Excel output}"'
	
end