{smcl}
{* *! version 1.1  Published August 2, 2023}{...}
{p2colset 2 12 14 28}{...}
{right: Version 1.1 }
{p2col:{bf:dtabxl} {hline 2}}Tabulate univariate descriptive statistics in Excel{p_end}
{p2colreset}{...}


{marker syntax}{...}
{title:Syntax}

{p 8 18 2}
{cmdab:dtabxl} {varlist}
{ifin}
{cmd:using} {it:{help filename}}
[{cmd:,}
{it:{help dtabxl##options:options}  {help dtabxl##bioptions:bioptions} {help dtabxl##sigoptions:sigoptions}}]


{marker options}{...}
{synoptset 24 tabbed}{...}
{synopt:{it:options}}Description{p_end}
{synoptline}
{synopt:{cmdab:stats(}{it:{help dtabxl##statname:statname}} [{it:...}]{cmd:)}}specify custom list of statistics to report; default is to report {opt n}, {opt mean}, {opt sd}, {opt p25}, {opt median}, and {opt p75}{p_end}
{synopt:{opt sheetname(text)}}specify custom sheet name; default sheet name is "Descriptives"{p_end}
{synopt:{opt tablename(text)}}specify custom table name; default table name is "Descriptive statistics"{p_end}
{synopt:{opt replace}}required to overwrite an existing sheet in an existing Excel file; not required to add new sheets to an existing Excel file{p_end}
{synopt:{opt roundto(#)}}set number of decimal places to round to; # must be integer between zero and 26; default is # = 2{p_end}
{synopt:{opt nozeros}}set zeros to missing to calculate descriptive statistics{p_end}
{synopt:{opt bifurcate(bivar)}}Bifurcate sample based on {it:bivar}; {it:bivar} must be 1/0 indicator{p_end}
{synopt:{opt extrarows(#)}}insert extra rows between statistics; # must be integer between one and 10{p_end}
{synopt:{opt extracols(#)}}insert extra columns between statistics; # must be integer between one and 10{p_end}
{synoptline}
{p2colreset}{...}
{p 4 6 2}
  If {opt replace} specified, {opt dtabxl} overwrites sheet given by {opt sheetname(text)} (or "Descriptives" if no {opt sheetname(text)} given); it does not alter other sheets in {it:filename}
  {p_end}


{marker bioptions}{...}
{synoptset 24 tabbed}{...}
{synopt:{it:bioptions}}Description{p_end}
{synoptline}
{synopt:{opt switch}}tabulate statistics for {it:bivar} = 0 on left and {it:bivar} = 1 on right; default is {it:bivar} = 1 on left and {it:bivar} = 0 on right{p_end}
{synopt:{opt extrabicols(#)}}insert extra columns between {it:bivar} = 1/0 sides, and between difference column(s) if {opt testmean}/{opt testmedian} specified; # must be integer between one and 10{p_end}
{synopt:{opt testmean}}test means across subsamples using Stata's {help ttest} command; see {browse "www.zach.prof":zach.prof} for additional details{p_end}
{synopt:{opt testmedian}}test medians across subsamples using Stata's {help median} command; see {browse "www.zach.prof":zach.prof} for additional details{p_end}
{synoptline}
{p2colreset}{...}
{p 4 6 2}
  {it:bioptions} only allowed if {opt bifurcate} specified
  {p_end}


{marker sigoptions}{...}
{synoptset 24 tabbed}{...}
{synopt:{it:sigoptions}}Description{p_end}
{synoptline}
{synopt:{opt bold}}use boldfaced text to indicate statistical significance{p_end}
{synopt:{opt italic}}use italicized text to indicate statistical significance{p_end}
{synopt:{opt nostars}}do not use stars to indicate statistical significance{p_end}
{synopt:{opt sleft}}remove difference column(s) and display significance on left{p_end}
{synopt:{opt sright}}remove difference column(s) and display significance on right; can be used with {opt sleft} to display significance on both sides{p_end}
{synopt:{opt sig(#)}}set significance level; # must be between zero and one; p < # receives star, boldface, and/or italic{p_end}
{synopt:{opt 3stars(# # #)}}use three stars to indicate significance; # # # must contain three numbers between zero and one; order does not matter; p < the largest (smallest, other) # receives one (three, two) star(s){p_end}
{synoptline}
{p2colreset}{...}
{p 4 6 2}
  {it:sigoptions} only allowed if {opt testmean} or {opt testmedian} specified; Fisher's exact p-values used for medians
  {p_end}


{marker statname}{...}
{synoptset 12 tabbed}{...}
{synopt:{it:statname}}Definition{p_end}
{synoptline}
{synopt:{opt mean}} mean{p_end}
{synopt:{opt n}} number of nonmissing observations{p_end}
{synopt:{opt sum}} sum{p_end}
{synopt:{opt max}} maximum{p_end}
{synopt:{opt min}} minimum{p_end}
{synopt:{opt range}} range = {opt max} - {opt min}{p_end}
{synopt:{opt sd}} standard deviation{p_end}
{synopt:{opt var}} variance{p_end}
{synopt:{opt cv}} coefficient of variation ({cmd:sd/mean}){p_end}
{synopt:{opt semean}} standard error of mean ({cmd:sd/sqrt(n)}){p_end}
{synopt:{opt skew}} skewness{p_end}
{synopt:{opt kurt}} kurtosis{p_end}
{synopt:{opt p1}} 1st percentile{p_end}
{synopt:{opt p5}} 5th percentile{p_end}
{synopt:{opt p10}} 10th percentile{p_end}
{synopt:{opt p25}} 25th percentile{p_end}
{synopt:{opt p50}} 50th percentile{p_end}
{synopt:{opt p75}} 75th percentile{p_end}
{synopt:{opt p90}} 90th percentile{p_end}
{synopt:{opt p95}} 95th percentile{p_end}
{synopt:{opt p99}} 99th percentile{p_end}
{synopt:{opt iqr}} interquartile range = {opt p75} - {opt p25}{p_end}
{synopt:{opt median}} median (same as {opt p50}){p_end}
{synoptline}
{p2colreset}{...}


{marker description}{...}
{title:Description}

{pstd}{opt dtabxl} provides a variety of options for tabulating univariate descriptive statistics in Excel. These options are designed to:{p_end}

{p 8 12}(1) streamline the process of communicating results to coauthors, and{p_end}
{p 8 12}(2) minimize time spent populating descriptive statistics tables in Word.{p_end}


{marker examples}{...}
{title:Examples}

{pstd}
Examples provided at {browse "www.zach.prof":zach.prof}
{p_end}


{marker contact}{...}
{title:Author}

{pstd}
Zachary King{break}
Email: {browse "mailto:zacharyjking90@gmail.com":zacharyjking90@gmail.com}{break}
Website: {browse "www.zach.prof":zach.prof}{break}
SSRN: {browse "https://papers.ssrn.com/sol3/cf_dev/AbsByAuth.cfm?per_id=2623799":https://papers.ssrn.com}
{p_end}


{title:Acknowledgements}

{pstd}
I thank the following individuals for helpful feedback and suggestions on {opt dtabxl}, this help file, and the associated documentation on {browse "www.zach.prof":zach.prof}:{p_end}

{pstd}
Jesse Chan{break}
Rachel Flam{break}
Ben Osswald
{p_end}