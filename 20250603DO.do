// current workingdirectory
cd C:\Users\dionw\Documents\stat\Pieter

ssc install winsor2
ssc install asdoc
ssc install outreg2
 
// load data

*** country specific
import excel "DATA1.xlsx", sheet("GDp per capita") firstrow clear
reshape long GDP , i(date) j(target_nation) string
gen y = year(date)
drop date
replace GDP = ln(GDP)
ren GDP GDP_per_capita
ren target_nation TargetNation
save dta\GDP_per_capita, replace

import excel "DATA1.xlsx", sheet("GDP growth") firstrow clear
reshape long GDP , i(target_nation) j(y)
ren target_nation TargetNation
replace TargetNation = "USA" if TargetNation == "United States"
save dta\GDP, replace

import excel "DATA1.xlsx", sheet("Quality of institutions") firstrow clear
reshape long QOI , i(target_nation) j(y)
ren target_nation TargetNation
save dta\QOI, replace

import excel "DATA1.xlsx", sheet("quarterly interest") firstrow clear
reshape long interest , i(date) j(target_nation) string
gen q = quarter(date)
gen y = year(date)
drop date
ren target_nation TargetNation
save dta\interest_target, replace

use dta\interest_target, clear
ren TargetNation AcquirorNation
save dta\interest_acq, replace

import excel "DATA1.xlsx", sheet("monthly exchange rate") firstrow clear
reshape long cur , i(date) j(target_nation) string
gen m = month(date)
gen y = year(date)
drop date
gen ym = ym(y,m)
egen tn = group(target_nation)
xtset tn ym
gen cur_change = 100*((cur-l.cur)/l.cur)
replace cur_change = 1 if cur_change == .
ren target_nation TargetNation
save dta\cur, replace

import excel "DATA1.xlsx", sheet("EPU") firstrow clear
destring Year, replace force
gen ym = ym(Year,Month)
format ym %tm
tsset ym
dfuller EPUUSA
dfuller EPUCana
dfuller EPUChina
dfuller EPUMexico
dfuller EPUEurope
tsline EPUUSA
tsline EPUChina
reshape long EPU , i(ym) j(target_nation) string
ren target_nation TargetNation
save dta\EPU, replace

import excel "DATA1.xlsx", sheet("consumer price index") firstrow clear
reshape long CPI , i(date) j(target_nation) string
gen q = quarter(date)
gen y = year(date)
drop date
ren target_nation TargetNation
save dta\cpi, replace

*** us company is target
import excel "DATA1.xlsx", sheet("UStarget") firstrow clear
drop if AcquirorNation == "Jersey"
drop if AcquirorNation == "Ireland"
drop if AcquirorNation == "Guernsey"
drop if AcquirorNation == "Isle of Man"
replace AcquirorNation = "China" if AcquirorNation == "China (Mainland)"
tab AcquirorNation 
drop if AcquirorNation == "United States"
replace AcquirorNation = "Europe" if (AcquirorNation != "China") & (AcquirorNation != "Mexico") & (AcquirorNation != "Canada") & (AcquirorNation != "United States")
save dta/ustarget, replace

*** US company is acquiror
import excel "DATA1.xlsx", sheet("usacquiror") firstrow clear
drop TargetStockPrice4WeeksAfter
replace TargetNation = "China" if TargetNation == "China (Mainland)"
replace TargetNation = "Europe" if TargetNation == "United Kingdom"
replace TargetNation = "USA" if TargetNation == "United States"
drop if TargetNation == "Jersey"
drop if TargetNation == "Guernsey"
drop if TargetNation == "Isle of Man"
replace TargetNation = "Europe" if (TargetNation != "China") & (TargetNation != "Mexico") & (TargetNation != "Canada") & (TargetNation != "USA")

*append using dta/europe_target // europe + rest
append using dta/ustarget // us target erbij
replace TargetNation = "USA" if TargetNation == "United States"
replace AcquirorNation = "USA" if AcquirorNation == "United States"

gen type = ""
replace type = "US_US" if TargetNation == "USA" & AcquirorNation == "USA"
replace type = "US_target" if TargetNation == "USA" & AcquirorNation != "USA"
replace type = "US_acq" if TargetNation != "USA" & AcquirorNation == "USA"
tab type

duplicates drop

*** overzicht over landen
tab TargetNation AcquirorNation

*** dates/time
rename NumberOfDaysBetweenDateAnnou Ndays
replace Ndays = . if Ndays == -999

gen date_completed = DateOriginallyAnnounced+Ndays
replace date_completed = DateOriginallyAnnounced if Ndays == .
format date_completed %td

gen y = year(date_completed)
gen q = quarter(date_completed)
gen m = month(date_completed)
drop if y < 2014 | y > 2022

* (control) variables

* total assets (acquiror)
rename AcquirorTotalAssetsLast12Mon TA_acq
gen ln_assets = ln(TA_acq)
* Market-to-book ratio
gen MBratio = AcquirorStockPriceonAnnouncem/AcquirorBookValueperShareLas
sum MBratio, d
* leverage by assets
sum LeveragetoAssets, d
* cash by assets
gen cash = CashHoldingstototalassets
*replace cash = CashHoldingstototalassets if cash == .
drop CashHoldingstototalassets 

* dependent

* deal value
rename DealValueUSDMillions DealValue
gen DealValue_by_assets = DealValue/ TA_acq
sum DealValue_by_assets, d

*search mvpatterns
mvpatterns ln_assets MBratio LeveragetoAssets cash DealValue
sum ln_assets MBratio LeveragetoAssets cash

gen ln_DealValue_by_assets = ln(DealValue_by_assets)
hist ln_DealValue_by_assets

replace DealStatus = "Withdrawn" if DealStatus != "Completed"

*** merge Canada & Mexico
gen target_nation = TargetNation
replace target_nation = "Canada&Mexico" if target_nation == "Canada" | target_nation == "Mexico"

*** Tariff periods
*drop tariff
gen tariff = "0pre"
replace tariff = "1post" if target_nation == "Canada&Mexico" & ///
 date_completed >= date("2018-03-08","YMD") & date_completed < date("2019-05-20","YMD")
replace tariff = "1post" if target_nation == "China" & date_completed >= date("2018-03-08","YMD")
replace tariff = "1post" if target_nation == "Europe" & ///
 date_completed >= date("2018-03-08","YMD") & date_completed < date("2022-11-31","YMD")
replace tariff = "" if target_nation == "Canada&Mexico" & date_completed >= date("2019-05-20","YMD")
replace tariff = "" if target_nation == "Europe" & date_completed > date("2022-12-01","YMD")
replace tariff = "1post" if target_nation == "USA" & date_completed > date("2018-03-08","YMD")

drop if tariff == ""


gen lagged_tariff = "0pre"
replace lagged_tariff = "1post" if target_nation == "Canada&Mexico" & ///
 date_completed >= date("2018-06-08","YMD") & date_completed < date("2019-08-20","YMD")
replace lagged_tariff = "1post" if target_nation == "China" & date_completed >= date("2018-06-08","YMD")
replace lagged_tariff = "1post" if target_nation == "Europe" & ///
 date_completed >= date("2018-06-08","YMD") & date_completed < date("2023-02-28","YMD")
replace lagged_tariff = "" if target_nation == "Canada&Mexico" & date_completed >= date("2019-08-20","YMD")
replace lagged_tariff = "" if target_nation == "Europe" & date_completed > date("2023-02-01","YMD")
replace lagged_tariff = "1post" if target_nation == "USA" & date_completed > date("2018-06-08","YMD")

save dta\DATA_17142, replace

use dta\DATA_17142, clear
merge m:1 TargetNation y using dta/QOI //, nogen
drop if _merge == 2
drop _merge
gen ym = ym(y,m)
merge m:1 TargetNation ym using dta/EPU
drop if _merge == 2
drop _merge
merge m:1 TargetNation y using dta/GDP_per_capita, nogen

merge m:1 TargetNation y using dta/GDP, nogen
merge m:1 TargetNation y q using dta/interest_target //, nogen
keep if _merge == 3
ren interest interest_target
drop _merge
merge m:1 AcquirorNation y q using dta/interest_acq
keep if _merge == 3
ren interest interest_acq
drop _merge
gen interest_rate_diff = interest_acq - interest_target

merge m:1 TargetNation y m using dta/cur

keep if _merge == 3 | _merge == 1
drop _merge
merge m:1 TargetNation y q using dta/cpi //, nogen
keep if _merge == 3
drop _merge
save dta\DATA_19382B, replace

use dta\DATA_19382B, replace
asdoc tab type DealStatus, save(output/Noverview.doc) replace row
tab type DealStatus , chi2
asdoc tab Year, save(output/Noverview_year.doc) replace
gen payment_method = ""
replace PercentageofCash = 100 if PercentageofCash > 100 & PercentageofCash != .
replace PercentageofStock= 100 if PercentageofStock> 100 & PercentageofStock!= .
replace payment_method = "cash" if PercentageofCash == 100
replace payment_method = "stock" if PercentageofStock == 100
replace payment_method = "mixed" if PercentageofStock < 100 & PercentageofCash == .
replace payment_method = "mixed" if PercentageofCash < 100 & PercentageofStock == .
replace payment_method = "mixed" if PercentageofStock < 100 & PercentageofCash < 100
replace payment_method = "mixed" if PercentageofStock == 100 & PercentageofCash == 100
tab payment_method
sum Percentageof*

gen perc_stock = PercentageofStock
replace perc_stock = 100 - PercentageofCash if PercentageofStock == .
replace perc_stock = 50 if PercentageofStock == 100 & PercentageofCash == 100
sum perc_stock

gen perc_cash = PercentageofCash
replace perc_cash = 100 - PercentageofStock if PercentageofCash == .
replace perc_cash = 50 if PercentageofStock == 100 & PercentageofCash == 100
replace perc_stock = . if DealStatus == "Withdrawn" 
replace perc_cash = . if DealStatus == "Withdrawn" 

drop if y > 2022

asdoc tab type payment_method, save(output/Noverview_payment.doc) replace row
tab type payment_method, chi2

asdoc tab tariff target_nation, save(output/n_obs_over_period_nation.doc) replace row
tab y target_nation

asdoc tab y type,  save(output/n_obs_over_year.doc) replace row

asdoc tab tariff target_nation, save(output/n_obs_over_period_nation.doc) replace 

save dta/DATA_all, replace
use dta/DATA_all, clear

gen Public = (AcquirorPublicStatus == "Public")
replace Public = . if (AcquirorPublicStatus == "")
gen diversify = DIF_SIC1SAMESIC0
rename PremiumPaid4WeeksPriortoA DealPremium
gen ln_DealValue = ln(DealValue)
gen dWithdrawn = DealStatus == "Withdrawn"
gen dStock100 = perc_stock == 100
replace dStock100 = . if perc_stock == .
replace cur_change = 0 if TargetNation == "USA"

sum DealPremium, d
winsor2 DealPremium, cuts(1 99) replace

mvpatterns DealPremium ln_assets MBratio cash LeveragetoAssets Public diversify ln_DealValue

* drop missing values
qui reg  ln_assets MBratio cash LeveragetoAssets Public diversify ln_DealValue
keep if e(sample) // drop cases with missing data (listwise)

sum ln_assets MBratio cash LeveragetoAssets, d
winsor2 MBratio cash LeveragetoAssets, cuts(2.5 97.5) replace

global countryvars GDP_per_capita GDP interest_rate_diff cur_change CPI EPU QOI
global controlvars ln_assets MBratio cash LeveragetoAssets Public diversify  ln_DealValue  

bys DealStatus: asdoc sum DealPremium perc_stock dStock100  dWithdrawn $controlvars $countryvars, stat(count mean sd min p25 median p75 max) replace save(output/descriptives_overall.doc)
bys tariff: asdoc sum DealPremium perc_stock dStock100 dWithdrawn $controlvars $countryvars, stat(count mean sd min p25 median p75 max) save(output/descriptives_overall.doc)

asdoc pwcorr DealPremium perc_stock dStock100 dWithdrawn $controlvars $countryvars , stat(count mean sd min p25 median p75 max) replace save(output/correlations_overall.doc) star(all) dec(2)

foreach var of varlist DealPremium perc_stock dStock100 dWithdrawn $controlvars $countryvars {
asdoc ttest `var' if type == "US_acq", by(tariff) save(output/ttests_us_acq.doc) rowappend
}
foreach var of varlist DealPremium perc_stock dStock100 dWithdrawn $controlvars $countryvars {
asdoc ttest `var' if type == "US_target", by(tariff) save(output/ttests_us_target.doc) rowappend
}

egen Tariff = group(tariff)
egen lagged_Tariff = group(lagged_tariff)
drop target_nation
egen target_nation = group(TargetNation), label
egen Type = group(type), label

label define tariff_label 1 "pre" 2 "post"
label values Tariff tariff_label
label values lagged_Tariff tariff_label
************************************************* STANDARDIZED *****************************************
* standardized betas: X en Y beide standardized
reg DealPremium i.Tariff i.Type i.Tariff#i.Type $controlvars i.Year if DealStatus != "Withdrawn"  & type != "US_target" , robust beta
outreg2 using voorbeeld_betas.doc, stat(beta)
summarize DealPremium if e(sample)
gen ZDealPremium = (DealPremium - r(mean)) / r(sd) 
* alleen afhankelijke
reg ZDealPremium i.Tariff i.Type i.Tariff#i.Type $controlvars i.Year if DealStatus != "Withdrawn"  & type != "US_target" , robust
outreg2 using voorbeeld_standardized.doc
************************************************* EIND STANDARDIZED *****************************************
testparm i.Year
outreg2 using output/H1.doc, replace label
margins i.Tariff#i.Type 
marginsplot
rvfplot
reg DealPremium i.lagged_Tariff i.Type i.lagged_Tariff#i.Type  $controlvars i.Year if DealStatus != "Withdrawn"  & type != "US_target" , robust
outreg2 using output/H1.doc, label

reg DealPremium i.Tariff i.Type i.Tariff#i.Type $controlvars if DealStatus != "Withdrawn"  & type != "US_target" , robust
outreg2 using output/H1_zonderYearFE.doc, replace label
asdoc vif, save(vif_H1) replace
margins i.Tariff#i.Type 
marginsplot
rvfplot
reg DealPremium i.lagged_Tariff i.Type i.lagged_Tariff#i.Type  $controlvars if DealStatus != "Withdrawn"  & type != "US_target" , robust
outreg2 using output/H1_zonderYearFE.doc, label
asdoc vif, save(vif_H1) 

//withdrawn
tab DealStatus dWithdrawn

logit dWithdrawn i.Tariff i.Type i.Tariff#i.Type $controlvars i.Year if type != "US_target"
outreg2 using output/H2.doc, replace label
outreg2 using output/H2.doc, eform ctitle("odds ratio") label
logit dWithdrawn i.lagged_Tariff i.Type i.lagged_Tariff#i.Type $controlvars i.Year if type != "US_target"
outreg2 using output/H2.doc, label
outreg2 using output/H2.doc, eform ctitle("odds ratio") label
margins i.lagged_Tariff#i.Type 
marginsplot

// zonder year FE
logit dWithdrawn i.Tariff i.Type i.Tariff#i.Type $controlvars if type != "US_target"
outreg2 using output/H2_zonderYearFE.doc, replace label
outreg2 using output/H2_zonderYearFE.doc, eform ctitle("odds ratio") label
logit dWithdrawn i.lagged_Tariff i.Type i.lagged_Tariff#i.Type $controlvars if type != "US_target"
outreg2 using output/H2_zonderYearFE.doc, label
outreg2 using output/H2_zonderYearFE.doc, eform ctitle("odds ratio") label
margins i.lagged_Tariff#i.Type 
marginsplot

qui reg dWithdrawn i.Tariff i.Type i.Tariff#i.Type $controlvars if type != "US_target" 
asdoc vif, save(vif_H2) replace

// payment
reg perc_stock i.Tariff i.Type i.Tariff#i.Type $controlvars if DealStatus != "Withdrawn"  & type != "US_target"  , robust
asdoc vif, save(vif_H3) replace
outreg2 using output/H3.doc, replace label
logit dStock100 i.Tariff i.Type i.Tariff#i.Type $controlvars if DealStatus != "Withdrawn"  & type != "US_target" 
outreg2 using output/H3.doc, label
outreg2 using output/H3.doc, eform label
reg perc_stock i.lagged_Tariff i.Type i.lagged_Tariff#i.Type $controlvars if DealStatus != "Withdrawn"  & type != "US_target"  , robust
asdoc vif, save(vif_H3)
outreg2 using output/H3.doc, label
logit dStock100 i.lagged_Tariff  i.Type i.lagged_Tariff#i.Type $controlvars if DealStatus != "Withdrawn"  & type != "US_target" 
outreg2 using output/H3.doc, label
outreg2 using output/H3.doc, eform label

reg perc_stock i.Tariff i.Type i.Tariff#i.Type $controlvars i.Year if DealStatus != "Withdrawn"  & type != "US_target"  , robust
outreg2 using output/H3_yearFE.doc, label replace
asdoc vif, save(vif_H3)
logit dStock100 i.Tariff i.Type i.Tariff#i.Type $controlvars i.Year  if DealStatus != "Withdrawn"  & type != "US_target" 
outreg2 using output/H3_yearFE.doc, label
outreg2 using output/H3_yearFE.doc, eform label
reg perc_stock i.lagged_Tariff i.Type i.lagged_Tariff#i.Type $controlvars i.Year  if DealStatus != "Withdrawn"  & type != "US_target"  , robust
asdoc vif, save(vif_H3)
outreg2 using output/H3_yearFE.doc, label
logit dStock100 i.lagged_Tariff  i.Type i.lagged_Tariff#i.Type $controlvars i.Year if DealStatus != "Withdrawn"  & type != "US_target" 
outreg2 using output/H3_yearFE.doc, label
outreg2 using output/H3_yearFE.doc, eform label


// voor data bij 0 deals invullen (per yq)
use dta\cur, clear 
replace cur_change = cur_change/100 
replace cur_change = cur_change + 1
gen cur_change3 = cur_change * l.cur_change * l2.cur_change
keep if m == 3 | m==6 | m == 9 | m == 12
gen yq = qofd(dofm(ym))
keep TargetNation yq cur_change3
format yq %tq
rename cur_change3 cur_change
save dta/cur_change_q, replace

use dta\EPU, clear
gen yq = qofd(dofm(ym))
collapse EPU, by(yq TargetNation)
save dta/EPU_q, replace


// voor acquiror = US
use dta/DATA_all, clear
gen n = 1
gen yq = yq(y,q)

drop if type == "US_target"

global countryvars GDP_per_capita GDP interest_rate_diff cur_change CPI EPU QOI

egen Tariff = group(tariff)
egen lagged_Tariff = group(lagged_tariff)
label define tariff_label 1 "pre" 2 "post"

collapse (sum) n=n (min) Tariff lagged_Tariff (mean) $countryvars  , by(TargetNation  yq)
label values lagged_Tariff Tariff tariff_label
egen nation = group(TargetNation), label
xtset nation yq
tsfill, full
drop if nation == 1 & yq >= 238
drop if nation == 4 & yq >= 238
replace n = 0 if n == .
replace Tariff = 2 if l.Tariff == 2
replace Tariff = 2 if l.Tariff == 2
replace Tariff = 1 if f.Tariff == 1
replace Tariff = 1 if f.Tariff == 1
replace Tariff = 1 if yq < 233 & nation == 4
replace lagged_Tariff = l.Tariff if lagged_Tariff == .

gen y = yofd(dofq(yq))
replace TargetNation = "Canada" if nation == 1
replace TargetNation = "China" if nation == 2
replace TargetNation = "Europe" if nation == 3
replace TargetNation = "Mexico" if nation == 4
replace TargetNation = "USA" if nation == 5

merge m:1 y TargetNation using dta/GDP, update  //keep(matched)
drop if _merge == 2
drop _merge
merge m:1 y TargetNation using dta/QOI, update //keep(matched)
drop if _merge == 2
drop _merge
merge m:1 y TargetNation using dta/GDP_per_capita, update  //keep(matched)
drop if _merge == 2
drop _merge
merge 1:1 yq TargetNation using dta/cur_change_q, update  //keep(matched)
drop if _merge == 2
drop _merge
replace cur_change = 1 if nation == 5
merge 1:1 yq TargetNation using dta/EPU_q, update  //keep(matched)
drop if _merge == 2
drop _merge
merge 1:1 yq TargetNation using dta/cpi, update  //keep(matched)
drop if _merge == 2
drop _merge

gen AcquirorNation = "USA"
merge m:1 AcquirorNation y q using dta/interest_acq
drop if _merge == 2
drop _merge
ren interest int_US
merge 1:1 yq TargetNation using dta/interest_target  //keep(matched)
drop if _merge == 2
drop _merge
replace interest_rate_diff = int_US - interest if interest_rate_diff == .

format yq %tq
tab yq Tariff
replace cur_change = 0 if TargetNation == "USA"
replace TargetNation = "0USA" if TargetNation == "USA"
egen target_nation = group(TargetNation), label

bys target_nation: sum n

poisson n i.Tariff i.target_nation Tariff#i.target_nation $countryvars 
outreg2 using output/H4.doc, replace label
poisson n i.lagged_Tariff i.target_nation lagged_Tariff#i.target_nation $countryvars
outreg2 using output/H4.doc, label
nbreg n i.Tariff i.target_nation Tariff#i.target_nation $countryvars
outreg2 using output/H4.doc, label
outreg2 using output/H4.doc, label stat(coef) eform
nbreg n i.lagged_Tariff i.target_nation lagged_Tariff#i.target_nation $countryvars
outreg2 using output/H4.doc, label
outreg2 using output/H4.doc, label stat(coef) eform

save dta/H4, replace

// robustness check

use dta/DATA_all, clear
gen n = 1
gen yq = yq(y,q)

drop if type == "US_acq"

global countryvars GDP_per_capita GDP interest_rate_diff cur_change CPI EPU QOI

egen Tariff = group(tariff)
egen lagged_Tariff = group(lagged_tariff)
label define tariff_label 1 "pre" 2 "post"

tab AcquirorNation type
tab TargetNation type
drop TargetNation
ren AcquirorNation TargetNation

collapse (sum) n=n (min) Tariff lagged_Tariff (mean) $countryvars  , by(TargetNation yq)
label values lagged_Tariff Tariff tariff_label
egen nation = group(TargetNation), label
xtset nation yq
tsfill, full
drop if nation == 1 & yq >= 238
drop if nation == 4 & yq >= 238
replace n = 0 if n == .
replace Tariff = 2 if l.Tariff == 2
replace Tariff = 2 if l.Tariff == 2
replace Tariff = 1 if f.Tariff == 1
replace Tariff = 1 if f.Tariff == 1
replace Tariff = 1 if yq < 233 & nation == 4
replace lagged_Tariff = l.Tariff if lagged_Tariff == .

gen y = yofd(dofq(yq))
replace TargetNation = "Canada" if nation == 1
replace TargetNation = "China" if nation == 2
replace TargetNation = "Europe" if nation == 3
replace TargetNation = "Mexico" if nation == 4
replace TargetNation = "USA" if nation == 5

merge m:1 y TargetNation using dta/GDP, update  //keep(matched)
drop if _merge == 2
drop _merge
merge m:1 y TargetNation using dta/QOI, update //keep(matched)
drop if _merge == 2
drop _merge
merge m:1 y TargetNation using dta/GDP_per_capita, update  //keep(matched)
drop if _merge == 2
drop _merge
merge 1:1 yq TargetNation using dta/cur_change_q, update  //keep(matched)
drop if _merge == 2
drop _merge
replace cur_change = 1 if nation == 5
merge 1:1 yq TargetNation using dta/EPU_q, update  //keep(matched)
drop if _merge == 2
drop _merge
merge 1:1 yq TargetNation using dta/cpi, update  //keep(matched)
drop if _merge == 2
drop _merge

ren TargetNation AcquirorNation 
gen TargetNation = "USA" //gen AcquirorNation = "USA"

merge m:1 AcquirorNation y q using dta/interest_acq
drop if _merge == 2
drop _merge
ren interest int_US
merge m:1 yq TargetNation using dta/interest_target  //keep(matched)
drop if _merge == 2
drop _merge
replace interest_rate_diff = int_US - interest if interest_rate_diff == .

format yq %tq
tab yq Tariff
replace cur_change = 0 if AcquirorNation == "USA"
replace AcquirorNation = "0USA" if AcquirorNation == "USA"
egen acq_nation = group(AcquirorNation), label

bys acq_nation: sum n

xtline n if acq_nation >1 , overlay

poisson n i.Tariff i.acq_nation Tariff#i.acq_nation $countryvars 
outreg2 using output/robustness.doc, replace label
nbreg n i.Tariff i.acq_nation Tariff#i.acq_nation $countryvars
outreg2 using output/robustness.doc, label
outreg2 using output/robustness.doc, label stat(coef) eform

save dta/robustness, replace





















// canada & mexico
gen tmp = group_tariff if TargetNation == "Canada"
bys yq: egen tmp2 = mean(tmp) 
replace group_tariff = tmp2 if TargetNation == "0USA"

poisson n i.group_tariff i.target_nation group_tariff#i.target_nation $cvars if TargetNation == "Canada" | TargetNation == "Mexico" | TargetNation == "0USA"
outreg2 using output/H4.doc, ctitle("Canada & Mexico") replace label
margins group_tariff#i.target_nation 
marginsplot

drop tmp tmp2
// europe
gen tmp = group_tariff if TargetNation == "Europe"
bys yq: egen tmp2 = mean(tmp) 
replace group_tariff = tmp2 if TargetNation == "0USA"

poisson n i.group_tariff i.target_nation group_tariff#i.target_nation $cvars if TargetNation == "Europe" | TargetNation == "0USA"
outreg2 using output/H4.doc, ctitle("Europe") label
margins group_tariff#i.target_nation 
marginsplot

drop tmp tmp2
// china
gen tmp = group_tariff if TargetNation == "China"
bys yq: egen tmp2 = mean(tmp) 
replace group_tariff = tmp2 if TargetNation == "0USA"

poisson n i.group_tariff i.target_nation group_tariff#i.target_nation $cvars if TargetNation == "China" | TargetNation == "0USA"
outreg2 using output/H4.doc, ctitle("China") label
margins group_tariff#i.target_nation 
marginsplot

// code om te checken
tab yq group_tariff if TargetNation == "0USA" | TargetNation == "China"







poisson n group_tariff##i.target_nation $cvars
poisson n i.group_tariff i.target_nation group_tariff#i.target_nation $cvars



nbreg n i.group_tariff i.target_nation  group_tariff#i.target_nation $cvars



*** descriptive + correlation overall H1 & H2
drop if type == "US_target"

qui reg ln_assets MBratio cash LeveragetoAssets
keep if e(sample) // drop cases with missing data (listwise)
winsor2 MBratio cash Leveragetoassets, cuts(2.5 97.5) replace

drop nation_period

tab nation_period
asdoc tabstat DealValue, by(nation_period) stat(count median) save(output/n_obs_over_period_nation.doc) replace

save 20250325DATA, replace


mvpatterns ln_DealValue_by_assets ln_assets MBratio cash LeveragetoAssetstariff GDP_per_capita GDP interest_rate_diff cur_change CPI

*** H1
use dta\DATA_19342, replace
keep if DealStatus == "Completed"


bys target_nation: tab y tariff



bys target_nation: asdoc sum ln_DealValue_by_assets ln_assets MBratio cash Leveragetoassets, stat(count mean sd min p25 median p75 max) replace save(output/descriptives_H1.doc)
asdoc sum ln_DealValue_by_assets ln_assets MBratio cash Leveragetoassets, stat(count mean sd min p25 median p75 max) save(output/descriptives_H1.doc)
bys tariff: asdoc sum ln_DealValue_by_assets ln_assets MBratio cash Leveragetoassets, stat(count mean sd min p25 median p75 max) save(output/descriptives_H1.doc)
asdoc tab target_nation tariff, row

gen dEurope = target_nation == "Europe"
gen dEurope = target_nation == "Europe"

asdoc pwcorr ln_DealValue_by_assets ln_assets MBratio cash Leveragetoassets, star(all) save(output/correlations_H1.doc) replace



reg ln_DealValue_by_assets i.target_nation ln_assets MBratio cash LeveragetoAssetsif type == "

