--declare @year int; set @year = 2022
--declare @quarter int; set @quarter = 3
--declare @carrierId varchar(3); set @carrierId = 'JHL'

use Magnastar

declare @startDate datetime
declare @endDate datetime
set @startdate = convert(datetime, str(@year))
set @enddate = dateadd(ms, -3*@quarter, dateadd(mm, 3*@quarter, @startdate))

--TRIAL BALANCE
select
  CASE WHEN COA.carrID IS NULL THEN m.carrierID ELSE COA.carrID END as carrier
 ,CASE WHEN COA.accountNumber IS NULL THEN m.cpaAccountNumber ELSE COA.accountNumber END as accountNumber
 ,COA.accountName
 ,coalesce(sum(m.amount),0) as amount
 
from (	select *
		from dbo.MtdJournal
		where JournalDate between @startDate and @endDate and carrierId = @carrierId
		) as m
	full outer join (	select carrID, accountNumber, accountName
						from dbo.[McCamish COA]
						group by carrID, accountNumber, accountName
					) as COA on COA.accountNumber = m.cpaAccountNumber and COA.carrID = m.carrierID

where m.carrierID = @carrierId or coa.carrID = @carrierID
group by CASE WHEN COA.carrID IS NULL THEN m.carrierID ELSE COA.carrID END
		,CASE WHEN COA.accountNumber IS NULL THEN m.cpaAccountNumber ELSE COA.accountNumber END
		,COA.accountName
order by CASE WHEN COA.accountNumber IS NULL THEN m.cpaAccountNumber ELSE COA.accountNumber END




--OVERHEAD
declare @priorYear varchar(4)
declare @priorQuarter int

IF (@quarter > 1) SELECT @priorYear = @year, @priorQuarter = @Quarter -1
ELSE SELECT @priorYear = convert(varchar(4),convert(int,@year)-1), @priorQuarter = 4

set @startdate = dateadd(q, @quarter - 1, convert(datetime, str(@year)))
set @enddate = dateadd(ms, -3, dateadd(mm, 3, @startdate))

IF (@carrierID = 'PR1' or @carrierID = 'PR2')

(select
  f00.carrierId
 ,cast(f00.issueYear as int)
 ,@quarter
 ,count(f00.policyNumber) as polCount
 ,sum(f00.specifiedAmount) as faceSum100
 ,sum(f00.specifiedAmount*(coalesce(n.modco100,.5))) as faceSumModco
 ,sum(F02.fundValue) as avSum100
 ,sum(F02.fundValue*(coalesce(n.modco100,.5))) as avSumModco
 ,convert(float, max(expenses.overheadPerQuarter))*convert(float, sum(f00.specifiedAmount*(coalesce(n.modco100,.5))))/convert(float, total) as overhead
 ,CASE WHEN (f00.issueYear = '2002' and f00.issueMonth < '07') or (f00.issueYear < '2002') THEN 'S' 
	   WHEN (f00.issueYear > '2008') OR (f00.issueYear ='2008' AND f00.issuemonth > '03') THEN 'M'
	   ELSE 'SM' 
  END as swissMIndicator


from dbo.vF00BaseCoverageQuarterly f00
  left outer join dbo.F02MonthiversaryRecord f02 
    on f02.carrierId = f00.carrierId and f02.policyNumber = f00.policyNumber
      and f02.activityYear = f00.activityYear and f02.activityMonth = f00.activityMonth
  inner join expenses on expenses.carrierId = f00.carrierId and @endDate between expenses.startDate and expenses.endDate
   		inner join (	

						select cID.carrierID, sum(f.specifiedAmount*coalesce(n.modco100,.5)) as TOTAL
						from dbo.vF00BaseCoverageQuarterly f
							inner join dbo.CarrierIDAssociations cID on cID.alternateID = f.carrierID 
							left outer join dbo.nonmodco n on f.policyNumber = n.policyNumber
						where cID.carrierID = @carrierId and activityYear = @year and activityQuarter = @quarter
						group by cID.carrierId) as groupedF00 on f00.carrierID = groupedF00.carrierID
 
left outer join dbo.nonmodco n on f00.policyNumber = n.policyNumber
where f00.carrierId = @carrierId
  and f00.activityYear = @year
  and f00.activityQuarter = @quarter
group by f00.carrierID, f00.issueYear, CASE WHEN (f00.issueYear = '2002' and f00.issueMonth < '07') or (f00.issueYear < '2002') THEN 'S' WHEN (f00.issueYear > '2008') OR (f00.issueYear ='2008' AND f00.issuemonth > '03') THEN 'M' ELSE 'SM' END, total 
)

ELSE
(select
  f00.carrierId
 ,cast(f00.issueYear as int)
 ,@quarter
 ,count(f00.policyNumber) as polCount
 ,sum(f00.specifiedAmount) as faceSum100
 ,sum(f00.specifiedAmount*(coalesce(n.modco100,.5))) as faceSumModco
 ,sum(F02.fundValue) as avSum100
 ,sum(F02.fundValue*(coalesce(n.modco100,.5))) as avSumModco
 ,convert(float, max(expenses.overheadPerQuarter))*convert(float, sum(f00.specifiedAmount*(coalesce(n.modco100,.5))))/convert(float, total) as overhead

from dbo.vF00BaseCoverageQuarterly f00
  left outer join dbo.F02MonthiversaryRecord f02 
    on f02.carrierId = f00.carrierId and f02.policyNumber = f00.policyNumber
      and f02.activityYear = f00.activityYear and f02.activityMonth = f00.activityMonth
  inner join expenses on expenses.carrierId = f00.carrierId and @endDate between expenses.startDate and expenses.endDate
  		inner join (	
						select cID.carrierID, sum(f.specifiedAmount*coalesce(n.modco100,.5)) as TOTAL
						from dbo.vF00BaseCoverageQuarterly f
							inner join dbo.CarrierIDAssociations cID on cID.alternateID = f.carrierID 
							left outer join dbo.nonmodco n on f.policyNumber = n.policyNumber
						where cID.carrierID = @carrierId and activityYear = @year and activityQuarter = @quarter
						group by cID.carrierId) as groupedF00 on f00.carrierID = groupedF00.carrierID
  
left outer join dbo.nonmodco n on f00.policyNumber = n.policyNumber

where f00.carrierId = @carrierId
  and f00.activityYear = @year
  and f00.activityQuarter = @quarter
group by f00.carrierID, f00.issueYear, total
)
order by carrierId, issueYear




--Quarterly Data
DECLARE @datebilled1 varchar(6)
DECLARE @datebilled2 varchar(6)
DECLARE @datebilled3 varchar(6)

IF (@quarter > 1) SELECT @priorYear = @year, @priorQuarter = @Quarter -1
ELSE SELECT @priorYear = convert(varchar(4),convert(int,@year)-1), @priorQuarter = 4

set @startdate = dateadd(q, @quarter - 1, convert(datetime, str(@year)))
set @enddate = dateadd(ms, -3, dateadd(mm, 3, @startdate))


select
  f00.carrierId
 ,rtrim(f00.policyNumber) policyNumber
 ,f00.recordType
 ,f00.activityYear
 ,f00.activityMonth
 ,f00.planCode
 ,f00.issueAge
 ,f00.sexCode
 ,f00.smokingAndRiskCategory
 ,f00.substandardClass
 ,f00.secondaryInsuredIssueAge
 ,f00.secondaryInsuredSexCode
 ,f00.secondaryInsuredSmokingAndRiskCategory
 ,f00.secondaryInsuredSubstandardClass
 ,f00.secondaryInsuredStatus
 ,f00.deathBenifitOption as DB_Option
 ,f00.doliTest
 ,f00.issueDay
 ,f00.issueYear
 ,f00.issueMonth
 ,f00.disablementYear
 ,f00.disablementMonth
 ,f00.currentStatus
 ,f00.modeOfPayment
 ,f00.specifiedAmount
 ,f00.targetPremiumComm1
 ,f00.targetPremiumcomm2
 ,f00.sevenPayPremium
 ,f00.issueState
 ,f00.policyType
 ,f00.caseNumber
 ,f00.caseVersion
 ,coalesce(nm.modco100, .5) as [Modco%]
 --******************************************************************************
--5/22/2008 MRJ: This code incorrectly incorrectly caps the premium at the 10 pay
--				This code should include all premium paid
--				Changed to use j instead of pyp 

--,coalesce(pyp.firstYear, 0) as firstYearPremiums
--,coalesce(pyp.renewal, 0) as renewalPremiums
--******************************************************************************

,coalesce(j.[30101.1], 0) as firstYearPremiums
,coalesce(j.[30102.1], 0) as renewalPremiums


 --max(min(f00.sevenPayPremium  - priorPeriodFirstYearPremiums, firstYearPremiums),0) * percentOfFirstYearPremiumExpenseDeferable
 ,coalesce(pyp.firstYear, 0)
     * expenses.percentOfFirstYearPremiumExpenseDeferable as expensePercentOfFirstYearPremiumDeferrable

 --max(min(f00.sevenPayPremium  - priorPeriodFirstYearPremiums, firstYearPremiums),0) * percentOfFirstYearPremiumExpenseRecurring
 ,coalesce(pyp.firstYear, 0)
     * expenses.percentOfFirstYearPremiumExpenseRecurring as expensePercentOfFirstYearPremiumRecurring

 --max(min(f00.sevenPayPremium  - priorPeriodRenewalPremiums, renewalPremiums),0) * percentOfRenewalPremiumExpense
 ,coalesce(pyp.renewal, 0)
     * expenses.percentOfRenewalPremiumExpense as expensePercentOfRenewalPremium

--******************************************************************************
--5/13/2008 MRJ: This code incorrectly calls fees.percentOfAssetsPerQuarter
--				This code needs to call expenses.percentOfAssetsPerQuarter
--				Changed to call correct table

-- --fees.percentOfAssetsPerQuarter * (IF f00.currentStatus = a then f02.fundValue else 0)
-- ,fees.percentOfAssetsPerQuarter * case f00.currentStatus when 'a' then f02.fundValue else 0 end as expensePercentOfAssets
--******************************************************************************

--expenses.percentOfAssetsPerQuarter * (IF f00.currentStatus = a then f02.fundValue else 0)
 ,expenses.percentOfAssetsPerQuarter * case f00.currentStatus when 'a' then f02.fundValue else 0 end as expensePercentOfAssets


  --IF policy existed in prior period then 0 else (appropriate smokingExpense based on f00.specifiedAmount and f00.smokingAndRiskCategory) * (if single then 1 elsif double then 2)
 ,CASE WHEN pf02.fundValue IS NOT NULL THEN 0 ELSE 
    CASE
      WHEN (f00.specifiedAmount <= 500000) AND (f00.smokingAndRiskCategory IN (4,5)) THEN expenses.smoking45Under500k
      WHEN (f00.specifiedAmount > 500000) AND (f00.specifiedAmount < 1000000) AND (f00.smokingAndRiskCategory IN (4,5)) THEN expenses.smoking45500kTo1m
	  WHEN (f00.specifiedAmount >= 1000000) AND (f00.smokingAndRiskCategory IN (4,5)) THEN expenses.smoking45Over1m
      WHEN (f00.specifiedAmount <= 5000000) AND (f00.smokingAndRiskCategory IN (1,2,3)) THEN expenses.smoking123Under5m
      WHEN (f00.specifiedAmount > 5000000) AND (f00.smokingAndRiskCategory IN (1,2,3)) THEN expenses.smoking123Over5m
      ELSE NULL
    END
  END * CASE WHEN f00.secondaryInsuredIssueAge IS NOT NULL THEN 2 ELSE 1 END as expensePerPolicyDeferrable

 ,expenses.perPolicyrecurringPerQuarter as expensePerPolicyRecurring

 ,expenses.rbcPercentOfAssetsPerQuarter * F02.fundValue as expenseRbcRent

 --max(min(f00.sevenPayPremium  - priorPeriodFirstYearPremiums, firstYearPremiums),0) * deferrablePercentOfFirstYearPremium
 ,coalesce(pyp.firstYear, 0)
     * fees.deferrablePercentOfFirstYearPremium as magnastarFeePercentOfFirstYearPremiumDeferrable

 --max(min(f00.sevenPayPremium  - priorPeriodFirstYearPremiums, firstYearPremiums),0) * recurringPercentOfFirstYearPremium
 ,coalesce(pyp.firstYear, 0)
     * fees.recurringPercentOfFirstYearPremium as magnastarFeePercentOfFirstYearPremiumRecurring


--TODO fee = max{0, min[7-pay premium - min(7-pay premium, premiums paid within the calendar year but prior to the reporting period), premium paid in the reporting period]
 -- OLD max(min(f00.sevenPayPremium  - priorPeriodRenewalPremiums, renewalPremiums),0) * percentOfRenewalPremium
 ,coalesce(pyp.renewal, 0)
     * fees.percentOfRenewalPremium as magnastarFeePercentOfRenewalPremium


--******************************************************************************
--5/13/2008 MRJ: This code incorrectly sets the %Asset fee as the fee %
--				This code needs to set the %Asset fee as [fee % * Asset]
--				Changed to correct calculation

-- ,fees.percentOfAssetsPerQuarter as magnastarFeePercentOfAssets
--******************************************************************************

,fees.percentOfAssetsPerQuarter * case f00.currentStatus when 'a' then f02.fundValue else 0 end as magnastarFeePercentOfAssets

 
--******************************************************************************
--5/13/2008 MRJ: This code incorrectly charges this fee every quarter
--				This code needs to charge this fee at issue
--				Changed to correct calculation

-- ,fees.perPolicyAtIssue as magnastarFeePerPolicyDeferrable
--******************************************************************************

--IF policy existed in prior period then 0 else magnastarFeePerPolicyDeferrable
 ,CASE WHEN pf02.fundValue IS NOT NULL THEN 0 ELSE fees.perPolicyAtIssue
  END as expensePerPolicyDeferrable

 
 ,fees.perPolicyPerQuarter as magnastarFeePerPolicyRecurring
 ,coalesce(j.[71001.1], 0) as commissionsPercentOfFirstYearPremium
 ,coalesce(j.[71002.1], 0) as commissionsPercentOfRenewalPremium
 ,coalesce(j.[71007.1], 0) as commissionsPercentOfAssets
 ,coalesce(j.[71008.1], 0) as commissionsService

 
 ,coalesce(j.[61000.1], 0) as benefitsPaidOnFullSurrender
 ,coalesce(j.[61010.1], 0) as benefitsPaidOnPartialSurrender
 ,coalesce(j.[60500.1], 0) as benefitsPaidOnClaims

 ,F02.fundValue as currentAccountValue

 ,dbo.fnMoneyMax(F02.cashValue - F02.fundValue,0) as currentGeneralAccountReserveForEcv

--******************************************************************************
--5/28/2008 MRJ: This code does not floor the current reserve offset at 0
--				This code needs to floor the current reserve offset at 0
--				Wrapped the existing code in a max function

----F02.fundValue - max(F02.cashValue, polySystems.statutoryReserve)
-- ,F02.fundValue -  dbo.fnMoneyMax(F02.cashValue, uc.statutoryReserve) as currentGeneralAccountReserveOffsetForSurrenderCharges
--******************************************************************************

 --F02.fundValue - max(F02.cashValue, polySystems.statutoryReserve)
 ,dbo.fnMoneyMax(F02.fundValue -  dbo.fnMoneyMax(F02.cashValue, uc.statutoryReserve),0) as currentGeneralAccountReserveOffsetForSurrenderCharges


 --polySystems.unearnedCoi
 ,uc.unearnedCoi as currentGeneralAccountReserveForUnearnedCoi

 ,coalesce(pf02.fundValue,0) as beginningAccountValue

 ,coalesce(dbo.fnMoneyMax(pf02.cashValue - pf02.fundValue,0),0) as beginningAccountReserveForEcv

--******************************************************************************
--5/28/2008 MRJ: This code does not floor the beginning reserve offset at 0
--				This code needs to floor the beginning reserve offset at 0
--				Wrapped the existing code in a max function

----pf02.fundValue - max(pf02.cashValue, Previous_polySystems.statutoryReserve)
-- ,pf02.fundValue - dbo.fnMoneyMax(pf02.cashValue, puc.statutoryReserve) as beginningGeneralAccountReserveOffsetForSurrenderCharges
--******************************************************************************

 --pf02.fundValue - max(pf02.cashValue, Previous_polySystems.statutoryReserve)
 ,dbo.fnMoneyMax(pf02.fundValue - dbo.fnMoneyMax(pf02.cashValue, puc.statutoryReserve),0) as beginningGeneralAccountReserveOffsetForSurrenderCharges


 --Previous_polySystems.unearnedCoi
 ,puc.unearnedCoi as beginningGeneralAccountReserveForUnearnedCoi

 ,va.variableAppreciation

--******************************************************************************
--5/28/2008 MRJ: This code does was left as a "to do" item
--				Added code to calculate earnedInterest
--TODO
-- ,'' as earnedInterest
--******************************************************************************

--GA earned Interest = (GA beginning + GA current)*.5*(Quarterly Interest)
--- Cases: GA Beginning is null, GA Current is null, both null, neither null
 ,CASE WHEN @carrierID LIKE 'JH%' THEN
	(CASE WHEN coalesce(dbo.fnMoneyMax(pf02.cashValue - pf02.fundValue,0),0) + puc.unearnedCoi - dbo.fnMoneyMax(pf02.fundValue - dbo.fnMoneyMax(pf02.cashValue, puc.statutoryReserve),0) IS NULL
            THEN 0 +(dbo.fnMoneyMax(F02.cashValue - F02.fundValue,0) - dbo.fnMoneyMax(F02.fundValue -  dbo.fnMoneyMax(F02.cashValue, uc.statutoryReserve),0))
        WHEN  dbo.fnMoneyMax(F02.cashValue - F02.fundValue,0) + uc.unearnedCoi - dbo.fnMoneyMax(F02.fundValue -  dbo.fnMoneyMax(F02.cashValue, uc.statutoryReserve),0) IS NULL
            THEN 0 + (coalesce(dbo.fnMoneyMax(pf02.cashValue - pf02.fundValue,0),0) - dbo.fnMoneyMax(pf02.fundValue - dbo.fnMoneyMax(pf02.cashValue, puc.statutoryReserve),0))
       WHEN (coalesce(dbo.fnMoneyMax(pf02.cashValue - pf02.fundValue,0),0) + puc.unearnedCoi - dbo.fnMoneyMax(pf02.fundValue - dbo.fnMoneyMax(pf02.cashValue, puc.statutoryReserve),0) + dbo.fnMoneyMax(F02.cashValue - F02.fundValue,0) + uc.unearnedCoi - dbo.fnMoneyMax(F02.fundValue -  dbo.fnMoneyMax(F02.cashValue, uc.statutoryReserve),0))  IS NULL
            THEN 0 
       WHEN (coalesce(dbo.fnMoneyMax(pf02.cashValue - pf02.fundValue,0),0) + puc.unearnedCoi - dbo.fnMoneyMax(pf02.fundValue - dbo.fnMoneyMax(pf02.cashValue, puc.statutoryReserve),0) + dbo.fnMoneyMax(F02.cashValue - F02.fundValue,0) + uc.unearnedCoi - dbo.fnMoneyMax(F02.fundValue -  dbo.fnMoneyMax(F02.cashValue, uc.statutoryReserve),0))  IS NOT NULL
            THEN (coalesce(dbo.fnMoneyMax(pf02.cashValue - pf02.fundValue,0),0) - dbo.fnMoneyMax(pf02.fundValue - dbo.fnMoneyMax(pf02.cashValue, puc.statutoryReserve),0) 
                  + dbo.fnMoneyMax(F02.cashValue - F02.fundValue,0) - dbo.fnMoneyMax(F02.fundValue -  dbo.fnMoneyMax(F02.cashValue, uc.statutoryReserve),0))
	END) ELSE 
	(CASE WHEN coalesce(dbo.fnMoneyMax(pf02.cashValue - pf02.fundValue,0),0) + puc.unearnedCoi - dbo.fnMoneyMax(pf02.fundValue - dbo.fnMoneyMax(pf02.cashValue, puc.statutoryReserve),0) IS NULL
            THEN 0 +(dbo.fnMoneyMax(F02.cashValue - F02.fundValue,0) + uc.unearnedCoi - dbo.fnMoneyMax(F02.fundValue -  dbo.fnMoneyMax(F02.cashValue, uc.statutoryReserve),0))
        WHEN  dbo.fnMoneyMax(F02.cashValue - F02.fundValue,0) + uc.unearnedCoi - dbo.fnMoneyMax(F02.fundValue -  dbo.fnMoneyMax(F02.cashValue, uc.statutoryReserve),0) IS NULL
            THEN 0 + (coalesce(dbo.fnMoneyMax(pf02.cashValue - pf02.fundValue,0),0) + puc.unearnedCoi - dbo.fnMoneyMax(pf02.fundValue - dbo.fnMoneyMax(pf02.cashValue, puc.statutoryReserve),0))
       WHEN (coalesce(dbo.fnMoneyMax(pf02.cashValue - pf02.fundValue,0),0) + puc.unearnedCoi - dbo.fnMoneyMax(pf02.fundValue - dbo.fnMoneyMax(pf02.cashValue, puc.statutoryReserve),0) + dbo.fnMoneyMax(F02.cashValue - F02.fundValue,0) + uc.unearnedCoi - dbo.fnMoneyMax(F02.fundValue -  dbo.fnMoneyMax(F02.cashValue, uc.statutoryReserve),0))  IS NULL
            THEN 0 
       WHEN (coalesce(dbo.fnMoneyMax(pf02.cashValue - pf02.fundValue,0),0) + puc.unearnedCoi - dbo.fnMoneyMax(pf02.fundValue - dbo.fnMoneyMax(pf02.cashValue, puc.statutoryReserve),0) + dbo.fnMoneyMax(F02.cashValue - F02.fundValue,0) + uc.unearnedCoi - dbo.fnMoneyMax(F02.fundValue -  dbo.fnMoneyMax(F02.cashValue, uc.statutoryReserve),0))  IS NOT NULL
            THEN (coalesce(dbo.fnMoneyMax(pf02.cashValue - pf02.fundValue,0),0) + puc.unearnedCoi - dbo.fnMoneyMax(pf02.fundValue - dbo.fnMoneyMax(pf02.cashValue, puc.statutoryReserve),0) 
                  + dbo.fnMoneyMax(F02.cashValue - F02.fundValue,0) + uc.unearnedCoi - dbo.fnMoneyMax(F02.fundValue -  dbo.fnMoneyMax(F02.cashValue, uc.statutoryReserve),0))
	END)
 END *.5 * dbo.EarnedRate.earnedRate - COALESCE(j.[23410.1],0) as earnedInterest



 ,coalesce(j.[35000.1], 0) as breakage
 ,coalesce(j.[55000.1], 0) as breakageExpense
 ,coalesce(j.[30121.1], 0) as premLoadAdditionalSalesChg
 ,coalesce(j.[30190.1], 0) as premLoadAdditionalPremTax
 ,coalesce(j.[30191.1], 0) as premLoadAdditionalDacTax
 ,coalesce(j.[20220.1], 0) as transferNetPremiumToSa
 ,coalesce(j.[20600.1], 0) as transferFromSaDeathClaim
 ,coalesce(j.[20620.1], 0) as transferFromSaPartialSurrender
 ,coalesce(j.[20610.1], 0) as transferFromSaSurrender
 ,coalesce(j.[20410.1], 0) as transferFromSaMAndEFeeDed
 ,coalesce(j.[20310.1], 0) as transferFromSaCoiBaseCoi
 ,coalesce(j.[20320.1], 0) as transferFromSaCoiTermCoi
 ,coalesce(j.[20330.1], 0) as transferFromSaCoiBenefitCoi
 ,coalesce(j.[20340.1], 0) as transferFromSaCoiExtraCoi
 ,coalesce(j.[20450.1], 0) as transferFromSaXferFee
 ,coalesce(j.[20420.1], 0) as transferFromSaSalesChgDed
 ,coalesce(j.[20460.1], 0) as transferFromSaIssueCharge
 ,coalesce(j.[20700.1], 0) as transferFromSaLoan
 ,coalesce(j.[20710.1], 0) as transferFromSaLoanInterest
 ,coalesce(j.[20720.1], 0) as transferToSaLoanEarnings
 ,coalesce(j.[23102.1], 0) as reserveLoanFund
 ,coalesce(j.[40700.1], 0) as policyLoanInterest
 ,coalesce(j.[23430.1], 0) as interestOnLoanedFunds
 ,coalesce(tax.qtdPremiumTax, 0) premiumTax
 ,me.meRefund
 ,coalesce(pyp.firstYear, 0) as PYPFirstYear
 ,coalesce(pyp.renewal, 0) as PYPRenewal



from dbo.vF00BaseCoverageQuarterly f00
  left outer join dbo.F02MonthiversaryRecord f02 
    on f02.carrierId = f00.carrierId and f02.policyNumber = f00.policyNumber
      and f02.activityYear = f00.activityYear and f02.activityMonth = f00.activityMonth
  left outer join dbo.vF02MonthiversaryRecordQuarterly pf02 
    on pf02.carrierId = f00.carrierId and pf02.policyNumber = f00.policyNumber
      and pf02.activityYear = @priorYear and pf02.activityQuarter = @priorQuarter
  left outer join dbo.NonModco nm on nm.policyNumber = f00.policyNumber
  left outer join vMtdJournalByCpaAccountNumberQuarterlyFinalDeaths j
    on j.carrierId = f00.carrierId and j.policyNumber = f00.policyNumber and j.maxJournalDate between @startDate and @endDate
  /*  
  left outer join dbo.MtdJournalHistory(@carrierId, dateadd(ms,-3,@startDate)) jh
    on jh.carrierId = f00.carrierId and jh.policyNumber = f00.policyNumber
  */
  left outer join dbo.PolicyYearPremiumsByQuarter(@carrierId, @year, @quarter) pyp on pyp.carrierId = f00.carrierId and pyp.policyNumber = f00.policyNumber
  inner join fees on fees.carrierId = f00.carrierId and @endDate between fees.startDate and fees.endDate
  inner join expenses on expenses.carrierId = f00.carrierId and @endDate between expenses.startDate and expenses.endDate
  left outer join dbo.UnearnedCoi uc ON uc.carrierId = f00.carrierId and uc.policyNumber = f00.policyNumber
      and uc.year = @year and uc.quarter = @quarter
  left outer join dbo.UnearnedCoi puc ON puc.carrierId = f00.carrierId and puc.policyNumber = f00.policyNumber
      and puc.year = @priorYear and puc.quarter = @priorQuarter
  left join dbo.vQtdPremiumTax tax on tax.[year] = @year and tax.[quarter] = @quarter and tax.policyNumber = f00.policyNumber
  left outer join dbo.vF15VariableAppreciationQuarterly va on va.carrierId = f00.carrierId and va.policyNumber = f00.policyNumber
      and va.activityYear = f00.activityYear and va.activityQuarter = f00.activityQuarter

--******************************************************************************
--6/10/2008 MRJ: Added inner join on dbo.F15MERefundQuarterly	
--******************************************************************************	
left outer join dbo.vF15MERefundQuarterly me on me.carrierId = f00.carrierId and me.policyNumber = f00.policyNumber
      and me.activityYear = f00.activityYear and me.activityQuarter = f00.activityQuarter
--******************************************************************************
--5/30/2008 MRJ: Added inner join on dbo.EarnedRate	
--******************************************************************************	
	inner join EarnedRate on EarnedRate.carrierId = f00.carrierId and @endDate between EarnedRate.startDate and EarnedRate.endDate

		
where f00.carrierId = @carrierId
  and f00.activityYear = @year
  and f00.activityQuarter = @quarter
order by carrierId, f00.policyNumber