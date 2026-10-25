--declare @year varchar(4); set @year = 2020
--declare @quarter int; set @quarter = 2
--declare @carrierId varchar(3); set @carrierId = 'SLF'

use Magnastar

--Admin Allowance
select 
f00.CarrierID as Carrier,
f00.PolicyNumber as Policy,
f00.ActivityYear as [Activity Year],
f00.ActivityQuarter as [Activity Quarter] ,
f00.IssueYear as [Issue Year],
f00.IssueMonth as [Issue Month],
f00.IssueDay as [Issue Day],
f00.CurrentStatus as [Current Status],
f02.FundValue [FundValue],
aa.AdminAllowanceRate/4/100 as [Admin Allowance Quarterly Rate],
f02.fundvalue*aa.AdminAllowanceRate/4/100 as [Admin Allowance]


from PPVA.vF00BaseCoverageQuarterly f00

  left outer join PPVA.F02MonthiversaryRecord f02 
    on f02.carrierId = f00.carrierId and f02.policyNumber = f00.policyNumber
      and f02.activityYear = f00.activityYear and f02.activityMonth = f00.activityMonth
		
	left outer join PPVA.AdminAllowance aa
			on f00.policynumber=aa.policyNumber
		
where PolicyDuration=datepart(YYYY,dateadd(mm,1,convert(datetime,f00.activityyear+f00.activitymonth+'01'))-convert(datetime,f00.issueyear+f00.issuemonth+f00.issueday))-1899
and f02.fundvalue>=AVBandLow and f02.fundvalue<AVBandHigh
and f00.activityyear=@year and f00.activityquarter=@quarter 

order by f00.policynumber


--Breakage
select  
mtd.carrierid,
mtd.policynumber,
sum(case when mtd.cpaAccountNumber like '35000.1%' 
		then amount 
		else 0 
		end) as Breakage,

sum(case when mtd.cpaAccountNumber like '55000.1%' 
		then amount 
		else 0 
		end) as BreakageExp,

sum(case when mtd.cpaAccountNumber like '35000.1%' then amount else 0 end)
+sum(case when mtd.cpaAccountNumber like '55000.1%'	then amount else 0 	end) as NetBreakage

from PPVA.MtdJournal mtd
left outer join dbo.[McCamish COA] coa
on mtd.carrierid=coa.CarrID and mtd.cpaAccountNumber=coa.AccountNumber
where 
((case when mtd.cpaAccountNumber like '35000.1%' 
		then amount 
		else 0 
		end) <> 0
or

(case when mtd.cpaAccountNumber like'55000.1%' 
		then amount 
		else 0 
		end)<> 0)

and datepart(YYYY,journaldate)=@year and datepart(qq, journaldate)=@quarter
--and carrierID=@carrierId

group by carrierId, policyNumber, journalDate
