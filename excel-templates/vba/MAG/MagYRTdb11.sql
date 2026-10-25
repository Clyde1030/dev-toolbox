--declare @year int; set @year = 2022
--declare @quarter int; set @quarter = 4
--declare @CarrierID varchar(3); set @carrierID = 'SLD'

declare @month int
set @month = @quarter*3
declare @YearMonth varchar(6)
set @yearmonth = concat(@year, right(concat('0',@month),2))


declare @ID varchar(3)
set @ID = case
	when @CarrierID = 'SLD' then '1SL' 
	when @CarrierID = 'NYL' then '1NY' 
	when @CarrierID = 'PLA' or @CarrierID = 'PLN' then '1PM' 
	when @CarrierID = 'JHL' then '1JH' 
	when @CarrierID = 'PR1' then 'PRU' end

select @carrierid CarrierID
,@quarter [Quarter], ReinsurerCompany,format(sum(totalPremium),'N2') 'Total Premium'
	from vbilling

where cedingcompany = @ID
and dateBilled IN (@YearMonth,@YearMonth-1,@YearMonth-2)
and cedingcompany+reinsurancetreaty in (
'PRUSQMAGGIB','1JHSWMAGF+A','1JHSWMAGS+C','1JHINMAGS+C','1JHLNMAGS+B','1JHSWMAGS+B','1JHSWMAGS-C',
'1JHINMAGS-C','1JHSWMAGF-A','1JHINMAGF-A','1JHINMAGF+A','1JHSWMBGUNE','1JHSWMBGS-E','1JHSWMBGS+E',
'1JHSWMAGF+B','1JHSWMBGS+F','1JHLQMAGGIA','1JHRQMAGGIA','1JHLNMAGGIA','1JHRGMAGGIA','1JHSWMAGFUA',
'1JHRGMAGFUA','1JHLNMAGUNA','1JHRGMAGUNA','1JHLNMAGUNB','1JHSWMAGUNB','1JHTAMAGFUA','1JHLQMAGGIB',
'1JHSQMAGGIB','1JHLNMAGGIB','1JHSWMAGGIB','1JHSWMAGUNC','1JHINMAGUNC','1JHINMAGFUA','1JHSQMAGGIC',
'1JHIQMAGGIC','1JHSWMAGGIC','1JHINMAGGIC','1JHSQMBGGIE','1JHSWMBGGIE','1JHSQMAGGIE','1JHCQMAGGIA',
'1JHRQMAGGIC','1JHSWMAGGIE','1JHCTMAGGIA','1JHRGMAGGIC','1JHSWMBGUNF','1JHSWMAGFUB','1NYINMAGGIC',
'1NYINMAGUNC','1NYIQMAGGIC','1NYMQMAGGKC','1NYMUMAGGKC','1NYMUMAGGKD','1NYSQMAGGIC','1NYSQMAGGKC',
'1NYSWMAGGIC','1NYSWMAGGKC','1NYSWMAGGKD','1NYSWMAGUNC','1PMSWMAGUNE','1PMRGMAGUNC','1PMSWMAGUNC',
'1PMINMAGUNC','1PMSWMAGS+D','1PMRGMAGS+B','1SLSLMAGFUA','1SLSLMAGF+A','PRUSWMAGGIB','PRUSWMAGFUA',
'PRUINMAGFUA','PRUSWMAGS+C','PRUINMAGS+C','PRUSWMAGS-E','PRURGMAGS-C','PRUSWMAGS+E','PRURGMAGS+C',
'PRURGMAGFUB','PRUSWMAGFUB','PRUSWMAGUNE','PRURGMAGUNC','1SLSLMAGF+B','1SLSLMAGF-A','1SLSLMAGF-B',
'1SLSLMAGFUB','1SLSLMAGTPA','1SLSLFACMAG','1SLSLFACTPA','1SLSLFACTPB')

and (
	(@CarrierID = 'PLA' and policyNumber like 'VM7%') 
	or (@CarrierID = 'PLN' and policyNumber like 'VM6%') 
	or (@CarrierID not in ('PLA','PLN'))
)

group by reinsurercompany
order by reinsurercompany

