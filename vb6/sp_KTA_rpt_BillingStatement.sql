--sp_KTA_rpt_BillingStatement 'N00042032', 'L01', '0221G', '0221G', '', '', '201802', '1'

CREATE PROCEDURE [dbo].[sp_KTA_rpt_BillingStatement]
(
	@CompanyID VARCHAR(10),
	@LocationID VARCHAR(3),
	@UnitNoFrom VARCHAR(10),
	@UnitNoThru VARCHAR(10),
	@GroupIDFrom VARCHAR(10),
	@GroupIDThru VARCHAR(10),
	@Period VARCHAR(6),
	--@UserID VARCHAR(10), 
	@ByWhat VARCHAR(1)	-- 1-> by unit, 2-> by group
)
AS
-- sp_KTA_rpt_BillingStatement 'N00042032','L01','0233F','0233F','','','201610',1
/*DECLARE @CompanyID VARCHAR(10),
		@LocationID VARCHAR(3),
		@UnitNoFrom VARCHAR(10),
		@UnitNoThru VARCHAR(10),
		@GroupIDFrom VARCHAR(10),
		@GroupIDThru VARCHAR(10),
		@Period VARCHAR(6),
		@UserID VARCHAR(10), 
		@ByWhat VARCHAR(1)	-- 1-> by unit, 2-> by group

SET @CompanyID='N00042032'
SET @LocationID='L01'
SET @UnitNoFrom='0628D'
SET @UnitNoThru='0628D'
SET @GroupIDFrom=''
SET @GroupIDThru=''
SET @Period='202512'
SET @ByWhat='1'
SET @UserID='sa'
*/
SET NOCOUNT ON
DECLARE @V_PeriodCutOff VARCHAR(6)='202510' 


CREATE TABLE #temp_H_BS
(
	KeyHD INT IDENTITY(1,1) NOT NULL,
	CompanyID VARCHAR(10),
	LocationID VARCHAR(3),
	OwnerName VARCHAR(100),
	[Address] VARCHAR(255),
	Address1 VARCHAR(100),
	Address2 VARCHAR(100),
	Address3 VARCHAR(100),
	Address4 VARCHAR(100),
	ContactPerson VARCHAR(100),
	UnitNo VARCHAR(5),
	Tower VARCHAR(2),
	[Floor] VARCHAR(2),
	[Number] VARCHAR(2),
	BillingNo VARCHAR(10),
	BillDate SMALLDATETIME,
	DueDate SMALLDATETIME,
	Previous MONEY,
	Payment MONEY,
	NewBilling MONEY,
	NewBalance MONEY,
	CurrencyID VARCHAR(3),
	BillingTo VARCHAR(1)
) 

CREATE TABLE #temp_D_BS
(
	KeyHD INT,
	SeqNo INT,
	[Description] VARCHAR(100),
	Amount MONEY,
	CR VARCHAR(1)
) 

CREATE TABLE #temp_D_BS_ELC
(
	KeyHD INT,
	MeterNo VARCHAR(20),
	RateGroup VARCHAR(10),
	MeterFactor MONEY,
	Wattage MONEY,
	Beginning MONEY,
	Ending MONEY,
	UsageFee1 MONEY,
	UsageFee2 MONEY,
	FeeUsage1 MONEY,
	FeeUsage2 MONEY,
	BPJU MONEY,
	FeeBPJU MONEY,
	TotalAmount MONEY
)

CREATE TABLE #temp_D_BS_WTR
(
	KeyHD INT,
	MeterNo VARCHAR(20),
	RateGroup VARCHAR(10),
	MeterFactor MONEY,
	PipeSize MONEY,
	Beginning MONEY,
	Ending MONEY,
	MeterUsage1 MONEY, 
	MeterUsage2 MONEY,
	MeterUsage3 MONEY,
	UsageFee1 MONEY,
	UsageFee2 MONEY,
	UsageFee3 MONEY,
	FeeMeter MONEY,
	FeeUsage MONEY,
	TotalAmount MONEY
)

CREATE TABLE #temp_D_BS_Phone
(
	KeyHD INT,
	Nomer VARCHAR(20),
	BiayaBulanan MONEY,
	BiayaLokal MONEY,
	BiayaSLJJ MONEY,
	BiayaSLI MONEY,
	BiayaAirtime MONEY,
	BiayaJapati MONEY,
	PulsaLokal INT,
	PulsaSLJJ INT,
	BiayaAdmin MONEY,
	PPN MONEY,
	TotalAmount MONEY
)
DECLARE @PPN MONEY
SET @PPN=CASE WHEN @Period<'202206' THEN 1.1 ELSE 1.11 END 

INSERT #temp_H_BS
select h.companyid,h.locationid,ownername,billingaddress,billingaddress1,billingaddress2,billingaddress3,billingaddress4,
contactperson,h.unitno,tower,[floor],number,h.billingno,h.BillDate,h.DueDate,h.Previous,h.Payment,h.NewBilling,h.NewBalance,
h.CurrencyID,billingto 
from kta_t_h_billing h  
INNER JOIN KTA_M_Unit kmu ON kmu.CompanyID=h.CompanyID AND kmu.LocationID=h.LocationID AND kmu.UnitNo=h.UnitNo AND kmu.SendEmail='y'
inner join kta_m_owner y on kmu.companyid=y.companyid and kmu.locationid=y.locationid and kmu.ownerid=y.ownerid 
inner join kta_m_roomtype z on z.companyid=kmu.companyid and z.locationid=kmu.locationid and z.roomtype=kmu.roomtype 
where h.companyid=@CompanyID and h.locationid=@LocationID
AND ((@ByWhat=1 AND h.UnitNo BETWEEN CASE WHEN @UnitNoFrom='' THEN h.UnitNo ELSE @UnitNoFrom END AND CASE WHEN @UnitNoThru='' THEN h.UnitNo ELSE @UnitNoThru END) OR  
	 (@ByWhat=2 AND kmu.GroupID BETWEEN CASE WHEN @GroupIDFrom='' THEN kmu.GroupID ELSE @GroupIDFrom END AND CASE WHEN @GroupIDThru='' THEN kmu.GroupID ELSE @GroupIDThru END) OR 
	  @ByWhat='')
and convert(varchar(6),h.billdate,112)=@Period
order by h.billingno

INSERT #temp_D_BS
SELECT a.KeyHD, 1, 'Saldo Terdahulu / Previous Balance', a.Previous, CASE WHEN a.Previous<0 THEN 'Y' ELSE 'N' END
FROM #temp_H_BS a

INSERT #temp_D_BS
select keyhd1, 
(SELECT ISNULL(MAX(seqno),0)+1 FROM #temp_D_BS WHERE KeyHD=keyhd1)+ROW_NUMBER() OVER (PARTITION BY keyhd1 ORDER BY keyhd1,d.RefNo  )-1,
CASE WHEN d.RecordType = 4 THEN 'Pembayaran / Payment Thank You '
	 WHEN d.RecordType = 3 THEN 'Credit Note '+x.[Description]+' '
	 ELSE 'Refund ' END + d.RefNo+' '+CAST(DAY(d.RKDate) AS VARCHAR)+' '+DATENAME(MONTH,d.RKDate)+' '+CAST(YEAR(d.RKDate) AS VARCHAR), SUM(d.Amount), 'Y'
FROM
( 
	SELECT a.*, 
	StartDate=cast(convert(varchar(6),DATEADD(MONTH,-1, z2.ActivePeriod+1)-1,112)+    
     RIGHT('00'+cast(CASE WHEN convert(varchar(6),z2.ActivePeriod,112)<=convert(varchar(6),z.StartingNewBilling,112) then z.PreviousBillDate else z.BillDate END     
     AS VARCHAR(2)),2) AS SMALLDATETIME),
    EndDate=cast(CONVERT(VARCHAR(6),z2.ActivePeriod,112)+ RIGHT('00'+cast(z.BillDate AS VARCHAR(2)),2) AS SMALLDATETIME)-1,
    PrevPeriod=DATEADD(MONTH,-1, z2.ActivePeriod+1)-1,
    x1.keyhd keyhd1 
	FROM kta_t_h_outstanding a 
	INNER JOIN M_Active_Period z2 ON z2.AutoNum='KCAR'+a.CompanyID+a.LocationID
	INNER JOIN KTA_M_Penalty z ON z.CompanyID=a.CompanyID AND z.LocationID=a.LocationID
	INNER JOIN #temp_H_BS x1 ON x1.UnitNo=a.UnitNo AND x1.CompanyID=a.CompanyID AND x1.LocationID=a.LocationID  
	WHERE txtype in ('B','W')  
) h 
inner JOIN
--(
--	SELECT *, RecordType1=CASE WHEN RecordType IN (2,4) THEN 4 WHEN RecordType = 3 THEN 3 ELSE 0 END FROM KTA_T_D_Outstanding 
--) d 
KTA_T_D_Outstanding d
on d.keyhd=h.keyhd 
inner join kta_m_income x on x.companyid=h.companyid and x.locationid=h.locationid and x.incomecode=h.incomecode and recordtype in (3,4,5)
where refdate <= enddate and ((txtype='w' and convert(varchar(6),documentdate,112)=convert(varchar(6),prevperiod,112)
and convert(varchar(6),refdate,112)>=convert(varchar(6),prevperiod,112)) or (txtype='w' and convert(varchar(6),documentdate,112)<convert(varchar(6),prevperiod,112) and refdate>=startdate
) or (txtype='b' and refdate>=startdate))
AND d.RefNo<>'-'
--group BY KeyHD1, d.RefNo, d.RKDate, d.RecordType1, x.[Description]
group BY KeyHD1, CASE WHEN d.RecordType = 4 THEN 'Pembayaran / Payment Thank You '
	 WHEN d.RecordType = 3 THEN 'Credit Note '+x.[Description]+' '
	 ELSE 'Refund ' END , d.RefNo,d.RKDate

INSERT #temp_D_BS
SELECT a.KeyHD, (SELECT ISNULL(MAX(seqno),0)+1 FROM #temp_D_BS WHERE KeyHD=a.keyhd)+ROW_NUMBER() OVER (PARTITION BY a.keyhd ORDER BY a.keyhd,MAX(a.SeqNo)  )-1,
CASE WHEN a.[Description]<>'' THEN a.[Description]
ELSE  
CASE WHEN a.IncomeCode='TEC' THEN 'Telephone Reimbursement Phone No. '+a.PhoneNo
	 WHEN a.IncomeCode = 'WOM' THEN 'Work Order Ref.Invoice No. '+a.InvoiceNo+' '+CAST(DAY(a.InvoiceDate) AS VARCHAR)+' '+DATENAME(MONTH,a.InvoiceDate)+' '+CAST(YEAR(a.InvoiceDate) AS VARCHAR)		--CONVERT(VARCHAR(20),a.InvoiceDate,106)
	 WHEN a.IncomeCode='PEC' THEN 'Late Payment Charges '+ DATENAME(MONTH,a.BillDate-DAY(a.BillDate))+' '+CAST(YEAR(a.BillDate-DAY(a.BillDate)) AS VARCHAR) --RIGHT(CONVERT(VARCHAR(20), a.BillDate-DAY(a.BillDate), 106), 8) 
	 WHEN a.IncomeCode = 'MFE' THEN 'Maintenance Fee '+ DATENAME(MONTH,a.billdate)+' '+CAST(YEAR(a.billdate) AS VARCHAR)  + --SUBSTRING(CONVERT(VARCHAR(11), a.BillDate, 113), 4, 8) + 
	 ' ('+ REPLACE(CONVERT(VARCHAR,kmrt.Area,1), '.00','') + ' Sqm x ' + CASE WHEN a.CurrencyID='IDR' THEN 'Rp. '+ REPLACE(CONVERT(VARCHAR,kmrt.MaintenanceFee+kmrt.SinkingFund,1), '.00','') ELSE 'US$ '+CONVERT(VARCHAR,kmrt.MaintenanceFee+kmrt.SinkingFund,1) END+')'     
	 ELSE d.[Description] 
	 END 
END, 
SUM(a.Amount),'N'
  
FROM
(
	SELECT c.KeyHD,a.SeqNo,  
	IncomeCode=CASE WHEN a.IncomeCode IN ('WOL','WOM','WOH') THEN 'WOM'
					 WHEN a.IncomeCode IN ('SKF','MFE') THEN 'MFE'
					 ELSE a.IncomeCode END,
	a.Amount, a.InvoiceNo, a.InvoiceDate, a.PhoneNo, a.[Description], b.CompanyID, b.LocationID, b.UnitNo, b.BillDate, b.CurrencyID
	FROM KTA_T_D_Billing a
	INNER JOIN KTA_T_H_Billing b ON a.KeyHD=b.KeyHD
	INNER JOIN #temp_H_BS c ON c.CompanyID=b.CompanyID AND c.LocationID=b.LocationID AND c.UnitNo=b.UnitNo AND c.BillingNo=b.BillingNo
) a 
inner join kta_m_income d on d.incomecode=a.incomecode and d.companyid=a.companyid and d.locationid=a.locationid
INNER JOIN KTA_M_Unit kmu ON kmu.CompanyID=a.CompanyID AND kmu.LocationID=a.LocationID AND kmu.UnitNo=a.UnitNo
INNER JOIN KTA_M_RoomType kmrt ON kmrt.CompanyID=a.CompanyID AND kmrt.LocationID=a.LocationID AND kmrt.RoomType=kmu.RoomType
GROUP BY a.KeyHD, a.IncomeCode, a.PhoneNo, a.InvoiceNo, a.InvoiceDate, a.BillDate, kmrt.Area, kmrt.MaintenanceFee, kmrt.SinkingFund, a.CurrencyID, d.[Description], a.[Description] 
ORDER BY a.KeyHD

INSERT #temp_D_BS
SELECT a.KeyHD, (SELECT ISNULL(MAX(seqno),0)+1 FROM #temp_D_BS WHERE KeyHD=a.keyhd), 'Saldo Baru / New Balance', a.NewBalance, CASE WHEN a.NewBalance<0 THEN 'Y' ELSE 'N' END
FROM #temp_H_BS a

INSERT #temp_D_BS_Phone
SELECT c.KeyHD, a.PhoneNo, 
Fee=((a.Fee*10/9)*@PPN)+a.Other, [Local]=(a.[Local]*10/9)*@PPN, SLJJ=(a.SLJJ*10/9)*@PPN, International=(a.International*10/9)*@PPN, 
Airtime=(a.Airtime*10/9)*@PPN, Japati=(a.Japati*10/9)*@PPN, 
--a.Fee, a.[Local], a.SLJJ, a.International, 
--a.Airtime, a.Japati, 
a.Pulsa_Local, a.Pulsa_SLJJ, a.BeaAdmin, a.PPN, a.Total    
FROM KTA_T_D_Billing_Phone a INNER JOIN KTA_T_H_Billing b ON b.KeyHD=a.KeyHD
INNER JOIN #temp_H_BS c ON c.CompanyID=b.CompanyID AND c.LocationID=b.LocationID AND c.BillingNo=b.BillingNo

INSERT #temp_D_BS_ELC
SELECT KeyHD, MeterNo, RateGroup, MeterFactor, Wattage, Beginning1, Ending1, UsageFee1, UsageFee2,
FeeUsage1, FeeUsage2, BPJU, FeeBPJU, TotalAmount 
FROM
(
SELECT c.KeyHD, a.MeterNo, a.RateGroup, a.MeterFactor, a.Wattage, a.Beginning1, a.Ending1, b.UsageFee1, b.UsageFee2,
a.FeeUsage1, a.FeeUsage2, b.BPJU, a.FeeBPJU, a.TotalAmount, seqno=ROW_NUMBER() OVER ( PARTITION BY a.MeterNo ORDER BY b.StartPeriod DESC )
from utl_t_elc_consumption a INNER JOIN utl_m_elc_rate b ON a.RateGroup=b.RateGroup AND b.isActive=1
INNER JOIN #temp_H_BS c ON c.CompanyID=a.CompanyID AND c.LocationID=a.LocationID AND c.UnitNo=a.TenantID 
where a.period=CONVERT(VARCHAR(20),c.BillDate-DAY(c.BillDate),106) AND a.Period>=b.StartPeriod
) a WHERE a.seqno=1

INSERT #temp_D_BS_WTR
SELECT KeyHD, MeterNo, RateGroup, MeterFactor, PipeSize, Beginning, Ending, 
CASE WHEN @Period<@V_PeriodCutOff THEN a.MeterUsage1+a.MeterUsage2+a.MeterUsage3 ELSE a.MeterUsage1 END, 
CASE WHEN @Period<@V_PeriodCutOff THEN 0 ELSE a.MeterUsage2 END, 
CASE WHEN @Period<@V_PeriodCutOff THEN 0 ELSE a.MeterUsage3 END, 
usagefee1, 
CASE WHEN @Period<@V_PeriodCutOff THEN 0 ELSE a.UsageFee2 END, 
CASE WHEN @Period<@V_PeriodCutOff THEN 0 ELSE a.UsageFee3 END, 
biaya,
FeeUsage, TotalAmount
FROM
(
SELECT c.KeyHD, a.MeterNo, b.RateGroup, b.MeterFactor, b.PipeSize, a.Beginning, a.Ending, a.MeterUsage1, a.MeterUsage2, a.MeterUsage3, h.usagefee1, h.UsageFee2, h.UsageFee3, a.FeeMeter+a.FeeMaintenance biaya,
a.FeeUsage, a.TotalAmount, seqno=ROW_NUMBER() OVER ( PARTITION BY a.MeterNo ORDER BY h.StartPeriod DESC ) 
from utl_t_wtr_consumption a INNER JOIN utl_m_wtr_meter b ON a.CompanyID=b.CompanyID AND a.LocationID=b.LocationID AND a.MeterNo=b.MeterNo AND b.Terminated='n'
INNER JOIN (utl_m_wtr_h_rate h INNER JOIN utl_m_wtr_d_rate d ON d.KeyHD=h.KeyHD) ON h.rategroup=b.RateGroup AND b.PipeSize=d.pipesize
INNER JOIN #temp_H_BS c ON c.CompanyID=a.CompanyID AND c.LocationID=a.LocationID AND c.UnitNo=b.TenantID 
where a.Period=CONVERT(VARCHAR(20),c.BillDate-DAY(c.BillDate),106) AND a.Period>=h.startperiod
) a WHERE a.seqno=1
--------------------------------------------------------------------------------------------------------------
-- Main report
--------------------------------------------------------------------------------------------------------------
--; with rcte
--as
--(
--	select	keyhd,groupid, countrow, startrow=17-countrow
--	from	
--	(
--		SELECT keyhd,groupid,17-MAX(seqno) countrow
--		FROM
--		(
--			SELECT keyhd,groupid, seqno=ROW_NUMBER() OVER ( PARTITION BY keyhd,groupid ORDER BY seqno)
--			FROM
--			(
--				SELECT a.KeyHD, GroupID=DENSE_RANK() OVER (PARTITION BY a.keyhd ORDER BY b.SeqNo/18), b.[Description], b.Amount, b.seqno
--				FROM #temp_H_BS a INNER JOIN #temp_D_BS b ON a.KeyHD=b.KeyHD
--			) a
--		) a
--		GROUP BY keyhd,a.groupid
--		HAVING 17-MAX(seqno)>0
--	) a
--	union all
--	select	keyhd,groupid,  countrow = countrow - 1, startrow
--	from	rcte
--	where	countrow > 1
--)
--select	a.keyhd,a.groupid, SubGroupID=1, seqno=startrow + ROW_NUMBER() OVER (PARTITION BY a.keyhd ORDER BY a.groupid)
--, '' OwnerName, '' Address1, '' Address2, '' Address3, '' Address4, b.UnitNo, b.Tower, b.[Floor], '' Number,
--b.BillingNo, b.BillDate, b.DueDate, b.Previous, b.Payment, b.NewBilling, b.NewBalance, b.CurrencyID, b.BillingTo,
--'' [Description], 0 Amount, BillingDesc='', 
--[desc11]='', [desc12]='', [desc13]='', [desc14]='',
--[desc21]='', [desc22]='', [desc23]='', [desc24]='',
--'' CR, '' Nomer, 0 BiayaBulanan, 0 BiayaLokal, 0 BiayaSLJJ, 0 BiayaSLI,
--0 BiayaAirtime, 0 BiayaJapati, 0 PulsaLokal, 0 PulsaSLJJ, 0 BiayaAdmin, 0 PPN,
--'' MeterNo, '' RateGroup, 0 MeterFactor, 0 Wattage, 0 PipeSize, 0 Beginning, 0 Ending, 0 UsageFee, 0 UsageFee1,
--0 UsageFee2, 0 FeeMeter, 0 FeeUsage, 0 FeeUsage1, 0 FeeUsage2, 0 BPJU, 0 FeeBPJU,0 TotalAmount
--INTO #temp
--from rcte a, #temp_H_BS b WHERE a.KeyHD=b.KeyHD


	SELECT
	--(CASE 
	--	WHEN a.GroupID = 10 THEN 0
	--	ELSE
	--		(CASE 
	--			WHEN row_number() OVER (PARTITION BY a.companyid, a.locationid, a.KeyHD, a.GroupID ORDER BY a.companyid, a.locationid, a.KeyHD, a.GroupID, a.SeqNo DESC) <> 1 THEN 0
	--			ELSE row_number() OVER (PARTITION BY a.companyid, a.locationid, a.KeyHD, a.GroupID ORDER BY a.companyid, a.locationid, a.KeyHD, a.GroupID, a.SeqNo DESC)				
	--		END) 
	--END) AS IsBill,
	a.*,
	[desc11]=d1.BankDes1, [desc12]=d1.BankDes2, [desc13]=d1.BankDes3, [desc14]=d1.BankDes4,
	[desc21]=d2.BankDes1, [desc22]=d2.BankDes2, [desc23]=d2.BankDes3, [desc24]=d2.BankDes4
	, SeqNo1=CASE WHEN subgroupid=1 THEN 0 ELSE ROW_NUMBER() OVER (ORDER BY subgroupid) END 
	--INTO #temp_Report
	FROM
	(
	SELECT a.companyid, a.locationid, a.KeyHD, GroupID=DENSE_RANK() OVER (PARTITION BY a.keyhd ORDER BY b.SeqNo/18), SubGroupID=1, b.SeqNo, a.OwnerName, a.Address1, a.Address2, a.Address3, a.Address4, a.UnitNo, a.Tower, a.[Floor], a.Number,
	a.BillingNo, a.BillDate, a.DueDate, a.Previous, a.Payment, a.NewBilling, a.NewBalance, a.CurrencyID, a.BillingTo,
	b.[Description], b.Amount, BillingDesc=c.[Description], 
	b.CR, '' Nomer, 0 BiayaBulanan, 0 BiayaLokal, 0 BiayaSLJJ, 0 BiayaSLI,
	0 BiayaAirtime, 0 BiayaJapati, 0 PulsaLokal, 0 PulsaSLJJ, 0 BiayaAdmin, 0 PPN,
	'' MeterNo, '' RateGroup, 0 MeterFactor, 0 Wattage, 0 PipeSize, 0 Beginning, 0 Ending, 0 UsageFee1,
	0 UsageFee2, 0 FeeMeter, 0 FeeUsage, 0 FeeUsage1, 0 FeeUsage2, 0 BPJU, 0 FeeBPJU, 
	0 TotalAmount, 0 MeterUsage1, 0 MeterUsage2, 0 MeterUsage3, 0 UsageFee3
	FROM #temp_H_BS a INNER JOIN #temp_D_BS b ON a.KeyHD=b.KeyHD
	INNER JOIN KTA_M_BillDes c ON c.CompanyID=a.CompanyID AND c.LocationID=a.LocationID
	UNION ALL 
	SELECT a.companyid, a.locationid, a.KeyHD, GroupID=10, SubGroupID=2, SeqNo=ROW_NUMBER() OVER (PARTITION BY a.keyhd ORDER BY b.nomer) , a.OwnerName, '' Address1, '' Address2, '' Address3, '' Address4, a.UnitNo, '' Tower, '' [Floor], '' Number,
	a.BillingNo, a.BillDate, a.DueDate, a.Previous, a.Payment, a.NewBilling, a.NewBalance, a.CurrencyID, a.BillingTo,
	'' [Description], 0 Amount, '' BillingDesc, 
	'' CR, b.Nomer, b.BiayaBulanan, b.BiayaLokal, b.BiayaSLJJ, b.BiayaSLI,
	b.BiayaAirtime, b.BiayaJapati, b.PulsaLokal, b.PulsaSLJJ, b.BiayaAdmin, b.PPN,
	'' MeterNo, '' RateGroup, 0 MeterFactor, 0 Wattage, 0 PipeSize, 0 Beginning, 0 Ending, 0 UsageFee1,
	0 UsageFee2, 0 FeeMeter, 0 FeeUsage, 0 FeeUsage1, 0 FeeUsage2, 0 BPJU, 0 FeeBPJU, 
	b.TotalAmount, 0 MeterUsage1, 0 MeterUsage2, 0 MeterUsage3, 0 UsageFee3
	FROM #temp_H_BS a INNER JOIN #temp_D_BS_Phone b ON a.KeyHD=b.KeyHD
	UNION ALL
	SELECT a.companyid, a.locationid, a.KeyHD, GroupID=10, SubGroupID=3, SeqNo=ROW_NUMBER() OVER (PARTITION BY a.keyhd ORDER BY b.meterno), a.OwnerName, '' Address1, '' Address2, '' Address3, '' Address4, a.UnitNo, '' Tower, '' [Floor], '' Number,
	a.BillingNo, a.BillDate, a.DueDate, a.Previous, a.Payment, a.NewBilling, a.NewBalance, a.CurrencyID, a.BillingTo,
	'' [Description], 0 Amount, '' BillingDesc, 
	'' CR, '' Nomer, 0 BiayaBulanan, 0 BiayaLokal, 0 BiayaSLJJ, 0 BiayaSLI,
	0 BiayaAirtime, 0 BiayaJapati, 0 PulsaLokal, 0 PulsaSLJJ, 0 BiayaAdmin, 0 PPN,
	b.MeterNo, b.RateGroup, b.MeterFactor, b.Wattage, 0 PipeSize, b.Beginning, b.Ending, b.UsageFee1,
	b.UsageFee2, 0 FeeMeter, 0 FeeUsage, b.FeeUsage1, b.FeeUsage2, b.BPJU, b.FeeBPJU, 
	b.TotalAmount, 0 MeterUsage1, 0 MeterUsage2, 0 MeterUsage3, 0 UsageFee3
	FROM #temp_H_BS a INNER JOIN #temp_D_BS_ELC b ON a.KeyHD=b.KeyHD
	UNION ALL
	SELECT a.companyid, a.locationid, a.KeyHD, GroupID=10, SubGroupID=4, SeqNo=ROW_NUMBER() OVER (PARTITION BY a.keyhd ORDER BY b.meterno), a.OwnerName, '' Address1, '' Address2, '' Address3, '' Address4, a.UnitNo, '' Tower, '' [Floor], '' Number,
	a.BillingNo, a.BillDate, a.DueDate, a.Previous, a.Payment, a.NewBilling, a.NewBalance, a.CurrencyID, a.BillingTo,
	'' [Description], 0 Amount, '' BillingDesc, 
	'' CR, '' Nomer, 0 BiayaBulanan, 0 BiayaLokal, 0 BiayaSLJJ, 0 BiayaSLI,
	0 BiayaAirtime, 0 BiayaJapati, 0 PulsaLokal, 0 PulsaSLJJ, 0 BiayaAdmin, 0 PPN,
	b.MeterNo, b.RateGroup, b.MeterFactor, 0 Wattage, b.PipeSize, b.Beginning, b.Ending, b.UsageFee1,
	b.UsageFee2, b.FeeMeter, b.FeeUsage, 0 FeeUsage1, 0 FeeUsage2, 0 BPJU, 0 FeeBPJU, 
	b.TotalAmount, b.MeterUsage1, b.MeterUsage2, b.MeterUsage3, b.UsageFee3
	FROM #temp_H_BS a INNER JOIN #temp_D_BS_WTR b ON a.KeyHD=b.KeyHD
	--UNION ALL
	--SELECT * FROM #temp
	) a
	INNER JOIN (SELECT TOP 1 * FROM KTA_M_BillBank ORDER BY SeqNo) d1 ON d1.CompanyID=a.CompanyID AND d1.LocationID=a.LocationID 
	INNER JOIN (SELECT TOP 1 * FROM KTA_M_BillBank ORDER BY SeqNo DESC) d2 ON d2.CompanyID=a.CompanyID AND d2.LocationID=a.LocationID

--ORDER BY keyhd, subgroupid, groupid, case when (subgroupid=1 and amount>0) OR subgroupid<>1 then a.seqno ELSE 999 END
ORDER BY keyhd, subgroupid, groupid, case when subgroupid=1 OR (subgroupid<>1 and amount>0) then a.seqno ELSE 999 END  
--------------------------------------------------------------------------------------------------------------
DROP TABLE  #temp_H_BS
DROP TABLE  #temp_D_BS
DROP TABLE  #temp_D_BS_ELC
DROP TABLE  #temp_D_BS_WTR
DROP TABLE  #temp_D_BS_Phone
--DROP TABLE	#temp
