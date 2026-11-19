/*
 KTA Web Billing - TVP validation and preview layer.
 SQL Server 2008 compatible. No persistent staging or batch tables.
*/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

/* Remove objects from the previous physical-staging design, if it was deployed. */
IF OBJECT_ID('dbo.usp_KTA_Billing_CreateBatch','P') IS NOT NULL DROP PROCEDURE dbo.usp_KTA_Billing_CreateBatch
GO
IF OBJECT_ID('dbo.usp_KTA_Billing_ValidateBatch','P') IS NOT NULL DROP PROCEDURE dbo.usp_KTA_Billing_ValidateBatch
GO
IF OBJECT_ID('dbo.KTA_W_BillingMember','U') IS NOT NULL DROP TABLE dbo.KTA_W_BillingMember
GO
IF OBJECT_ID('dbo.KTA_W_BillingCard','U') IS NOT NULL DROP TABLE dbo.KTA_W_BillingCard
GO
IF OBJECT_ID('dbo.KTA_W_BillingMonthly','U') IS NOT NULL DROP TABLE dbo.KTA_W_BillingMonthly
GO
IF OBJECT_ID('dbo.KTA_W_BillingBatch','U') IS NOT NULL DROP TABLE dbo.KTA_W_BillingBatch
GO

IF OBJECT_ID('dbo.usp_KTA_Billing_Validate','P') IS NOT NULL DROP PROCEDURE dbo.usp_KTA_Billing_Validate
GO
CREATE PROCEDURE dbo.usp_KTA_Billing_Validate
 @CompanyID varchar(10), @LocationID varchar(3),
 @BillingMember dbo.Type_Tbl_C4 READONLY,
 @BillingCard dbo.Type_Tbl_C7 READONLY,
 @BillingMonthly dbo.Type_Tbl_C7 READONLY,
 @CardUnavailableConfirmed bit=0, @MonthlyUnavailableConfirmed bit=0
AS
BEGIN
 SET NOCOUNT ON;
 CREATE TABLE #BillingMember(RowNo int IDENTITY(1,1),UnitNo varchar(5),DateTrans smalldatetime,[Description] varchar(100),Amount money);
 INSERT #BillingMember(UnitNo,DateTrans,[Description],Amount) SELECT LTRIM(RTRIM(C1)),CONVERT(smalldatetime,C2,120),C3,CONVERT(money,C4) FROM @BillingMember;
 CREATE TABLE #BillingCard(RowNo int IDENTITY(1,1),ChargingCardName varchar(100),ChargingCardNo varchar(100),Currency varchar(3),TotalPowerCost money,TotalPowerConsumption decimal(18,4),PeriodFrom smalldatetime,PeriodThru smalldatetime);
 INSERT #BillingCard(ChargingCardName,ChargingCardNo,Currency,TotalPowerCost,TotalPowerConsumption,PeriodFrom,PeriodThru) SELECT LTRIM(RTRIM(C1)),C2,C3,CONVERT(money,C4),CASE WHEN ISNULL(C5,'')='' THEN NULL ELSE CONVERT(decimal(18,4),C5) END,CONVERT(smalldatetime,C6,120),CONVERT(smalldatetime,C7,120) FROM @BillingCard;
 CREATE TABLE #BillingMonthly(RowNo int IDENTITY(1,1),DriverAccount varchar(100),SubtotalAmount money,SubtotalPowerUsage decimal(18,4),PeriodFrom smalldatetime,PeriodThru smalldatetime);
 INSERT #BillingMonthly(DriverAccount,SubtotalAmount,SubtotalPowerUsage,PeriodFrom,PeriodThru) SELECT LTRIM(RTRIM(C1)),CONVERT(money,C2),CASE WHEN ISNULL(C3,'')='' THEN NULL ELSE CONVERT(decimal(18,4),C3) END,CONVERT(smalldatetime,C4,120),CONVERT(smalldatetime,C5,120) FROM @BillingMonthly;
 DECLARE @Period smalldatetime,@PeriodCAR smalldatetime,@MonthStart smalldatetime,@MonthEnd smalldatetime,
         @Message varchar(2000);
 SELECT @Period=ActivePeriod FROM dbo.M_Active_Period WHERE AutoNum='KAR'+@CompanyID+@LocationID;
 SELECT @PeriodCAR=ActivePeriod FROM dbo.M_Active_Period WHERE AutoNum='KCAR'+@CompanyID+@LocationID;
 SET @Message='';
 IF @Period IS NULL SET @Message=@Message+'Active KAR period was not found. ';
 IF @PeriodCAR IS NULL SET @Message=@Message+'Active KCAR period was not found. ';
 IF @Period IS NOT NULL
 BEGIN
  SET @MonthStart=DATEADD(month,DATEDIFF(month,0,@Period),0);
  SET @MonthEnd=DATEADD(month,1,@MonthStart);
 END
 IF @PeriodCAR IS NOT NULL AND @Period IS NOT NULL AND CONVERT(varchar(6),@PeriodCAR,112)<>CONVERT(varchar(6),@Period,112)
  SET @Message=@Message+'Active KAR and KCAR periods are not aligned for processing. ';
 IF NOT EXISTS(SELECT 1 FROM #BillingMember)
  SET @Message=@Message+'Billing_member file is required and must contain data. ';
 IF NOT EXISTS(SELECT 1 FROM #BillingCard) AND @CardUnavailableConfirmed=0
  SET @Message=@Message+'Upload Card Payment Statistics or confirm it is unavailable. ';
 IF NOT EXISTS(SELECT 1 FROM #BillingMonthly) AND @MonthlyUnavailableConfirmed=0
  SET @Message=@Message+'Upload Billing Report Monthly or confirm it is unavailable. ';
 IF @Period IS NOT NULL AND EXISTS(SELECT 1 FROM #BillingMember WHERE DateTrans<@MonthStart OR DateTrans>=@MonthEnd)
  SET @Message=@Message+'Billing_member contains transaction dates outside active KAR period. ';
 IF @Period IS NOT NULL AND EXISTS(SELECT 1 FROM #BillingCard WHERE PeriodFrom<@MonthStart OR PeriodThru>=@MonthEnd OR PeriodFrom>PeriodThru)
  SET @Message=@Message+'Card Payment Statistics contains invalid Statistical Period. ';
 IF @Period IS NOT NULL AND EXISTS(SELECT 1 FROM #BillingMonthly WHERE PeriodFrom<@MonthStart OR PeriodThru>=@MonthEnd OR PeriodFrom>PeriodThru)
  SET @Message=@Message+'Billing Report Monthly contains invalid From-Thru period. ';
 IF EXISTS(SELECT 1 FROM #BillingMember m WHERE NOT EXISTS(
    SELECT 1 FROM dbo.KTA_M_Unit u WHERE u.CompanyID=@CompanyID AND u.LocationID=@LocationID
     AND u.UnitNo=m.UnitNo AND u.Tower<>'' AND u.OwnerID<>''
     AND EXISTS(SELECT 1 FROM dbo.KTA_T_H_Billing h WHERE h.CompanyID=u.CompanyID AND h.LocationID=u.LocationID
       AND h.UnitNo=u.UnitNo AND CONVERT(varchar(6),h.BillDate,112)=CONVERT(varchar(6),DATEADD(month,-1,@Period),112))))
  SET @Message=@Message+'Billing_member contains units that are not eligible for this billing run. ';
 IF EXISTS(SELECT 1 FROM #BillingCard c WHERE NOT EXISTS(
    SELECT 1 FROM dbo.KTA_M_Unit u WHERE u.CompanyID=@CompanyID AND u.LocationID=@LocationID
     AND u.UnitNo=c.ChargingCardName AND u.Tower<>'' AND u.OwnerID<>''
     AND EXISTS(SELECT 1 FROM dbo.KTA_T_H_Billing h WHERE h.CompanyID=u.CompanyID AND h.LocationID=u.LocationID
       AND h.UnitNo=u.UnitNo AND CONVERT(varchar(6),h.BillDate,112)=CONVERT(varchar(6),DATEADD(month,-1,@Period),112))))
  SET @Message=@Message+'Card Payment Statistics contains Charging Card Name values that are not eligible unit numbers. ';
 IF EXISTS(SELECT 1 FROM #BillingMonthly m WHERE NOT EXISTS(
    SELECT 1 FROM dbo.KTA_M_Unit u WHERE u.CompanyID=@CompanyID AND u.LocationID=@LocationID
     AND u.UnitNo=m.DriverAccount AND u.Tower<>'' AND u.OwnerID<>''
     AND EXISTS(SELECT 1 FROM dbo.KTA_T_H_Billing h WHERE h.CompanyID=u.CompanyID AND h.LocationID=u.LocationID
       AND h.UnitNo=u.UnitNo AND CONVERT(varchar(6),h.BillDate,112)=CONVERT(varchar(6),DATEADD(month,-1,@Period),112))))
  SET @Message=@Message+'Billing Report Monthly contains Driver Account values that are not eligible unit numbers. ';
 IF NOT EXISTS(SELECT 1 FROM dbo.KTA_M_Penalty WHERE CompanyID=@CompanyID AND LocationID=@LocationID)
  SET @Message=@Message+'KTA_M_Penalty configuration was not found. ';
 IF EXISTS(SELECT 1 FROM dbo.KTA_M_Penalty WHERE CompanyID=@CompanyID AND LocationID=@LocationID AND (BillDate IS NULL OR BillDate=0))
  SET @Message=@Message+'KTA_M_Penalty.BillDate is not configured. ';
 IF NOT EXISTS(SELECT 1 FROM dbo.M_Rate WHERE @PeriodCAR IS NOT NULL AND CONVERT(varchar(6),[Date],112)=CONVERT(varchar(6),@PeriodCAR,112))
  SET @Message=@Message+'M_Rate for active CAR period was not found. ';
 IF NOT EXISTS(SELECT 1 FROM dbo.KTA_T_BillingPhone WHERE @Period IS NOT NULL AND CONVERT(varchar(8),BillDate,112)=
   CONVERT(varchar(8),DATEADD(day,-1,DATEADD(month,-1,DATEADD(day,1,@Period))),112))
  SET @Message=@Message+'Telephone billing source for the required prior period is missing. ';
 SELECT CASE WHEN @Message='' THEN 1 ELSE 2 END ValidationStatus,
        NULLIF(@Message,'') ValidationMessage,@Period ActivePeriod;
END
GO

IF OBJECT_ID('dbo.usp_KTA_Billing_Preview','P') IS NOT NULL DROP PROCEDURE dbo.usp_KTA_Billing_Preview
GO
CREATE PROCEDURE dbo.usp_KTA_Billing_Preview
 @CompanyID varchar(10), @LocationID varchar(3),
 @BillingMember dbo.Type_Tbl_C4 READONLY,
 @BillingCard dbo.Type_Tbl_C7 READONLY,
 @BillingMonthly dbo.Type_Tbl_C7 READONLY,
 @CardUnavailableConfirmed bit=0, @MonthlyUnavailableConfirmed bit=0
AS
BEGIN
 SET NOCOUNT ON;
 DECLARE @Status int,@Message varchar(2000),@Period smalldatetime;
 DECLARE @V TABLE(ValidationStatus int,ValidationMessage varchar(2000),ActivePeriod smalldatetime);
 INSERT @V EXEC dbo.usp_KTA_Billing_Validate @CompanyID,@LocationID,@BillingMember,@BillingCard,@BillingMonthly,@CardUnavailableConfirmed,@MonthlyUnavailableConfirmed;
 SELECT @Status=ValidationStatus,@Message=ValidationMessage,@Period=ActivePeriod FROM @V;
 IF @Status<>1
 BEGIN
  SELECT @Status ValidationStatus,@Message ValidationMessage,@Period ActivePeriod;
  RETURN;
 END
 SELECT 1 ValidationStatus,CAST(NULL AS varchar(2000)) ValidationMessage,@Period ActivePeriod;
 CREATE TABLE #BillingMember(RowNo int IDENTITY(1,1),UnitNo varchar(5),DateTrans smalldatetime,[Description] varchar(100),Amount money);
 INSERT #BillingMember(UnitNo,DateTrans,[Description],Amount) SELECT LTRIM(RTRIM(C1)),CONVERT(smalldatetime,C2,120),C3,CONVERT(money,C4) FROM #BillingMember;
 CREATE TABLE #BillingCard(RowNo int IDENTITY(1,1),ChargingCardName varchar(100),ChargingCardNo varchar(100),Currency varchar(3),TotalPowerCost money,TotalPowerConsumption decimal(18,4),PeriodFrom smalldatetime,PeriodThru smalldatetime);
 INSERT #BillingCard(ChargingCardName,ChargingCardNo,Currency,TotalPowerCost,TotalPowerConsumption,PeriodFrom,PeriodThru) SELECT LTRIM(RTRIM(C1)),C2,C3,CONVERT(money,C4),CASE WHEN ISNULL(C5,'')='' THEN NULL ELSE CONVERT(decimal(18,4),C5) END,CONVERT(smalldatetime,C6,120),CONVERT(smalldatetime,C7,120) FROM #BillingCard;
 CREATE TABLE #BillingMonthly(RowNo int IDENTITY(1,1),DriverAccount varchar(100),SubtotalAmount money,SubtotalPowerUsage decimal(18,4),PeriodFrom smalldatetime,PeriodThru smalldatetime);
 INSERT #BillingMonthly(DriverAccount,SubtotalAmount,SubtotalPowerUsage,PeriodFrom,PeriodThru) SELECT LTRIM(RTRIM(C1)),CONVERT(money,C2),CASE WHEN ISNULL(C3,'')='' THEN NULL ELSE CONVERT(decimal(18,4),C3) END,CONVERT(smalldatetime,C4,120),CONVERT(smalldatetime,C5,120) FROM @BillingMonthly;

 SELECT 'PCR' SourceCode,COUNT(*) RowCount,COUNT(DISTINCT UnitNo) UnitCount,SUM(Amount) TotalAmount FROM #BillingMember
 UNION ALL SELECT 'ELV-MONTHLY',COUNT(*),COUNT(DISTINCT DriverAccount),SUM(SubtotalAmount) FROM #BillingMonthly
 UNION ALL SELECT 'ELV-CARD',COUNT(*),COUNT(DISTINCT ChargingCardName),SUM(TotalPowerCost) FROM #BillingCard;
 SELECT 'PCR' SourceCode,UnitNo,COUNT(*) SourceRows,SUM(Amount) Amount FROM #BillingMember GROUP BY UnitNo
 UNION ALL SELECT 'ELV-MONTHLY',DriverAccount,COUNT(*),SUM(SubtotalAmount) FROM #BillingMonthly GROUP BY DriverAccount
 UNION ALL SELECT 'ELV-CARD',ChargingCardName,COUNT(*),SUM(TotalPowerCost) FROM #BillingCard GROUP BY ChargingCardName
 ORDER BY SourceCode,UnitNo;
 SELECT 'Billing_member' SourceName,'Uploaded' SourceStatus
 UNION ALL SELECT 'Card Payment Statistics',CASE WHEN EXISTS(SELECT 1 FROM #BillingCard) THEN 'Uploaded' ELSE 'Belum di-upload – dikonfirmasi tidak tersedia' END
 UNION ALL SELECT 'Billing Report Monthly',CASE WHEN EXISTS(SELECT 1 FROM #BillingMonthly) THEN 'Uploaded' ELSE 'Belum di-upload – dikonfirmasi tidak tersedia' END;
END
GO
