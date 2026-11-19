/*
 KTA Web Billing - SQL Server 2008 compatible staging and preview layer
 Branch: feature/web-billing-module
 This layer deliberately does not write KTA_T_H_Billing / AR journals.
*/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

IF OBJECT_ID('dbo.KTA_W_BillingBatch','U') IS NULL
BEGIN
 CREATE TABLE dbo.KTA_W_BillingBatch(
  BatchID uniqueidentifier NOT NULL PRIMARY KEY,
  CompanyID varchar(10) NOT NULL,
  LocationID varchar(3) NOT NULL,
  ActivePeriod smalldatetime NOT NULL,
  UpdateUser varchar(10) NOT NULL,
  CreatedAt datetime NOT NULL CONSTRAINT DF_KTA_W_BillingBatch_CreatedAt DEFAULT(GETDATE()),
  MemberUploaded bit NOT NULL CONSTRAINT DF_KTA_W_BillingBatch_Member DEFAULT(0),
  CardUploaded bit NOT NULL CONSTRAINT DF_KTA_W_BillingBatch_Card DEFAULT(0),
  MonthlyUploaded bit NOT NULL CONSTRAINT DF_KTA_W_BillingBatch_Monthly DEFAULT(0),
  CardUnavailableConfirmed bit NOT NULL CONSTRAINT DF_KTA_W_BillingBatch_CardNA DEFAULT(0),
  MonthlyUnavailableConfirmed bit NOT NULL CONSTRAINT DF_KTA_W_BillingBatch_MonthlyNA DEFAULT(0),
  ValidationStatus tinyint NOT NULL CONSTRAINT DF_KTA_W_BillingBatch_Status DEFAULT(0),
  ValidationMessage varchar(2000) NULL
 );
END
GO
IF OBJECT_ID('dbo.KTA_W_BillingMember','U') IS NULL
BEGIN
 CREATE TABLE dbo.KTA_W_BillingMember(
  BatchID uniqueidentifier NOT NULL, RowNo int NOT NULL,
  UnitNo varchar(5) NOT NULL, DateTrans smalldatetime NOT NULL,
  [Description] varchar(100) NULL, Amount money NOT NULL,
  CONSTRAINT PK_KTA_W_BillingMember PRIMARY KEY(BatchID,RowNo)
 );
 CREATE INDEX IX_KTA_W_BillingMember_BatchUnit ON dbo.KTA_W_BillingMember(BatchID,UnitNo,DateTrans);
END
GO
IF OBJECT_ID('dbo.KTA_W_BillingCard','U') IS NULL
BEGIN
 CREATE TABLE dbo.KTA_W_BillingCard(
  BatchID uniqueidentifier NOT NULL, RowNo int NOT NULL,
  ChargingCardName varchar(100) NOT NULL, ChargingCardNo varchar(100) NULL,
  Currency varchar(3) NULL, TotalPowerCost money NOT NULL,
  TotalPowerConsumption decimal(18,4) NULL,
  PeriodFrom smalldatetime NOT NULL, PeriodThru smalldatetime NOT NULL,
  CONSTRAINT PK_KTA_W_BillingCard PRIMARY KEY(BatchID,RowNo)
 );
 CREATE INDEX IX_KTA_W_BillingCard_BatchUnit ON dbo.KTA_W_BillingCard(BatchID,ChargingCardName);
END
GO
IF OBJECT_ID('dbo.KTA_W_BillingMonthly','U') IS NULL
BEGIN
 CREATE TABLE dbo.KTA_W_BillingMonthly(
  BatchID uniqueidentifier NOT NULL, RowNo int NOT NULL,
  DriverAccount varchar(100) NOT NULL, SubtotalAmount money NOT NULL,
  SubtotalPowerUsage decimal(18,4) NULL,
  PeriodFrom smalldatetime NOT NULL, PeriodThru smalldatetime NOT NULL,
  CONSTRAINT PK_KTA_W_BillingMonthly PRIMARY KEY(BatchID,RowNo)
 );
 CREATE INDEX IX_KTA_W_BillingMonthly_BatchUnit ON dbo.KTA_W_BillingMonthly(BatchID,DriverAccount);
END
GO

IF OBJECT_ID('dbo.usp_KTA_Billing_CreateBatch','P') IS NOT NULL DROP PROCEDURE dbo.usp_KTA_Billing_CreateBatch
GO
CREATE PROCEDURE dbo.usp_KTA_Billing_CreateBatch
 @CompanyID varchar(10), @LocationID varchar(3), @UpdateUser varchar(10),
 @BatchID uniqueidentifier OUTPUT, @ActivePeriod smalldatetime OUTPUT
AS
BEGIN
 SET NOCOUNT ON;
 SELECT @ActivePeriod=ActivePeriod FROM dbo.M_Active_Period
 WHERE AutoNum='KAR'+@CompanyID+@LocationID;
 IF @ActivePeriod IS NULL
 BEGIN
  RAISERROR('Active billing period KAR was not found.',16,1);
  RETURN;
 END
 SET @BatchID=NEWID();
 INSERT dbo.KTA_W_BillingBatch(BatchID,CompanyID,LocationID,ActivePeriod,UpdateUser)
 VALUES(@BatchID,@CompanyID,@LocationID,@ActivePeriod,@UpdateUser);
END
GO

IF OBJECT_ID('dbo.usp_KTA_Billing_ValidateBatch','P') IS NOT NULL DROP PROCEDURE dbo.usp_KTA_Billing_ValidateBatch
GO
CREATE PROCEDURE dbo.usp_KTA_Billing_ValidateBatch
 @BatchID uniqueidentifier
AS
BEGIN
 SET NOCOUNT ON;
 DECLARE @Period smalldatetime,@MonthStart smalldatetime,@MonthEnd smalldatetime,
         @Member bit,@Card bit,@Monthly bit,@CardNA bit,@MonthlyNA bit,
         @Message varchar(2000);
 SELECT @Period=ActivePeriod,@Member=MemberUploaded,@Card=CardUploaded,
        @Monthly=MonthlyUploaded,@CardNA=CardUnavailableConfirmed,
        @MonthlyNA=MonthlyUnavailableConfirmed
 FROM dbo.KTA_W_BillingBatch WHERE BatchID=@BatchID;
 IF @Period IS NULL
 BEGIN
  RAISERROR('Unknown billing BatchID.',16,1);
  RETURN;
 END
 DECLARE @PeriodCAR smalldatetime,@PeriodKAR smalldatetime,@CompanyID varchar(10),@LocationID varchar(3);
 SELECT @CompanyID=CompanyID,@LocationID=LocationID FROM dbo.KTA_W_BillingBatch WHERE BatchID=@BatchID;
 SELECT @PeriodKAR=ActivePeriod FROM dbo.M_Active_Period WHERE AutoNum='KAR'+@CompanyID+@LocationID;
 SELECT @PeriodCAR=ActivePeriod FROM dbo.M_Active_Period WHERE AutoNum='KCAR'+@CompanyID+@LocationID;
 SET @Message='';
 IF @PeriodKAR IS NULL SET @Message=@Message+'Active KAR period was not found. ';
 IF @PeriodKAR IS NOT NULL AND CONVERT(varchar(6),@PeriodKAR,112)<>CONVERT(varchar(6),@Period,112)
   SET @Message=@Message+'Batch period is no longer the current active KAR period. ';
 IF @PeriodCAR IS NULL SET @Message=@Message+'Active KCAR period was not found. ';
 IF @PeriodCAR IS NOT NULL AND CONVERT(varchar(6),@Period,112)<>CONVERT(varchar(6),@PeriodCAR,112)
   SET @Message=@Message+'Active KAR and KCAR periods are not aligned for processing. ';
 SET @MonthStart=DATEADD(month,DATEDIFF(month,0,@Period),0);
 SET @MonthEnd=DATEADD(month,1,@MonthStart);
 IF @Member=0 SET @Message=@Message+'Billing_member file is required. ';
 IF @Card=0 AND @CardNA=0 SET @Message=@Message+'Confirm Card Payment Statistics is unavailable, or upload it. ';
 IF @Monthly=0 AND @MonthlyNA=0 SET @Message=@Message+'Confirm Billing Report Monthly is unavailable, or upload it. ';
 IF EXISTS(SELECT 1 FROM dbo.KTA_W_BillingMember WHERE BatchID=@BatchID
           AND (DateTrans<@MonthStart OR DateTrans>=@MonthEnd))
  SET @Message=@Message+'Billing_member contains transaction dates outside active KAR period. ';
 IF @Card=1 AND EXISTS(SELECT 1 FROM dbo.KTA_W_BillingCard WHERE BatchID=@BatchID
           AND (PeriodFrom<@MonthStart OR PeriodThru>=@MonthEnd OR PeriodFrom>PeriodThru))
  SET @Message=@Message+'Card Payment Statistics contains invalid Statistical Period. ';
 IF @Monthly=1 AND EXISTS(SELECT 1 FROM dbo.KTA_W_BillingMonthly WHERE BatchID=@BatchID
           AND (PeriodFrom<@MonthStart OR PeriodThru>=@MonthEnd OR PeriodFrom>PeriodThru))
  SET @Message=@Message+'Billing Report Monthly contains invalid From-Thru period. ';
 IF @Member=1 AND EXISTS(
    SELECT 1 FROM dbo.KTA_W_BillingMember m
    WHERE m.BatchID=@BatchID AND NOT EXISTS(
      SELECT 1 FROM dbo.KTA_M_Unit u
      WHERE u.CompanyID=@CompanyID AND u.LocationID=@LocationID AND u.UnitNo=m.UnitNo
        AND u.Tower<>'' AND u.OwnerID<>''
        AND EXISTS(SELECT 1 FROM dbo.KTA_T_H_Billing h
                   WHERE h.CompanyID=u.CompanyID AND h.LocationID=u.LocationID
                     AND h.UnitNo=u.UnitNo
                     AND CONVERT(varchar(6),h.BillDate,112)=
                         CONVERT(varchar(6),DATEADD(month,-1,@Period),112))
    ))
  SET @Message=@Message+'Billing_member contains units that are not eligible for this billing run. ';

 IF @Card=1 AND EXISTS(
    SELECT 1 FROM dbo.KTA_W_BillingCard c
    WHERE c.BatchID=@BatchID AND NOT EXISTS(
      SELECT 1 FROM dbo.KTA_M_Unit u
      WHERE u.CompanyID=@CompanyID AND u.LocationID=@LocationID AND u.UnitNo=c.ChargingCardName
        AND u.Tower<>'' AND u.OwnerID<>''
        AND EXISTS(SELECT 1 FROM dbo.KTA_T_H_Billing h
                   WHERE h.CompanyID=u.CompanyID AND h.LocationID=u.LocationID
                     AND h.UnitNo=u.UnitNo
                     AND CONVERT(varchar(6),h.BillDate,112)=
                         CONVERT(varchar(6),DATEADD(month,-1,@Period),112))
    ))
  SET @Message=@Message+'Card Payment Statistics contains Charging Card Name values that are not eligible unit numbers. ';

 IF @Monthly=1 AND EXISTS(
    SELECT 1 FROM dbo.KTA_W_BillingMonthly m
    WHERE m.BatchID=@BatchID AND NOT EXISTS(
      SELECT 1 FROM dbo.KTA_M_Unit u
      WHERE u.CompanyID=@CompanyID AND u.LocationID=@LocationID AND u.UnitNo=m.DriverAccount
        AND u.Tower<>'' AND u.OwnerID<>''
        AND EXISTS(SELECT 1 FROM dbo.KTA_T_H_Billing h
                   WHERE h.CompanyID=u.CompanyID AND h.LocationID=u.LocationID
                     AND h.UnitNo=u.UnitNo
                     AND CONVERT(varchar(6),h.BillDate,112)=
                         CONVERT(varchar(6),DATEADD(month,-1,@Period),112))
    ))
  SET @Message=@Message+'Billing Report Monthly contains Driver Account values that are not eligible unit numbers. ';

 IF @Member=1 AND NOT EXISTS(SELECT 1 FROM dbo.KTA_W_BillingMember WHERE BatchID=@BatchID)
  SET @Message=@Message+'Billing_member has no data rows. ';
 IF @Card=1 AND NOT EXISTS(SELECT 1 FROM dbo.KTA_W_BillingCard WHERE BatchID=@BatchID)
  SET @Message=@Message+'Card Payment Statistics has no data rows. ';
 IF @Monthly=1 AND NOT EXISTS(SELECT 1 FROM dbo.KTA_W_BillingMonthly WHERE BatchID=@BatchID)
  SET @Message=@Message+'Billing Report Monthly has no data rows. ';
 UPDATE dbo.KTA_W_BillingBatch
 SET ValidationStatus=CASE WHEN @Message='' THEN 1 ELSE 2 END,
     ValidationMessage=NULLIF(@Message,'')
 WHERE BatchID=@BatchID;
 SELECT ValidationStatus,ValidationMessage,ActivePeriod
 FROM dbo.KTA_W_BillingBatch WHERE BatchID=@BatchID;
END
GO

IF OBJECT_ID('dbo.usp_KTA_Billing_Preview','P') IS NOT NULL DROP PROCEDURE dbo.usp_KTA_Billing_Preview
GO
CREATE PROCEDURE dbo.usp_KTA_Billing_Preview
 @BatchID uniqueidentifier
AS
BEGIN
 SET NOCOUNT ON;
 IF NOT EXISTS(SELECT 1 FROM dbo.KTA_W_BillingBatch WHERE BatchID=@BatchID AND ValidationStatus=1)
 BEGIN
  RAISERROR('Batch must pass validation before preview.',16,1);
  RETURN;
 END
 DECLARE @Card bit,@Monthly bit;
 SELECT @Card=CardUploaded,@Monthly=MonthlyUploaded FROM dbo.KTA_W_BillingBatch WHERE BatchID=@BatchID;

 -- Result set 1: summary totals by source.
 SELECT 'PCR' SourceCode,COUNT(*) RowCount,COUNT(DISTINCT UnitNo) UnitCount,
        SUM(Amount) TotalAmount FROM dbo.KTA_W_BillingMember WHERE BatchID=@BatchID
 UNION ALL
 SELECT 'ELV-MONTHLY',COUNT(*),COUNT(DISTINCT DriverAccount),SUM(SubtotalAmount)
 FROM dbo.KTA_W_BillingMonthly WHERE BatchID=@BatchID
 UNION ALL
 SELECT 'ELV-CARD',COUNT(*),COUNT(DISTINCT ChargingCardName),SUM(TotalPowerCost)
 FROM dbo.KTA_W_BillingCard WHERE BatchID=@BatchID;

 -- Result set 2: unit-level detail, grouped/summed where source rows can repeat.
 SELECT 'PCR' SourceCode,UnitNo,COUNT(*) SourceRows,SUM(Amount) Amount
 FROM dbo.KTA_W_BillingMember WHERE BatchID=@BatchID GROUP BY UnitNo
 UNION ALL
 SELECT 'ELV-MONTHLY',DriverAccount,COUNT(*),SUM(SubtotalAmount)
 FROM dbo.KTA_W_BillingMonthly WHERE BatchID=@BatchID GROUP BY DriverAccount
 UNION ALL
 SELECT 'ELV-CARD',ChargingCardName,COUNT(*),SUM(TotalPowerCost)
 FROM dbo.KTA_W_BillingCard WHERE BatchID=@BatchID GROUP BY ChargingCardName
 ORDER BY SourceCode,UnitNo;

 -- Result set 3: optional-source status for the future XLSX validation sheet.
 SELECT 'Billing_member' SourceName,'Uploaded' SourceStatus
 UNION ALL SELECT 'Card Payment Statistics',
   CASE WHEN @Card=1 THEN 'Uploaded' ELSE 'Belum di-upload – dikonfirmasi tidak tersedia' END
 UNION ALL SELECT 'Billing Report Monthly',
   CASE WHEN @Monthly=1 THEN 'Uploaded' ELSE 'Belum di-upload – dikonfirmasi tidak tersedia' END;
END
GO
