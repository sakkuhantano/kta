USE [KTA]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

IF OBJECT_ID('dbo.usp_KTA_Billing_Post','P') IS NOT NULL DROP PROCEDURE dbo.usp_KTA_Billing_Post
GO
CREATE PROCEDURE dbo.usp_KTA_Billing_Post
 @BatchID uniqueidentifier,
 @CompanyID varchar(10),
 @LocationID varchar(3),
 @UpdateUser varchar(10),
 @Dummy bit = 0
AS
BEGIN
 SET NOCOUNT ON;
 SET XACT_ABORT ON;

 DECLARE @Period smalldatetime,@PeriodKAR smalldatetime,@PeriodCAR smalldatetime,@PrevPeriod smalldatetime,
         @BillDate smalldatetime,@DueDate smalldatetime,@TermDue int,
         @Penalty money,@MinPEC money,@HardCopyCharges money,@PPN money,
         @AcctRate money,@ProcessStatus tinyint,@Err varchar(4000);

 SELECT @Period=ActivePeriod FROM dbo.KTA_W_BillingBatch
 WHERE BatchID=@BatchID AND CompanyID=@CompanyID AND LocationID=@LocationID;
 IF @Period IS NULL
 BEGIN
  RAISERROR('Invalid billing batch/company/location.',16,1);
  RETURN;
 END

 SELECT @PeriodKAR=ActivePeriod FROM dbo.M_Active_Period
 WHERE AutoNum='KAR'+@CompanyID+@LocationID;
 IF @PeriodKAR IS NULL OR CONVERT(varchar(6),@PeriodKAR,112)<>CONVERT(varchar(6),@Period,112)
 BEGIN
  RAISERROR('Billing batch period is no longer the active KAR period.',16,1);
  RETURN;
 END

 SELECT @PeriodCAR=ActivePeriod FROM dbo.M_Active_Period
 WHERE AutoNum='KCAR'+@CompanyID+@LocationID;
 IF @PeriodCAR IS NULL
 BEGIN
  RAISERROR('Active CAR period was not found.',16,1);
  RETURN;
 END

 SELECT @TermDue=CAST(Term AS int),@Penalty=Penalty,@MinPEC=ISNULL(MinPEC,0),
        @HardCopyCharges=ISNULL(HardCopyCharges,0)
 FROM dbo.KTA_M_Penalty WHERE CompanyID=@CompanyID AND LocationID=@LocationID;
 IF @TermDue IS NULL
 BEGIN
  RAISERROR('KTA_M_Penalty not found.',16,1);
  RETURN;
 END

 SELECT TOP 1 @AcctRate=AcctRate FROM dbo.M_Rate
 WHERE CONVERT(varchar(6),[Date],112)=CONVERT(varchar(6),@PeriodCAR,112)
 ORDER BY [Date];
 IF @AcctRate IS NULL
 BEGIN
  RAISERROR('M_Rate for active CAR period was not found.',16,1);
  RETURN;
 END

 -- VB6 requires telephone billing rows for the last day of the month before KAR-1.
 IF NOT EXISTS(SELECT 1 FROM dbo.KTA_T_BillingPhone
               WHERE CONVERT(varchar(8),BillDate,112)=
                     CONVERT(varchar(8),DATEADD(day,-1,DATEADD(month,-1,DATEADD(day,1,@Period))),112))
 BEGIN
  RAISERROR('Telephone billing source is missing for the required prior period.',16,1);
  RETURN;
 END

 IF NOT EXISTS(SELECT 1 FROM dbo.KTA_W_BillingBatch WHERE BatchID=@BatchID AND ValidationStatus=1)
 BEGIN
  RAISERROR('Billing batch has not passed validation.',16,1);
  RETURN;
 END

 IF CONVERT(varchar(6),@Period,112)<>CONVERT(varchar(6),@PeriodCAR,112)
 BEGIN
  RAISERROR('Process can be done after EOM: active KAR and KCAR periods are not the same.',16,1);
  RETURN;
 END

 DECLARE @BillDateDay tinyint;
 SELECT @BillDateDay=BillDate FROM dbo.KTA_M_Penalty
 WHERE CompanyID=@CompanyID AND LocationID=@LocationID;
 IF @BillDateDay IS NULL OR @BillDateDay=0
 BEGIN
  RAISERROR('KTA_M_Penalty.BillDate is not configured.',16,1);
  RETURN;
 END
 SET @BillDate=DATEADD(day,@BillDateDay-1,DATEADD(month,DATEDIFF(month,0,@Period),0));
 SET @DueDate=DATEADD(day,@TermDue-1,@BillDate);
 SET @PrevPeriod=DATEADD(day,-1,DATEADD(month,DATEDIFF(month,0,@Period),0));

 BEGIN TRANSACTION;
 BEGIN TRY

   IF EXISTS(SELECT 1 FROM dbo.KTA_T_ProcessBilling WHERE CompanyID=@CompanyID
             AND LocationID=@LocationID AND Period=CONVERT(varchar(6),@PeriodCAR,112))
      RAISERROR('Billing period has already been processed.',16,1);

   -- PPN follows the legacy VB6 constant: Private Const PPN As Double = (11 / 100).
   -- Keep this hardcoded for now; move to configuration only when the legacy source is intentionally changed.
   SET @PPN = 0.11;

   -- The number is generated only during posting, inside this transaction.
   -- Required convention: dbo.sp_GetGenerateAutoNumID 'arkta', KBI+Company+Location.
   -- This block is intentionally isolated so preview never consumes a number.
   DECLARE @Units TABLE(
      RowID int IDENTITY(1,1) PRIMARY KEY,
      UnitNo varchar(5),OwnerID varchar(10),CurrencyID varchar(3),
      Area money,MaintenanceFee money,SinkingFund money,
      SendHardCopy varchar(1),HandOverDate smalldatetime,
      Previous money,Payment money
   );

   INSERT @Units(UnitNo,OwnerID,CurrencyID,Area,MaintenanceFee,SinkingFund,SendHardCopy,
                 HandOverDate,Previous,Payment)
   SELECT u.UnitNo,u.OwnerID,r.CurrencyID,r.Area,r.MaintenanceFee,r.SinkingFund,
          u.SendHardCopy,u.HandOverDate,
          ISNULL(pb.NewBalance,0),ISNULL(p.Payment,0)
   FROM dbo.KTA_M_Unit u
   INNER JOIN dbo.KTA_M_RoomType r ON r.CompanyID=u.CompanyID AND r.LocationID=u.LocationID
                                  AND r.RoomType=u.RoomType
   LEFT JOIN dbo.KTA_T_H_Billing pb ON pb.CompanyID=u.CompanyID AND pb.LocationID=u.LocationID
                                   AND pb.UnitNo=u.UnitNo
                                   AND CONVERT(varchar(6),pb.BillDate,112)=CONVERT(varchar(6),@PrevPeriod,112)
   LEFT JOIN (
      SELECT h.CompanyID,h.LocationID,h.UnitNo,
             SUM(CASE WHEN d.RefNo='-' THEN 0 ELSE d.Amount END) Payment
      FROM dbo.KTA_T_D_Outstanding d
      INNER JOIN dbo.vw_KTA_T_H_Outstanding h ON h.KeyHD=d.KeyHD
      WHERE h.CompanyID=@CompanyID AND h.LocationID=@LocationID
        AND h.TxType IN ('W','B') AND h.LocationNo=''
        AND d.RecordType IN ('3','4','5') AND d.RefNo<>'' AND d.RefDate<=h.EndDate
        AND ((h.TxType='W' AND CONVERT(varchar(6),h.DocumentDate,112)=CONVERT(varchar(6),h.PrevPeriod,112)
              AND CONVERT(varchar(6),d.RefDate,112)>=CONVERT(varchar(6),h.PrevPeriod,112))
          OR (h.TxType='W' AND CONVERT(varchar(6),h.DocumentDate,112)<CONVERT(varchar(6),h.PrevPeriod,112)
              AND d.RefDate>=h.StartDate)
          OR (h.TxType='B' AND d.RefDate>=h.StartDate))
      GROUP BY h.CompanyID,h.LocationID,h.UnitNo
   ) p ON p.CompanyID=u.CompanyID AND p.LocationID=u.LocationID AND p.UnitNo=u.UnitNo
   WHERE u.CompanyID=@CompanyID AND u.LocationID=@LocationID
     AND u.Tower<>'' AND u.OwnerID<>''
     AND EXISTS(SELECT 1 FROM dbo.KTA_T_H_Billing x
                WHERE x.CompanyID=u.CompanyID AND x.LocationID=u.LocationID
                  AND x.UnitNo=u.UnitNo
                  AND CONVERT(varchar(6),x.BillDate,112)=CONVERT(varchar(6),@PrevPeriod,112));

   /*
     Load LPC source rows once for all eligible units.
     This is the set-based equivalent of sp_KTA_selectLPC, avoiding 2,700
     separate stored-procedure executions while retaining its filters/order.
   */
   DECLARE @LPCPeriodText varchar(6);
   SET @LPCPeriodText=CONVERT(varchar(6),@Period,112);

   CREATE TABLE #LPC(
      KeyHD int,DocumentNo varchar(10),DocumentDate smalldatetime,DueDate smalldatetime,
      IncomeCode varchar(3),CurrencyID varchar(3),UnitNo varchar(10),OwnerID varchar(10),
      BillDate smalldatetime,RecordType varchar(2),RefNo varchar(10),RKDate smalldatetime,
      Amount money,StartDate smalldatetime,EndDate smalldatetime,OffsetAmount money
   );

   INSERT #LPC
   SELECT h.KeyHD,h.DocumentNo,h.DocumentDate,h.DueDate,h.IncomeCode,h.CurrencyID,h.UnitNo,h.OwnerID,
          ISNULL(z.BillDate,h.DocumentDate),d.RecordType,d.RefNo,d.RKDate,d.Amount,
          CAST(CONVERT(varchar(6),@PrevPeriod,112)+RIGHT('00'+CAST(
             CASE WHEN @LPCPeriodText<=CONVERT(varchar(6),z1.StartingNewBilling,112)
                  THEN z1.PreviousBillDate ELSE z1.BillDate END AS varchar(2)),2) AS smalldatetime),
          DATEADD(day,-1,CAST(@LPCPeriodText+RIGHT('00'+CAST(z1.BillDate AS varchar(2)),2) AS smalldatetime)),
          ISNULL(dd.Amount,0)
   FROM dbo.KTA_T_H_Outstanding h
   INNER JOIN dbo.KTA_T_D_Outstanding d ON d.KeyHD=h.KeyHD
   INNER JOIN (
       dbo.KTA_M_Owner x
       INNER JOIN dbo.KTA_M_Unit y
         ON x.CompanyID=y.CompanyID AND x.LocationID=y.LocationID AND x.OwnerID=y.OwnerID
   ) ON y.CompanyID=h.CompanyID AND y.LocationID=h.LocationID AND y.UnitNo=h.UnitNo
   INNER JOIN dbo.KTA_M_Penalty z1 ON z1.CompanyID=h.CompanyID AND z1.LocationID=h.LocationID
   INNER JOIN @Units eligible ON eligible.UnitNo=h.UnitNo
   LEFT JOIN dbo.KTA_T_D_Outstanding dd
     INNER JOIN dbo.KTA_T_H_Outstanding hh
       ON dd.KeyHD=hh.KeyHD AND dd.RecordType='1' AND dd.Amount<0
     ON hh.DocumentNo=h.DocumentNo AND hh.TxType=h.TxType AND hh.UnitNo=h.UnitNo
   LEFT JOIN (
       SELECT bh.CompanyID,bh.LocationID,bh.UnitNo,bh.BillDate,bd.InvoiceNo,bd.IncomeCode
       FROM dbo.KTA_T_D_Billing bd
       INNER JOIN dbo.KTA_T_H_Billing bh ON bd.KeyHD=bh.KeyHD
       WHERE bd.IncomeCode IN ('WOL','WOM')
   ) z ON z.CompanyID=h.CompanyID AND z.LocationID=h.LocationID AND z.UnitNo=h.UnitNo
       AND z.IncomeCode=h.IncomeCode AND z.InvoiceNo=h.DocumentNo
   WHERE x.Penalty='y'
     AND h.CompanyID=@CompanyID AND h.LocationID=@LocationID
     AND h.UnitNo IN (SELECT UnitNo FROM @Units)
     AND (h.PaidDate='' OR CONVERT(varchar(6),h.PaidDate,112)=CONVERT(varchar(6),@PrevPeriod,112))
     AND d.RefNo<>''
     AND EXISTS(
       SELECT a.CompanyID,a.LocationID,a.DocumentNo,a.TxType
       FROM dbo.KTA_T_H_Outstanding a
       INNER JOIN dbo.KTA_T_D_Outstanding b ON a.KeyHD=b.KeyHD
       WHERE b.RefNo<>''
         AND a.CompanyID=h.CompanyID AND a.LocationID=h.LocationID
         AND a.DocumentNo=h.DocumentNo AND a.TxType=h.TxType AND a.OwnerID=h.OwnerID
         AND a.DueDate<=DATEADD(day,-1,CAST(@LPCPeriodText+RIGHT('00'+CAST(z1.BillDate AS varchar(2)),2) AS smalldatetime))
         AND CASE WHEN b.RecordType='1' THEN b.RefDate ELSE b.RKDate END<=a.DueDate
         AND CASE WHEN b.RecordType='1' THEN b.RefDate ELSE b.RKDate END<=DATEADD(day,-1,CAST(@LPCPeriodText+RIGHT('00'+CAST(z1.BillDate AS varchar(2)),2) AS smalldatetime))
       GROUP BY a.CompanyID,a.LocationID,a.DocumentNo,a.TxType
       HAVING SUM(CASE WHEN b.RecordType='1' THEN b.Amount ELSE -b.Amount END)>0
     );

   CREATE CLUSTERED INDEX IX_LPC_Unit_KeyHD ON #LPC(UnitNo,KeyHD,RecordType,RKDate);

   DECLARE @RowID int,@MaxRowID int;
   SELECT @RowID=1,@MaxRowID=COUNT(*) FROM @Units;

   WHILE @RowID<=@MaxRowID
   BEGIN
      DECLARE @UnitNo varchar(5),@OwnerID varchar(10),@CurrencyID varchar(3),
              @Area money,@MFee money,@SFund money,@SendHardCopy varchar(1),
              @HandOverDate smalldatetime,@Previous money,@Payment money,
              @KeyHD int,@BillingNo varchar(10),@Seq int,@Total money;

      SELECT @UnitNo=UnitNo,@OwnerID=OwnerID,@CurrencyID=CurrencyID,@Area=Area,
             @MFee=MaintenanceFee,@SFund=SinkingFund,@SendHardCopy=SendHardCopy,
             @HandOverDate=HandOverDate,@Previous=Previous,@Payment=Payment
      FROM @Units WHERE RowID=@RowID;

      IF @Dummy=1
      BEGIN
        -- Dummy IDs are deterministic within the rolled-back preview transaction.
        -- Never call the production autonumber generator for a dummy statement.
        SET @BillingNo='D'+RIGHT('000000000'+CONVERT(varchar(9),@RowID),9);
      END
      ELSE
      BEGIN
        DECLARE @AutoNumResult TABLE(ID varchar(10),SeqNo int);
        DECLARE @AutoNum varchar(20);
        SET @AutoNum='KBI'+@CompanyID+@LocationID;
        INSERT INTO @AutoNumResult
        EXECUTE dbo.sp_GetGenerateAutoNumID 'arkta',@AutoNum,1;
        SELECT @BillingNo=ID FROM @AutoNumResult;
        IF ISNULL(@BillingNo,'')='' RAISERROR('Failed to generate BillingNo.',16,1);
      END

      INSERT dbo.KTA_T_H_Billing(CompanyID,LocationID,BillingNo,UnitNo,BillDate,DueDate,
        CurrencyID,Previous,Payment,NewBilling,NewBalance,PrintCode,RePrint,UpdateUser,UpdateDate)
      VALUES(@CompanyID,@LocationID,@BillingNo,@UnitNo,@BillDate,@DueDate,@CurrencyID,
        @Previous,@Payment,0,0,0,0,@UpdateUser,GETDATE());
      SET @KeyHD=SCOPE_IDENTITY();
      SET @Seq=0; SET @Total=0;

      SET @Seq=@Seq+1;
      INSERT dbo.KTA_T_D_Billing(KeyHD,SeqNo,IncomeCode,Amount,InvoiceNo,InvoiceDate,PhoneNo,Description)
      VALUES(@KeyHD,@Seq,'MFE',@Area*(@MFee+@SFund),'','19000101','',NULL);
      SET @Total=@Total+@Area*(@MFee+@SFund);

      IF @SendHardCopy='Y'
      BEGIN
        SET @Seq=@Seq+1;
        INSERT dbo.KTA_T_D_Billing(KeyHD,SeqNo,IncomeCode,Amount,InvoiceNo,InvoiceDate,PhoneNo,Description)
        VALUES(@KeyHD,@Seq,'PHC',@HardCopyCharges,'','19000101','',NULL);
        SET @Total=@Total+@HardCopyCharges;
      END

      INSERT dbo.KTA_T_D_Billing(KeyHD,SeqNo,IncomeCode,Amount,InvoiceNo,InvoiceDate,PhoneNo,Description)
      SELECT @KeyHD,@Seq+ROW_NUMBER() OVER(ORDER BY d.RefDate,d.RefNo),
             h.IncomeCode,d.Amount,d.RefNo,d.RefDate,'',NULL
      FROM dbo.KTA_T_H_Outstanding h
      INNER JOIN dbo.KTA_T_D_Outstanding d ON d.KeyHD=h.KeyHD
      WHERE h.CompanyID=@CompanyID AND h.LocationID=@LocationID AND h.UnitNo=@UnitNo
        AND h.TxType='W' AND h.LocationNo=''
        AND d.RecordType='1'
        AND CONVERT(varchar(6),h.DocumentDate,112)=CONVERT(varchar(6),@PrevPeriod,112);
      SELECT @Seq=ISNULL(MAX(SeqNo),@Seq),@Total=@Total+ISNULL(SUM(Amount),0)
      FROM dbo.KTA_T_D_Billing WHERE KeyHD=@KeyHD;

      INSERT dbo.KTA_T_D_Billing(KeyHD,SeqNo,IncomeCode,Amount,InvoiceNo,InvoiceDate,PhoneNo,Description)
      SELECT @KeyHD,@Seq+ROW_NUMBER() OVER(ORDER BY Period,Description),
             IncomeCode,Amount+VAT,'','19000101','',Description
      FROM dbo.KTA_T_CollectCall
      WHERE CompanyID=@CompanyID AND LocationID=@LocationID AND UnitNo=@UnitNo
        AND CONVERT(varchar(6),Period,112)=CONVERT(varchar(6),@PrevPeriod,112);
      SELECT @Seq=ISNULL(MAX(SeqNo),@Seq),@Total=@Total+ISNULL(SUM(Amount),0)
      FROM dbo.KTA_T_D_Billing WHERE KeyHD=@KeyHD;

      INSERT dbo.KTA_T_D_Billing_Phone(KeyHD,PhoneNo,Fee,Local,SLJJ,International,Airtime,JaPaTi,SubTotal,
        Pulsa_Local,Pulsa_SLJJ,BeaAdmin,PPN,Other,Total)
      SELECT @KeyHD,p.PhoneNo,p.Fee,p.Local,p.SLJJ,p.International,p.Airtime,p.JaPaTi,
        p.Fee+p.Local+p.SLJJ+p.International+p.Airtime+p.JaPaTi,
        p.Pulsa_Local,p.Pulsa_SLJJ,
        ROUND((p.Fee+p.Local+p.SLJJ+p.International+p.Airtime+p.JaPaTi)/9,0),
        ROUND(((p.Fee+p.Local+p.SLJJ+p.International+p.Airtime+p.JaPaTi)+
              ROUND((p.Fee+p.Local+p.SLJJ+p.International+p.Airtime+p.JaPaTi)/9,0))*0,0),
        0,0
      FROM dbo.KTA_T_BillingPhone p
      WHERE p.UnitNo=@UnitNo
        AND CONVERT(varchar(8),p.BillDate,112)=CONVERT(varchar(8),
             DATEADD(day,-1,DATEADD(month,-1,DATEADD(day,1,@Period))),112);

      UPDATE bp SET PPN=ROUND((bp.SubTotal+bp.BeaAdmin)*ISNULL(@PPN,0),0),
                    Total=bp.SubTotal+bp.BeaAdmin+ROUND((bp.SubTotal+bp.BeaAdmin)*ISNULL(@PPN,0),0)+bp.Other
      FROM dbo.KTA_T_D_Billing_Phone bp WHERE bp.KeyHD=@KeyHD;
      INSERT dbo.KTA_T_D_Billing(KeyHD,SeqNo,IncomeCode,Amount,InvoiceNo,InvoiceDate,PhoneNo,Description)
      SELECT @KeyHD,@Seq+ROW_NUMBER() OVER(ORDER BY PhoneNo),'TEC',Total,'','19000101',PhoneNo,''
      FROM dbo.KTA_T_D_Billing_Phone WHERE KeyHD=@KeyHD;
      SELECT @Seq=ISNULL(MAX(SeqNo),@Seq),@Total=@Total+ISNULL(SUM(Amount),0)
      FROM dbo.KTA_T_D_Billing WHERE KeyHD=@KeyHD;

      INSERT dbo.KTA_T_D_Billing(KeyHD,SeqNo,IncomeCode,Amount,InvoiceNo,InvoiceDate,PhoneNo,Description)
      SELECT @KeyHD,@Seq+ROW_NUMBER() OVER(ORDER BY DateTrans,[Description],RowNo),'PCR',Amount,'','19000101','',LEFT([Description],100)
      FROM dbo.KTA_W_BillingMember WHERE BatchID=@BatchID AND UnitNo=@UnitNo
        AND DateTrans>=DATEADD(month,DATEDIFF(month,0,@Period),0)
        AND DateTrans<DATEADD(month,DATEDIFF(month,0,@Period)+1,0);
      SELECT @Seq=ISNULL(MAX(SeqNo),@Seq),@Total=@Total+ISNULL(SUM(Amount),0)
      FROM dbo.KTA_T_D_Billing WHERE KeyHD=@KeyHD;

      IF EXISTS(SELECT 1 FROM dbo.KTA_W_BillingCard WHERE BatchID=@BatchID)
      BEGIN
        INSERT dbo.KTA_T_D_Billing(KeyHD,SeqNo,IncomeCode,Amount,InvoiceNo,InvoiceDate,PhoneNo,Description)
        SELECT @KeyHD,@Seq+ROW_NUMBER() OVER(ORDER BY RowNo),'ELV',ROUND(TotalPowerCost,0),'','19000101','',
               'Electricity Reimbursement for EV Charging EV-2, Usage Period '+
               DATENAME(month,PeriodFrom)+', '+CAST(YEAR(PeriodFrom) AS varchar)
        FROM dbo.KTA_W_BillingCard WHERE BatchID=@BatchID AND ChargingCardName=@UnitNo;
        SELECT @Seq=ISNULL(MAX(SeqNo),@Seq),@Total=@Total+ISNULL(SUM(Amount),0)
        FROM dbo.KTA_T_D_Billing WHERE KeyHD=@KeyHD;
      END

      IF EXISTS(SELECT 1 FROM dbo.KTA_W_BillingMonthly WHERE BatchID=@BatchID)
      BEGIN
        INSERT dbo.KTA_T_D_Billing(KeyHD,SeqNo,IncomeCode,Amount,InvoiceNo,InvoiceDate,PhoneNo,Description)
        SELECT @KeyHD,@Seq+ROW_NUMBER() OVER(ORDER BY RowNo),'ELV',ROUND(SubtotalAmount,0),'','19000101','',
               'Electricity Reimbursement for EV Charging EV-1, Usage Period '+
               DATENAME(month,PeriodFrom)+', '+CAST(YEAR(PeriodFrom) AS varchar)
        FROM dbo.KTA_W_BillingMonthly WHERE BatchID=@BatchID AND DriverAccount=@UnitNo;
        SELECT @Seq=ISNULL(MAX(SeqNo),@Seq),@Total=@Total+ISNULL(SUM(Amount),0)
        FROM dbo.KTA_T_D_Billing WHERE KeyHD=@KeyHD;
      END

      IF @HandOverDate<>'19000101' AND CONVERT(varchar(6),@HandOverDate,112)<CONVERT(varchar(6),@PeriodCAR,112)
      BEGIN
        INSERT dbo.KTA_T_D_Billing(KeyHD,SeqNo,IncomeCode,Amount,InvoiceNo,InvoiceDate,PhoneNo,Description)
        SELECT @KeyHD,@Seq+ROW_NUMBER() OVER(ORDER BY d.KeyHD),'ELC',d.Amount,'','19000101','',
               'Electricity Reimbursement, usage period '+DATENAME(month,h.BillDate)+', '+CAST(YEAR(h.BillDate) AS varchar)
        FROM dbo.AR_T_H_Billing h INNER JOIN dbo.AR_T_D_Billing d ON d.KeyHD=h.KeyHD
        WHERE h.CompanyID=@CompanyID AND h.LocationID=@LocationID
          AND CONVERT(varchar(6),h.BillDate,112)=CONVERT(varchar(6),@PrevPeriod,112)
          AND h.TenantID=@UnitNo AND d.IncomeCode='elc'
          AND h.Cancel='n' AND h.Flag='' AND h.SendNo<>'';
        SELECT @Seq=ISNULL(MAX(SeqNo),@Seq),@Total=@Total+ISNULL(SUM(Amount),0)
        FROM dbo.KTA_T_D_Billing WHERE KeyHD=@KeyHD;

        INSERT dbo.KTA_T_D_Billing(KeyHD,SeqNo,IncomeCode,Amount,InvoiceNo,InvoiceDate,PhoneNo,Description)
        SELECT @KeyHD,@Seq+ROW_NUMBER() OVER (ORDER BY d.KeyHD),'AIR',d.Amount,'','19000101','',
               'Water Reimbursement, usage period '+DATENAME(month,h.BillDate)+', '+CAST(YEAR(h.BillDate) AS varchar)
        FROM dbo.AR_T_H_Billing h INNER JOIN dbo.AR_T_D_Billing d ON d.KeyHD=h.KeyHD
        WHERE h.CompanyID=@CompanyID AND h.LocationID=@LocationID
          AND CONVERT(varchar(6),h.BillDate,112)=CONVERT(varchar(6),@PrevPeriod,112)
          AND h.TenantID=@UnitNo AND d.IncomeCode='air'
          AND h.Cancel='n' AND h.Flag='' AND h.SendNo<>'';
        SELECT @Seq=ISNULL(MAX(SeqNo),@Seq),@Total=@Total+ISNULL(SUM(Amount),0)
        FROM dbo.KTA_T_D_Billing WHERE KeyHD=@KeyHD;
      END

      /* Legacy VB6 LPC state machine over the batch-wide #LPC rows. */
      DECLARE @LPCBreak int=9,@LPCKeyHD int,@LPCDueDate smalldatetime,@LPCOuts money,
              @LPCMonth int=DAY(@PrevPeriod),@LPCTotPenalty money=0,@LPCDoc varchar(10),
              @LPCStart smalldatetime,@LPCEnd smalldatetime,@LPCCharges money=0,
              @LPCSeq int,@LPCCal money,@LPCBillDate smalldatetime,@LPCOwner varchar(10),
              @LPCUnit varchar(10),@LPCIncome varchar(3),@LPCCurrency varchar(3),
              @LPCOuts1 money=0,@LPCIncome1 varchar(3)='',@LPCRKDate1 smalldatetime,
              @LPCPeriod1 smalldatetime;

      SET @LPCPeriod1=CASE WHEN MONTH(@Period)=3
                           THEN DATEADD(day,1,DATEADD(month,DATEDIFF(month,0,@Period),0))
                           ELSE @PrevPeriod END;

      DECLARE lpc_cur CURSOR LOCAL FAST_FORWARD FOR
       SELECT KeyHD,DocumentNo,DocumentDate,DueDate,IncomeCode,CurrencyID,UnitNo,OwnerID,
              BillDate,RecordType,RefNo,RKDate,Amount,StartDate,EndDate,OffsetAmount
       FROM #LPC
       WHERE UnitNo=@UnitNo
       ORDER BY KeyHD,CASE WHEN RecordType='1' THEN 1 WHEN RecordType='3' THEN 2 ELSE 3 END,RKDate;

      DECLARE @nKey int,@nDoc varchar(10),@nDocDate smalldatetime,@nDue smalldatetime,
              @nIncome varchar(3),@nCur varchar(3),@nUnit varchar(10),@nOwner varchar(10),
              @nBill smalldatetime,@nRT varchar(2),@nRef varchar(10),@nRK smalldatetime,
              @nAmt money,@nStart smalldatetime,@nEnd smalldatetime,@nOffset money;

      OPEN lpc_cur;
      FETCH NEXT FROM lpc_cur INTO @nKey,@nDoc,@nDocDate,@nDue,@nIncome,@nCur,@nUnit,@nOwner,
        @nBill,@nRT,@nRef,@nRK,@nAmt,@nStart,@nEnd,@nOffset;

      WHILE @@FETCH_STATUS=0
      BEGIN
       IF @LPCBreak>0
       BEGIN
        SET @LPCCurrency=@nCur; SET @LPCOuts=@nAmt; SET @LPCKeyHD=@nKey; SET @LPCSeq=0;
        SET @LPCDueDate=@nDue; SET @LPCDoc=@nDoc; SET @LPCStart=@nStart; SET @LPCEnd=@nEnd;
        SET @LPCBillDate=@nDocDate; SET @LPCOwner=@nOwner; SET @LPCUnit=@nUnit; SET @LPCIncome=@nIncome;
        IF @LPCBreak>1 SET @LPCTotPenalty=0;
       END

       IF @nRT='4' AND @nRK>=@LPCStart AND @nRK<=@LPCEnd AND @nRK>@LPCDueDate
       BEGIN
        SET @LPCCal=CASE WHEN @LPCOuts>@nAmt THEN @nAmt ELSE @LPCOuts END;
        SET @LPCCharges=ROUND((DATEDIFF(day,@LPCStart,@nRK)+1.0)/NULLIF(@LPCMonth,0)*@LPCCal*(@Penalty/100.0),0);
        SET @LPCTotPenalty=@LPCTotPenalty+@LPCCharges; SET @LPCSeq=@LPCSeq+1;
        INSERT dbo.KTA_T_ProcessBilling_Calculation
        (CompanyID,LocationID,Period,UnitNo,BillingNo,IncomeCode,SeqNo,BillDate,DueDate,OwnerID,CurrencyID,
         BillAmount,CalDay,CalMonth,CalAmount,CalPersen,Charges,UpdateUser,UpdateDate)
        VALUES(@CompanyID,@LocationID,CONVERT(varchar(6),@PeriodCAR,112),@LPCUnit,@LPCDoc,@LPCIncome,@LPCSeq,
         @LPCBillDate,@LPCDueDate,@LPCOwner,@LPCCurrency,@LPCCal,DATEDIFF(day,@LPCStart,@nRK)+1,@LPCMonth,
         @LPCCal,@Penalty,@LPCCharges,@UpdateUser,GETDATE());
        SET @LPCRKDate1=@nRK;
       END

       IF @nRT<>'1' AND @nRK<=@nEnd SET @LPCOuts=@LPCOuts-@nAmt;

       FETCH NEXT FROM lpc_cur INTO @nKey,@nDoc,@nDocDate,@nDue,@nIncome,@nCur,@nUnit,@nOwner,
        @nBill,@nRT,@nRef,@nRK,@nAmt,@nStart,@nEnd,@nOffset;

       IF @@FETCH_STATUS<>0 SET @LPCBreak=9;
       ELSE IF @nKey<>@LPCKeyHD SET @LPCBreak=1;
       ELSE SET @LPCBreak=0;

       -- Legacy VB6 special case retained exactly: AIR negative balance for 202511.
       IF @LPCBreak=1 AND @LPCOuts<0 AND @LPCOuts1=0
          AND @LPCIncome='AIR' AND CONVERT(varchar(6),@LPCBillDate,112)='202511'
       BEGIN
          SET @LPCOuts1=@LPCOuts;
          SET @LPCIncome1=@LPCIncome;
       END

       IF @LPCBreak>0
       BEGIN
        IF @LPCOuts>0 AND CONVERT(varchar(6),@LPCDueDate,112)<=CONVERT(varchar(6),@LPCPeriod1,112)
        BEGIN
         SET @LPCCharges=ROUND(@LPCOuts*(DATEDIFF(day,@LPCStart,@LPCEnd)+1.0)/NULLIF(@LPCMonth,0)*(@Penalty/100.0),0);
         SET @LPCTotPenalty=@LPCTotPenalty+@LPCCharges; SET @LPCSeq=@LPCSeq+1;
         INSERT dbo.KTA_T_ProcessBilling_Calculation
         (CompanyID,LocationID,Period,UnitNo,BillingNo,IncomeCode,SeqNo,BillDate,DueDate,OwnerID,CurrencyID,
          BillAmount,CalDay,CalMonth,CalAmount,CalPersen,Charges,UpdateUser,UpdateDate)
         VALUES(@CompanyID,@LocationID,CONVERT(varchar(6),@PeriodCAR,112),@LPCUnit,@LPCDoc,@LPCIncome,@LPCSeq,
          @LPCBillDate,@LPCDueDate,@LPCOwner,@LPCCurrency,@LPCOuts,DATEDIFF(day,@LPCStart,@LPCEnd)+1,@LPCMonth,
          @LPCOuts,@Penalty,@LPCCharges,@UpdateUser,GETDATE());
        END

        IF @LPCBreak>1 AND @LPCOuts1<0 AND @LPCCharges<>0
        BEGIN
         SET @LPCCharges=ROUND(@LPCOuts1*(DATEDIFF(day,@LPCStart,@LPCRKDate1)+1.0)/NULLIF(@LPCMonth,0)*(@Penalty/100.0),0);
         SET @LPCTotPenalty=@LPCTotPenalty+@LPCCharges; SET @LPCSeq=@LPCSeq+1;
         INSERT dbo.KTA_T_ProcessBilling_Calculation
         (CompanyID,LocationID,Period,UnitNo,BillingNo,IncomeCode,SeqNo,BillDate,DueDate,OwnerID,CurrencyID,
          BillAmount,CalDay,CalMonth,CalAmount,CalPersen,Charges,UpdateUser,UpdateDate)
         VALUES(@CompanyID,@LocationID,CONVERT(varchar(6),@PeriodCAR,112),@LPCUnit,@LPCDoc,@LPCIncome1,@LPCSeq,
          @LPCBillDate,@LPCDueDate,@LPCOwner,@LPCCurrency,@LPCOuts1,DATEDIFF(day,@LPCStart,@LPCRKDate1)+1,@LPCMonth,
          @LPCOuts1,@Penalty,@LPCCharges,@UpdateUser,GETDATE());
        END

        SET @LPCOuts1=0; SET @LPCIncome1='';
        IF @LPCTotPenalty>@MinPEC
        BEGIN
         SET @Seq=@Seq+1;
         INSERT dbo.KTA_T_D_Billing(KeyHD,SeqNo,IncomeCode,Amount,InvoiceNo,InvoiceDate,PhoneNo,Description)
         VALUES(@KeyHD,@Seq,'PEC',@LPCTotPenalty,'','19000101','','');
         SET @Total=@Total+@LPCTotPenalty;
        END
       END
      END
      CLOSE lpc_cur; DEALLOCATE lpc_cur;

      UPDATE dbo.KTA_T_H_Billing
      SET NewBilling=@Total,NewBalance=Previous+@Total-Payment
      WHERE KeyHD=@KeyHD;

      SET @RowID=@RowID+1;
   END

   DROP TABLE #LPC;

   /* Outstanding + Sales Journal: exact VB6 grouping boundaries. */
   ;WITH BI AS
   (
    SELECT h.BillingNo,h.CompanyID,h.LocationID,h.UnitNo,h.OwnerID,h.BillDate,h.DueDate,h.CurrencyID,
           d.IncomeCode,SUM(d.Amount) Amount
    FROM dbo.KTA_T_H_Billing h
    JOIN dbo.KTA_T_D_Billing d ON d.KeyHD=h.KeyHD
    WHERE h.CompanyID=@CompanyID AND h.LocationID=@LocationID
      AND CONVERT(varchar(6),h.BillDate,112)=CONVERT(varchar(6),@PeriodCAR,112)
      AND h.PrintCode=0 AND d.InvoiceNo='' AND d.Amount<>0
    GROUP BY h.BillingNo,h.CompanyID,h.LocationID,h.UnitNo,h.OwnerID,h.BillDate,h.DueDate,h.CurrencyID,d.IncomeCode
   )
   INSERT dbo.KTA_T_H_Outstanding
   (CompanyID,LocationID,TxType,DocumentNo,IncomeCode,DocumentDate,DueDate,OwnerID,UnitNo,LocationNo,
    CurrencyID,PaidDate,CancelDate,Status,UpdateUser,UpdateDate)
   SELECT CompanyID,LocationID,'B',BillingNo,IncomeCode,BillDate,DueDate,OwnerID,UnitNo,'',
          CurrencyID,'19000101','19000101',0,@UpdateUser,GETDATE() FROM BI;

   ;WITH BI AS
   (
    SELECT h.BillingNo,h.BillDate,h.DueDate,h.UnitNo,h.OwnerID,h.CurrencyID,d.IncomeCode,SUM(d.Amount) Amount
    FROM dbo.KTA_T_H_Billing h JOIN dbo.KTA_T_D_Billing d ON d.KeyHD=h.KeyHD
    WHERE h.CompanyID=@CompanyID AND h.LocationID=@LocationID
      AND CONVERT(varchar(6),h.BillDate,112)=CONVERT(varchar(6),@PeriodCAR,112)
      AND h.PrintCode=0 AND d.InvoiceNo='' AND d.Amount<>0
    GROUP BY h.BillingNo,h.BillDate,h.DueDate,h.UnitNo,h.OwnerID,h.CurrencyID,d.IncomeCode
   )
   INSERT dbo.KTA_T_D_Outstanding
   (KeyHD,RecordType,RefNo,RefDate,RKDate,Amount,PaymentNo,CN,UpdateUser,UpdateDate)
   SELECT oh.KeyHD,'1',BI.BillingNo,BI.BillDate,BI.DueDate,BI.Amount,'',NULL,@UpdateUser,GETDATE()
   FROM BI JOIN dbo.KTA_T_H_Outstanding oh
     ON oh.CompanyID=@CompanyID AND oh.LocationID=@LocationID AND oh.TxType='B'
    AND oh.DocumentNo=BI.BillingNo AND oh.IncomeCode=BI.IncomeCode;

   ;WITH SJ AS
   (
    SELECT h.BillingNo,h.BillDate,h.DueDate,h.UnitNo,h.CurrencyID,SUM(d.Amount) TotalAmount
    FROM dbo.KTA_T_H_Billing h JOIN dbo.KTA_T_D_Billing d ON d.KeyHD=h.KeyHD
    WHERE h.CompanyID=@CompanyID AND h.LocationID=@LocationID
      AND CONVERT(varchar(6),h.BillDate,112)=CONVERT(varchar(6),@PeriodCAR,112)
      AND h.PrintCode=0 AND d.InvoiceNo=''
    GROUP BY h.BillingNo,h.BillDate,h.DueDate,h.UnitNo,h.CurrencyID
   )
   INSERT dbo.AR_T_H_SalesJournal
   (CompanyID,LocationID,DocumentNo,RecordType,PaymentFirst,RefNo,[Date],DueDate,CurrencyID,OriginalAmount,
    TenantID,SuiteNo,PrintCode,Status,UpdateUser,UpdateDate,Posted_to_GL,Category,ARType,VAT,PPH23)
   SELECT @CompanyID,@LocationID,BillingNo,'1',0,'',BillDate,DueDate,CurrencyID,TotalAmount,
          UnitNo,'',1,0,@UpdateUser,GETDATE(),0,'','B',0,0 FROM SJ;

   DECLARE sj_cur CURSOR LOCAL FAST_FORWARD FOR
    SELECT h.BillingNo,d.IncomeCode,SUM(d.Amount),MAX(d.PhoneNo),
           MAX(i.CrAccount),MAX(i.DeptID),MAX(i.Description),h.CurrencyID,h.UnitNo
    FROM dbo.KTA_T_H_Billing h
    JOIN dbo.KTA_T_D_Billing d ON d.KeyHD=h.KeyHD
    JOIN dbo.KTA_M_Income i ON i.CompanyID=h.CompanyID AND i.LocationID=h.LocationID AND i.IncomeCode=d.IncomeCode
    WHERE h.CompanyID=@CompanyID AND h.LocationID=@LocationID
      AND CONVERT(varchar(6),h.BillDate,112)=CONVERT(varchar(6),@PeriodCAR,112)
      AND h.PrintCode=0 AND d.InvoiceNo=''
    GROUP BY h.BillingNo,d.IncomeCode,h.CurrencyID,h.UnitNo
    ORDER BY h.BillingNo,d.IncomeCode;

   DECLARE @SBillingNo varchar(10),@SIncome varchar(3),@SAmount money,@SPhone varchar(15),
           @SCr varchar(10),@SDept varchar(10),@SDesc varchar(50),@SCur varchar(3),@SUnit varchar(5),
           @SJSKey int,@SJSeq tinyint,@MFEArea money,@MFEMFee money,@MFESF money,
           @SKFAccount varchar(10),@SKFDept varchar(10),@SKFDesc varchar(50);

   SELECT @SKFAccount=CrAccount,@SKFDept=DeptID,@SKFDesc=Description
   FROM dbo.KTA_M_Income WHERE CompanyID=@CompanyID AND LocationID=@LocationID AND IncomeCode='SKF';

   OPEN sj_cur;
   FETCH NEXT FROM sj_cur INTO @SBillingNo,@SIncome,@SAmount,@SPhone,@SCr,@SDept,@SDesc,@SCur,@SUnit;
   WHILE @@FETCH_STATUS=0
   BEGIN
    SELECT @SJSKey=KeyHD FROM dbo.AR_T_H_SalesJournal
    WHERE CompanyID=@CompanyID AND LocationID=@LocationID AND DocumentNo=@SBillingNo;
    SELECT @SJSeq=ISNULL(MAX(SeqNo),0) FROM dbo.AR_T_D_SalesJournal WHERE KeyHD=@SJSKey;

    IF @SIncome='MFE'
    BEGIN
     SELECT TOP 1 @MFEArea=r.Area,@MFEMFee=r.MaintenanceFee,@MFESF=r.SinkingFund
     FROM dbo.KTA_M_Unit u JOIN dbo.KTA_M_RoomType r
       ON r.CompanyID=u.CompanyID AND r.LocationID=u.LocationID AND r.RoomType=u.RoomType
     WHERE u.CompanyID=@CompanyID AND u.LocationID=@LocationID AND u.UnitNo=@SUnit;

     SET @SJSeq=@SJSeq+1;
     INSERT dbo.AR_T_D_SalesJournal
     (KeyHD,SeqNo,SuiteNo,IncomeID,AccountCr,ProfitCenterNo,Description1,Description2,Description3,CurrencyID,
      OriginalAmount,RpAmount,USAmount,VAT,TaxCode)
     VALUES(@SJSKey,@SJSeq,'','MFE',@SCr,@SDept,@SDesc,'','',@SCur,@MFEArea*@MFEMFee,
      CASE WHEN @SCur='IDR' THEN @MFEArea*@MFEMFee ELSE @MFEArea*@MFEMFee*@AcctRate END,
      CASE WHEN @SCur='IDR' THEN (@MFEArea*@MFEMFee)/NULLIF(@AcctRate,0) ELSE @MFEArea*@MFEMFee END,0,NULL);

     SET @SJSeq=@SJSeq+1;
     INSERT dbo.AR_T_D_SalesJournal
     (KeyHD,SeqNo,SuiteNo,IncomeID,AccountCr,ProfitCenterNo,Description1,Description2,Description3,CurrencyID,
      OriginalAmount,RpAmount,USAmount,VAT,TaxCode)
     VALUES(@SJSKey,@SJSeq,'','SKF',@SKFAccount,@SKFDept,@SKFDesc,'','',@SCur,@MFEArea*@MFESF,
      CASE WHEN @SCur='IDR' THEN @MFEArea*@MFESF ELSE @MFEArea*@MFESF*@AcctRate END,
      CASE WHEN @SCur='IDR' THEN (@MFEArea*@MFESF)/NULLIF(@AcctRate,0) ELSE @MFEArea*@MFESF END,0,NULL);
    END
    ELSE
    BEGIN
     SET @SJSeq=@SJSeq+1;
     INSERT dbo.AR_T_D_SalesJournal
     (KeyHD,SeqNo,SuiteNo,IncomeID,AccountCr,ProfitCenterNo,Description1,Description2,Description3,CurrencyID,
      OriginalAmount,RpAmount,USAmount,VAT,TaxCode)
     VALUES(@SJSKey,@SJSeq,'',@SIncome,@SCr,@SDept,@SDesc,
      CASE WHEN @SIncome='TEC' THEN 'Phone No. '+ISNULL(@SPhone,'') ELSE '' END,'',@SCur,@SAmount,
      CASE WHEN @SCur='IDR' THEN @SAmount ELSE @SAmount*@AcctRate END,
      CASE WHEN @SCur='IDR' THEN @SAmount/NULLIF(@AcctRate,0) ELSE @SAmount END,0,NULL);
    END
    FETCH NEXT FROM sj_cur INTO @SBillingNo,@SIncome,@SAmount,@SPhone,@SCr,@SDept,@SDesc,@SCur,@SUnit;
   END
   CLOSE sj_cur; DEALLOCATE sj_cur;

   UPDATE sj SET OriginalAmount=x.TotalAmount
   FROM dbo.AR_T_H_SalesJournal sj
   JOIN (
    SELECT h.BillingNo,SUM(d.Amount) TotalAmount
    FROM dbo.KTA_T_H_Billing h JOIN dbo.KTA_T_D_Billing d ON d.KeyHD=h.KeyHD
    WHERE h.CompanyID=@CompanyID AND h.LocationID=@LocationID
      AND CONVERT(varchar(6),h.BillDate,112)=CONVERT(varchar(6),@PeriodCAR,112)
      AND h.PrintCode=0 AND d.InvoiceNo=''
    GROUP BY h.BillingNo
   ) x ON x.BillingNo=sj.DocumentNo
   WHERE sj.CompanyID=@CompanyID AND sj.LocationID=@LocationID;

   INSERT dbo.KTA_T_ProcessBilling
   (CompanyID,LocationID,Period,Status,UpdateUser,UpdateDate)
   VALUES(@CompanyID,@LocationID,CONVERT(varchar(6),@PeriodCAR,112),1,@UpdateUser,GETDATE());

   DELETE dbo.KTA_T_BillingPhone
   WHERE CONVERT(varchar(8),BillDate,112)=
     CONVERT(varchar(8),DATEADD(day,-1,DATEADD(month,-1,DATEADD(day,1,@Period))),112);

   COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
   IF @@TRANCOUNT>0 ROLLBACK TRANSACTION;
   SET @Err=ERROR_MESSAGE();
   RAISERROR(@Err,16,1);
 END CATCH
END
GO
