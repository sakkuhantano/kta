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
 @UpdateUser varchar(10)
AS
BEGIN
 SET NOCOUNT ON;
 SET XACT_ABORT ON;

 DECLARE @Period smalldatetime,@PeriodCAR smalldatetime,@PrevPeriod smalldatetime,
         @BillDate smalldatetime,@DueDate smalldatetime,@TermDue int,
         @Penalty money,@MinPEC money,@HardCopyCharges money,@PPN money,
         @AcctRate money,@ProcessStatus tinyint,@Err varchar(4000);

 SELECT @Period=ActivePeriod FROM dbo.KTA_W_BillingBatch
 WHERE BatchID=@BatchID AND CompanyID=@CompanyID AND LocationID=@LocationID;
 IF @Period IS NULL RAISERROR('Invalid billing batch/company/location.',16,1);

 SELECT @PeriodCAR=ActivePeriod FROM dbo.M_Active_Period
 WHERE AutoNum='KCAR'+@CompanyID+@LocationID;
 IF @PeriodCAR IS NULL RAISERROR('Active CAR period was not found.',16,1);

 SELECT @TermDue=CAST(Term AS int),@Penalty=Penalty,@MinPEC=ISNULL(MinPEC,0),
        @HardCopyCharges=ISNULL(HardCopyCharges,0)
 FROM dbo.KTA_M_Penalty WHERE CompanyID=@CompanyID AND LocationID=@LocationID;
 IF @TermDue IS NULL RAISERROR('KTA_M_Penalty not found.',16,1);

 SELECT TOP 1 @AcctRate=AcctRate FROM dbo.M_Rate
 WHERE CONVERT(varchar(6),[Date],112)=CONVERT(varchar(6),@PeriodCAR,112)
 ORDER BY [Date];
 IF @AcctRate IS NULL RAISERROR('M_Rate for active CAR period was not found.',16,1);

 IF NOT EXISTS(SELECT 1 FROM dbo.KTA_W_BillingBatch WHERE BatchID=@BatchID AND ValidationStatus=1)
  RAISERROR('Billing batch has not passed validation.',16,1);

 SET @BillDate=DATEADD(day,@TermDue*0,@Period);
 SELECT @BillDate=DATEADD(day,0,DATEADD(day,
       DAY(@BillDate)*0 + ISNULL((SELECT BillDate FROM dbo.KTA_M_Penalty WHERE CompanyID=@CompanyID AND LocationID=@LocationID),0)-1,
       DATEADD(month,DATEDIFF(month,0,@Period),0)));
 IF @BillDate IS NULL OR DAY(@BillDate)=0 SET @BillDate=@Period;
 SET @DueDate=DATEADD(day,@TermDue,@BillDate)-1;
 SET @PrevPeriod=DATEADD(day,-1,DATEADD(month,DATEDIFF(month,0,@Period),0));

 BEGIN TRANSACTION;
 BEGIN TRY

   IF EXISTS(SELECT 1 FROM dbo.KTA_T_ProcessBilling WHERE CompanyID=@CompanyID
             AND LocationID=@LocationID AND Period=CONVERT(varchar(6),@PeriodCAR,112))
      RAISERROR('Billing period has already been processed.',16,1);

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
        AND d.RecordType IN ('3','4','5') AND d.RefDate<=h.EndDate
      GROUP BY h.CompanyID,h.LocationID,h.UnitNo
   ) p ON p.CompanyID=u.CompanyID AND p.LocationID=u.LocationID AND p.UnitNo=u.UnitNo
   WHERE u.CompanyID=@CompanyID AND u.LocationID=@LocationID
     AND u.Tower<>'' AND u.OwnerID<>''
     AND EXISTS(SELECT 1 FROM dbo.KTA_T_H_Billing x
                WHERE x.CompanyID=u.CompanyID AND x.LocationID=u.LocationID
                  AND x.UnitNo=u.UnitNo
                  AND CONVERT(varchar(6),x.BillDate,112)=CONVERT(varchar(6),@PrevPeriod,112));

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

      DECLARE @AutoNumResult TABLE(ID varchar(10),SeqNo int);
      DECLARE @AutoNum varchar(20);
      SET @AutoNum='KBI'+@CompanyID+@LocationID;
      INSERT INTO @AutoNumResult
      EXECUTE dbo.sp_GetGenerateAutoNumID 'arkta',@AutoNum,1;
      SELECT @BillingNo=ID FROM @AutoNumResult;
      IF ISNULL(@BillingNo,'')='' RAISERROR('Failed to generate BillingNo.',16,1);

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
      SELECT @KeyHD,@Seq+ROW_NUMBER() OVER(ORDER BY d.RefDate,d.RefNo),'',
             d.Amount,d.RefNo,d.RefDate,'',NULL
      FROM dbo.KTA_T_H_Outstanding h
      INNER JOIN dbo.KTA_T_D_Outstanding d ON d.KeyHD=h.KeyHD
      WHERE h.CompanyID=@CompanyID AND h.LocationID=@LocationID AND h.UnitNo=@UnitNo
        AND h.TxType='W' AND h.LocationNo=''
        AND d.RecordType='1'
        AND CONVERT(varchar(6),h.DocumentDate,112)=CONVERT(varchar(6),@PrevPeriod,112);
      -- Normalize the income code after bulk insert by joining source rows.
      UPDATE b SET IncomeCode=d.IncomeCode
      FROM dbo.KTA_T_D_Billing b
      INNER JOIN dbo.KTA_T_D_Outstanding d ON b.InvoiceNo=d.RefNo
      INNER JOIN dbo.KTA_T_H_Outstanding h ON h.KeyHD=d.KeyHD
      WHERE b.KeyHD=@KeyHD AND h.TxType='W' AND d.RecordType='1';
      SELECT @Seq=ISNULL(MAX(SeqNo),@Seq),@Total=@Total+ISNULL(SUM(Amount),0)
      FROM dbo.KTA_T_D_Billing WHERE KeyHD=@KeyHD;

      INSERT dbo.KTA_T_D_Billing(KeyHD,SeqNo,IncomeCode,Amount,InvoiceNo,InvoiceDate,PhoneNo,Description)
      SELECT @KeyHD,@Seq+ROW_NUMBER() OVER(ORDER BY Period,Description),
             IncomeCode,Amount,'','19000101','',Description
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
      SELECT @KeyHD,@Seq+ROW_NUMBER() OVER(ORDER BY RowNo),'PCR',Amount,'','19000101','',LEFT([Description],100)
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

      UPDATE dbo.KTA_T_H_Billing
      SET NewBilling=@Total,NewBalance=Previous+@Total-Payment
      WHERE KeyHD=@KeyHD;

      SET @RowID=@RowID+1;
   END

   -- Outstanding and Sales Journal are intentionally left for the next
   -- verified porting stage because their grouping boundary is billing/income
   -- dependent and must exactly match the VB6 c_Break behavior.
   RAISERROR('Posting stage incomplete: outstanding/sales-journal port must be completed before production use.',16,1);

   COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
   IF @@TRANCOUNT>0 ROLLBACK TRANSACTION;
   SET @Err=ERROR_MESSAGE();
   RAISERROR(@Err,16,1);
 END CATCH
END
GO
