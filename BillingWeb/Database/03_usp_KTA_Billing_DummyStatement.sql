/*
 KTA Web Billing - non-posting dummy statement preview.
 SQL Server 2008 compatible.
 The billing rows are created only inside this outer transaction and rolled back
 after the legacy Crystal report dataset has been returned.
 Dummy mode never calls sp_GetGenerateAutoNumID.
*/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

IF OBJECT_ID('dbo.usp_KTA_Billing_DummyStatement','P') IS NOT NULL
 DROP PROCEDURE dbo.usp_KTA_Billing_DummyStatement
GO
CREATE PROCEDURE dbo.usp_KTA_Billing_DummyStatement
 @BatchID uniqueidentifier,
 @CompanyID varchar(10),
 @LocationID varchar(3),
 @UpdateUser varchar(10),
 @UnitNoFrom varchar(10)='',
 @UnitNoThru varchar(10)='',
 @GroupIDFrom varchar(10)='',
 @GroupIDThru varchar(10)='',
 @ByWhat varchar(1)='1'
AS
BEGIN
 SET NOCOUNT ON;
 SET XACT_ABORT ON;

 DECLARE @Period smalldatetime;
 SELECT @Period=ActivePeriod
 FROM dbo.KTA_W_BillingBatch
 WHERE BatchID=@BatchID AND CompanyID=@CompanyID AND LocationID=@LocationID
   AND ValidationStatus=1;
 IF @Period IS NULL
 BEGIN
  RAISERROR('Dummy statement requires a validated batch for this company/location.',16,1);
  RETURN;
 END

 IF @ByWhat NOT IN ('1','2','')
 BEGIN
  RAISERROR('ByWhat must be 1 (unit), 2 (group), or blank (all).',16,1);
  RETURN;
 END

 BEGIN TRANSACTION;
 BEGIN TRY
  /*
    The posting procedure's transaction is nested under this transaction.
    Its COMMIT only decrements @@TRANCOUNT; this wrapper owns the final ROLLBACK.
    Dummy=1 uses D000000001... IDs and does not touch the real billing autonumber.
  */
  EXEC dbo.usp_KTA_Billing_Post
       @BatchID=@BatchID,
       @CompanyID=@CompanyID,
       @LocationID=@LocationID,
       @UpdateUser=@UpdateUser,
       @Dummy=1;

  /*
    Reuse the exact legacy report dataset and report filters/layout contract.
    Result rows are streamed to the caller before the outer transaction is rolled back.
  */
  EXEC dbo.sp_KTA_rpt_BillingStatement
       @CompanyID=@CompanyID,
       @LocationID=@LocationID,
       @UnitNoFrom=@UnitNoFrom,
       @UnitNoThru=@UnitNoThru,
       @GroupIDFrom=@GroupIDFrom,
       @GroupIDThru=@GroupIDThru,
       @Period=CONVERT(varchar(6),@Period,112),
       @ByWhat=@ByWhat;

  ROLLBACK TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @@TRANCOUNT>0 ROLLBACK TRANSACTION;
  DECLARE @Err varchar(4000);
  SET @Err=ERROR_MESSAGE();
  RAISERROR(@Err,16,1);
 END CATCH
END
GO
