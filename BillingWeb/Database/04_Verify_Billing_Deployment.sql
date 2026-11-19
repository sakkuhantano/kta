/*
 KTA Web Billing deployment preflight.
 Run after the legacy KTA schema, sp_KTA_rpt_BillingStatement,
 and BillingWeb/Database/01-03 scripts have been deployed.
 SQL Server 2008 compatible. This script is read-only.
*/
SET NOCOUNT ON;

DECLARE @Missing varchar(2000);
SET @Missing='';

IF OBJECT_ID('dbo.M_Active_Period','U') IS NULL SET @Missing=@Missing+'dbo.M_Active_Period; ';
IF OBJECT_ID('dbo.KTA_M_Penalty','U') IS NULL SET @Missing=@Missing+'dbo.KTA_M_Penalty; ';
IF OBJECT_ID('dbo.M_Rate','U') IS NULL SET @Missing=@Missing+'dbo.M_Rate; ';
IF OBJECT_ID('dbo.KTA_M_Unit','U') IS NULL SET @Missing=@Missing+'dbo.KTA_M_Unit; ';
IF OBJECT_ID('dbo.KTA_M_RoomType','U') IS NULL SET @Missing=@Missing+'dbo.KTA_M_RoomType; ';
IF OBJECT_ID('dbo.KTA_M_Owner','U') IS NULL SET @Missing=@Missing+'dbo.KTA_M_Owner; ';
IF OBJECT_ID('dbo.KTA_M_Income','U') IS NULL SET @Missing=@Missing+'dbo.KTA_M_Income; ';
IF OBJECT_ID('dbo.KTA_T_BillingPhone','U') IS NULL SET @Missing=@Missing+'dbo.KTA_T_BillingPhone; ';
IF OBJECT_ID('dbo.KTA_T_H_Billing','U') IS NULL SET @Missing=@Missing+'dbo.KTA_T_H_Billing; ';
IF OBJECT_ID('dbo.KTA_T_D_Billing','U') IS NULL SET @Missing=@Missing+'dbo.KTA_T_D_Billing; ';
IF OBJECT_ID('dbo.KTA_T_D_Billing_Phone','U') IS NULL SET @Missing=@Missing+'dbo.KTA_T_D_Billing_Phone; ';
IF OBJECT_ID('dbo.KTA_T_H_Outstanding','U') IS NULL SET @Missing=@Missing+'dbo.KTA_T_H_Outstanding; ';
IF OBJECT_ID('dbo.KTA_T_D_Outstanding','U') IS NULL SET @Missing=@Missing+'dbo.KTA_T_D_Outstanding; ';
IF OBJECT_ID('dbo.vw_KTA_T_H_Outstanding','V') IS NULL SET @Missing=@Missing+'dbo.vw_KTA_T_H_Outstanding; ';
IF OBJECT_ID('dbo.KTA_T_CollectCall','U') IS NULL SET @Missing=@Missing+'dbo.KTA_T_CollectCall; ';
IF OBJECT_ID('dbo.AR_T_H_Billing','U') IS NULL SET @Missing=@Missing+'dbo.AR_T_H_Billing; ';
IF OBJECT_ID('dbo.AR_T_D_Billing','U') IS NULL SET @Missing=@Missing+'dbo.AR_T_D_Billing; ';
IF OBJECT_ID('dbo.AR_T_H_SalesJournal','U') IS NULL SET @Missing=@Missing+'dbo.AR_T_H_SalesJournal; ';
IF OBJECT_ID('dbo.AR_T_D_SalesJournal','U') IS NULL SET @Missing=@Missing+'dbo.AR_T_D_SalesJournal; ';
IF OBJECT_ID('dbo.KTA_T_ProcessBilling','U') IS NULL SET @Missing=@Missing+'dbo.KTA_T_ProcessBilling; ';
IF OBJECT_ID('dbo.KTA_T_ProcessBilling_Calculation','U') IS NULL SET @Missing=@Missing+'dbo.KTA_T_ProcessBilling_Calculation; ';
IF OBJECT_ID('dbo.KTA_M_BillDes','U') IS NULL SET @Missing=@Missing+'dbo.KTA_M_BillDes; ';
IF OBJECT_ID('dbo.KTA_M_BillBank','U') IS NULL SET @Missing=@Missing+'dbo.KTA_M_BillBank; ';
IF OBJECT_ID('dbo.sp_GetGenerateAutoNumID','P') IS NULL SET @Missing=@Missing+'dbo.sp_GetGenerateAutoNumID; ';
IF OBJECT_ID('dbo.sp_KTA_rpt_BillingStatement','P') IS NULL SET @Missing=@Missing+'dbo.sp_KTA_rpt_BillingStatement; ';
IF OBJECT_ID('dbo.sp_KTA_selectLPC','P') IS NULL SET @Missing=@Missing+'dbo.sp_KTA_selectLPC; ';

IF @Missing<>''
BEGIN
 RAISERROR('Billing deployment is missing required KTA objects: %s',16,1,@Missing);
 RETURN;
END

IF COL_LENGTH('dbo.KTA_M_Penalty','BillDate') IS NULL
 OR COL_LENGTH('dbo.KTA_M_Penalty','StartingNewBilling') IS NULL
 OR COL_LENGTH('dbo.KTA_M_Penalty','PreviousBillDate') IS NULL
BEGIN
 RAISERROR('KTA_M_Penalty is missing billing/LPC date configuration columns.',16,1);
 RETURN;
END

IF COL_LENGTH('dbo.KTA_T_BillingPhone','BillDate') IS NULL
 OR COL_LENGTH('dbo.KTA_T_BillingPhone','PhoneNo') IS NULL
 OR COL_LENGTH('dbo.KTA_T_BillingPhone','Fee') IS NULL
 OR COL_LENGTH('dbo.KTA_T_BillingPhone','Other') IS NULL
BEGIN
 RAISERROR('KTA_T_BillingPhone does not match the expected legacy telephone source schema.',16,1);
 RETURN;
END

IF TYPE_ID('dbo.Type_Tbl_C4') IS NULL
 OR TYPE_ID('dbo.Type_Tbl_C7') IS NULL
 OR TYPE_ID('dbo.Type_Tbl_C7') IS NULL
 OR OBJECT_ID('dbo.usp_KTA_Billing_Validate','P') IS NULL
 OR OBJECT_ID('dbo.usp_KTA_Billing_Preview','P') IS NULL
 OR OBJECT_ID('dbo.usp_KTA_Billing_Post','P') IS NULL
 OR OBJECT_ID('dbo.usp_KTA_Billing_DummyStatement','P') IS NULL
BEGIN
 RAISERROR('Web Billing TVP/procedure deployment is incomplete. Run scripts 01, 02, and 03.',16,1);
 RETURN;
END

IF OBJECT_ID('dbo.KTA_W_BillingBatch','U') IS NOT NULL
 OR OBJECT_ID('dbo.KTA_W_BillingMember','U') IS NOT NULL
 OR OBJECT_ID('dbo.KTA_W_BillingCard','U') IS NOT NULL
 OR OBJECT_ID('dbo.KTA_W_BillingMonthly','U') IS NOT NULL
BEGIN
 RAISERROR('Legacy physical billing staging tables still exist. Review/drop them after confirming they are no longer needed.',16,1);
 RETURN;
END

SELECT
 CAST(SERVERPROPERTY('ProductVersion') AS varchar(30)) AS SqlServerVersion,
 'OK' AS DeploymentPreflight,
 'All required legacy objects and Web Billing objects are present.' AS Message;
GO
