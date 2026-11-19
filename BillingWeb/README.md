# KTA Web Billing module

Target: ASP.NET WebForms / .NET Framework 4.5 / C# / SQL Server 2008.
The module is intentionally standalone so the page/code/database scripts can be
copied into the actual KTA solution later.

## Web UI

Billing.aspx provides the complete upload/validation/preview/process workflow:

1. Start a new upload session; its ID exists only in ASP.NET Session/ViewState and is never stored in SQL.
2. Upload mandatory Billing_member.csv.
3. Upload optional Card Payment Statistics.xlsx.
4. Upload optional Billing Report Monthly.csv/.tsv.
5. If an optional source is unavailable, explicitly confirm that source is unavailable.
6. Run validation against the active KAR period.
7. Preview source control totals without creating billing numbers or posting billing.
8. Process only after validation succeeds; posting validates the same TVPs again inside SQL.
9. After successful posting, the same CAR period cannot be processed again.

The production copy should map ResolveContext() to the real KTA session conventions
for CompanyID, LocationID, and UserID. The standalone sample uses configurable
session-key names in Web.config.

## TVP architecture

There are no persistent billing staging or batch tables. The three source files are parsed into in-memory .NET DataTables held in server-side ASP.NET Session for the lifetime of the upload session. Each Validate, Preview, Process, and Dummy Statement call sends the rows to SQL Server using these table-valued parameters:

- `dbo.Type_Tbl_C4` for Billing_member (4 positional columns)
- `dbo.Type_Tbl_C7` for Card Payment Statistics (7 positional columns)
- `dbo.Type_Tbl_C7` for Billing Report Monthly (5 populated columns; remaining columns unused)

SQL validates the TVPs against the active KAR period and existing KTA master/transaction data. Preview reads the TVPs without posting. Final posting repeats validation and copies TVP rows into connection-local `#temp` tables for the legacy calculation flow; those temporary tables disappear automatically. Only existing KTA billing, outstanding, sales-journal, and process tables receive permanent transaction data.

Optional-source unavailable confirmations are held in ASP.NET Session and passed as bit parameters. If an uploaded source is invalid, validation fails; it is never silently treated as unavailable. Since upload data is session-scoped, users must keep the same web session until Process completes; an expired session requires re-upload.

## Database deployment order

1. Database/01_Billing_Staging_Validation_Preview.sql
2. Database/02_usp_KTA_Billing_Post.sql
3. Existing vb6/sp_KTA_rpt_BillingStatement.sql (deploy/update this legacy report procedure before the dummy wrapper)
4. Database/03_usp_KTA_Billing_DummyStatement.sql
5. Database/04_Verify_Billing_Deployment.sql (read-only object/schema preflight)

Script 01 reuses the existing shared TVP types (`dbo.Type_Tbl_C4` and `dbo.Type_Tbl_C7`) and creates only the Validate/Preview procedures. It also removes the old KTA_W_* staging tables and CreateBatch/ValidateBatch procedures from the previous design. No physical staging tables are needed.

## Billing period

The authoritative billing period is always:

SELECT ActivePeriod FROM dbo.M_Active_Period WHERE AutoNum = 'KAR' + @CompanyID + @LocationID;

The CAR period is read separately from:

SELECT ActivePeriod FROM dbo.M_Active_Period WHERE AutoNum = 'KCAR' + @CompanyID + @LocationID;

The browser never supplies the period used for validation/posting.

## Billing number generation

Posting uses the INSPEC transaction pattern, with discriminator arkta and KBI prefix:

DECLARE @AutoNumResult TABLE (ID VARCHAR(10), SeqNo INT);
DECLARE @AutoNum VARCHAR(20);
SET @AutoNum = 'KBI' + @CompanyID + @LocationID;
INSERT INTO @AutoNumResult EXECUTE dbo.sp_GetGenerateAutoNumID 'arkta', @AutoNum, 1;
SELECT @BillingNo = ID FROM @AutoNumResult;

The call is made only during final posting and inside the same transaction as billing,
outstanding, sales-journal, and process-period writes.

## Legacy calculation coverage

02_usp_KTA_Billing_Post.sql ports the main VB6 process in SQL:

- maintenance fee / sinking fund and hard-copy charge
- prior work-order outstanding
- collect call
- telephone detail and PPN = 11% hardcoded for now
- Billing_member / PCR
- EV-1 monthly reimbursement
- EV-2 card reimbursement
- ELC/AIR handover reimbursement
- LPC stateful calculation ported from sp_KTA_selectLPC; source rows are now loaded once for all eligible units (instead of executing the legacy procedure once per unit)
- PEC minimum-charge handling
- legacy AIR negative-balance special case
- billing header/detail totals and balances
- grouped KTA outstanding header/detail
- sales journal header/detail with MFE/SKF split
- KTA_T_ProcessBilling completion
- cleanup of the previous-period telephone source rows

All final writes are wrapped in one transaction with rollback on error.

## SQL Server 2008 compatibility

Scripts avoid SQL Server 2012+ features such as DATEFROMPARTS, TRY_CONVERT,
STRING_SPLIT, THROW, CREATE OR ALTER, and FORMAT.

## Excel dependency

The standalone WebForms project uses EPPlus 4.5.3.2 for reading the Card Payment
Statistics workbook. That legacy version targets .NET Framework 4.5 and is kept here
because the project target is VS2013/.NET 4.5.

## Dummy Billing Statement / Crystal Report

The Billing page includes a dummy statement action for a selected unit range. It:
- runs the same billing calculation/posting procedure in dummy mode;
- uses temporary D-prefixed BillingNo values and never calls sp_GetGenerateAutoNumID;
- runs inside an outer SQL transaction and rolls back after the legacy
  sp_KTA_rpt_BillingStatement dataset has been returned;
- uses the copied BillingWeb/Reports/rptKTA_BillingStatement_New.rpt file to export an inline PDF
  when the SAP Crystal Reports runtime is installed and registered in IIS;
- falls back to displaying the returned report dataset in GridView if Crystal runtime
  or report binding is unavailable.

The legacy report SP currently filters units to KTA_M_Unit.SendEmail='Y'. The dummy
screen uses its unit range filter and therefore follows that same legacy rule.

The dummy path temporarily writes billing/outstanding/journal rows inside a transaction
before rolling them back. It does not consume the production KBI autonumber. Before
enabling it on production IIS, verify that billing-table triggers do not perform
non-transactional external side effects. The Crystal PDF export also needs runtime
verification against the installed Crystal version and the uploaded .rpt file.

The XLSX preview workbook remains a later presentation enhancement; it is separate
from the billing calculation/posting logic.
