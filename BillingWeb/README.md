# KTA Web Billing module

Target: ASP.NET WebForms / .NET Framework 4.5 / C# / SQL Server 2008.
Reference UI conventions: INSPEC `Secure/M_InspectionType.aspx` and
`MasterPages/Site_BS_5.Master`.

## Current database layer

Apply `Database/01_Billing_Staging_Validation_Preview.sql` first. It creates
BatchID-scoped staging for the three upload sources, reads the authoritative
active period from `M_Active_Period` using `KAR + CompanyID + LocationID`,
validates source periods, and returns source summary/detail/status result sets.

The two EV files are optional individually, but absence must be explicitly
confirmed. An uploaded invalid file is not treated as absent. Billing_member
is mandatory. The active period is never taken from a browser field or file.

## Billing number generation

The posting procedure must use the INSPEC transaction pattern exactly, changing
the database discriminator to `arkta` and the prefix to KBI:

```sql
DECLARE @AutoNumResult TABLE (ID VARCHAR(10), SeqNo INT);
DECLARE @AutoNum VARCHAR(20);
SET @AutoNum = 'KBI' + @CompanyID + @LocationID;

INSERT INTO @AutoNumResult
EXECUTE dbo.sp_GetGenerateAutoNumID 'arkta', @AutoNum, 1;

SELECT @BillingNo = ID FROM @AutoNumResult;
```

This call belongs inside the same SQL transaction as the billing header/detail,
outstanding, sales journal, and process-period writes. Do not generate billing
numbers for preview.

## Posting boundary

Preview is non-posting: it must not create billing headers/details, outstanding,
sales journal rows, process-period completion, or consume billing numbers.
The final posting procedure still needs the complete legacy calculation port,
including LPC sequence/penalty behavior, telephone detail, ELC/AIR
reimbursements, and grouped outstanding/sales-journal creation. Those writes
must commit or roll back together.

## SQL compatibility

All scripts must remain compatible with SQL Server 2008: no TRY_CONVERT,
STRING_SPLIT, THROW, CREATE OR ALTER, or FORMAT.
