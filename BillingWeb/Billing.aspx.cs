using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Text;
using System.Web.UI;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace BillingWeb
{
    public partial class Billing : Page
    {
        private const string BatchViewStateKey = "KTA_Billing_SessionID";
        private string MemberKey { get { return "KTA_Billing_Member_" + BatchID; } }
        private string CardKey { get { return "KTA_Billing_Card_" + BatchID; } }
        private string MonthlyKey { get { return "KTA_Billing_Monthly_" + BatchID; } }
        private string CardNAKey { get { return "KTA_Billing_CardNA_" + BatchID; } }
        private string MonthlyNAKey { get { return "KTA_Billing_MonthlyNA_" + BatchID; } }
        private DataTable MemberData { get { return Session[MemberKey] as DataTable ?? EmptyMember(); } }
        private DataTable CardData { get { return Session[CardKey] as DataTable ?? EmptyCard(); } }
        private DataTable MonthlyData { get { return Session[MonthlyKey] as DataTable ?? EmptyMonthly(); } }
        private string ConnectionString { get { return ConfigurationManager.ConnectionStrings["KTA"].ConnectionString; } }

        private Guid BatchID
        {
            get
            {
                object v = ViewState[BatchViewStateKey];
                Guid id;
                return v != null && Guid.TryParse(v.ToString(), out id) ? id : Guid.Empty;
            }
            set { ViewState[BatchViewStateKey] = value.ToString(); }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                LoadContext();
                CreateBatch();
            }
        }

        private void LoadContext()
        {
            txtCompany.Text = ResolveContext("CompanyID", "KTA.CompanyID");
            txtLocation.Text = ResolveContext("LocationID", "KTA.LocationID");
            if (String.IsNullOrWhiteSpace(txtCompany.Text) || String.IsNullOrWhiteSpace(txtLocation.Text))
                throw new ApplicationException("CompanyID / LocationID belum tersedia pada Session. Atur appSettings KTA.CompanyID/KTA.LocationID atau sesuaikan ResolveContext() dengan session KTA.");
        }

        private string ResolveContext(string sessionKey, string appKey)
        {
            string configured = ConfigurationManager.AppSettings[appKey];
            if (!String.IsNullOrWhiteSpace(configured))
            {
                object v = Session[configured];
                if (v != null && !String.IsNullOrWhiteSpace(v.ToString())) return v.ToString();
            }
            object direct = Session[sessionKey];
            return direct == null ? "" : direct.ToString();
        }

        private string ResolveUser()
        {
            string configured = ConfigurationManager.AppSettings["KTA.UserIDSessionKey"];
            object v = !String.IsNullOrWhiteSpace(configured) ? Session[configured] : Session["UserID"];
            return v == null ? "WEB" : v.ToString();
        }

        private void CreateBatch()
        {
            BatchID = Guid.NewGuid(); // Session-scoped upload correlation only; never stored in SQL.
            Session[MemberKey] = EmptyMember();
            Session[CardKey] = EmptyCard();
            Session[MonthlyKey] = EmptyMonthly();
            Session[CardNAKey] = false;
            Session[MonthlyNAKey] = false;
            txtBatchID.Text = "Session " + BatchID.ToString("N").Substring(0, 12);
            using (SqlConnection cn = new SqlConnection(ConnectionString))
            using (SqlCommand cmd = new SqlCommand("SELECT ActivePeriod FROM dbo.M_Active_Period WHERE AutoNum='KAR'+@CompanyID+@LocationID", cn))
            {
                cmd.Parameters.Add("@CompanyID", SqlDbType.VarChar, 10).Value = txtCompany.Text.Trim();
                cmd.Parameters.Add("@LocationID", SqlDbType.VarChar, 3).Value = txtLocation.Text.Trim();
                cn.Open();
                object p = cmd.ExecuteScalar();
                if (p == null || p == DBNull.Value) throw new ApplicationException("Active KAR period tidak ditemukan.");
                txtActivePeriod.Text = Convert.ToDateTime(p).ToString("yyyy-MM");
            }
            lblStatus.Text = "Sesi upload baru siap. Upload Billing_member dan file optional.";
            lblStatus.CssClass = "alert alert-info d-block mb-0";
            SetProcessState(false);
        }

        protected void btnNew_Click(object sender, EventArgs e) { CreateBatch(); ClearPreview(); }

        protected void btnUploadMember_Click(object sender, EventArgs e)
        {
            UploadCsvMember();
        }

        protected void btnUploadMonthly_Click(object sender, EventArgs e)
        {
            UploadCsvMonthly();
        }

        protected void btnUploadCard_Click(object sender, EventArgs e)
        {
            UploadCardXlsx();
        }

        protected void chkCardNA_CheckedChanged(object sender, EventArgs e) { SetOptionalUnavailable("Card", chkCardNA.Checked); ValidateAndRefresh(); }
        protected void chkMonthlyNA_CheckedChanged(object sender, EventArgs e) { SetOptionalUnavailable("Monthly", chkMonthlyNA.Checked); ValidateAndRefresh(); }

        protected void btnValidate_Click(object sender, EventArgs e) { ValidateAndRefresh(); }

        protected void btnPreview_Click(object sender, EventArgs e)
        {
            ValidateAndRefresh();
            if (!CanPreview()) return;
            LoadPreview();
        }

        protected void btnDummyStatement_Click(object sender, EventArgs e)
        {
            try
            {
                ValidateAndRefresh();
                if (!CanPreview()) return;

                string unitFrom = (txtDummyUnitFrom.Text ?? "").Trim();
                string unitThru = (txtDummyUnitThru.Text ?? "").Trim();
                if (String.IsNullOrWhiteSpace(unitFrom) || String.IsNullOrWhiteSpace(unitThru))
                    throw new ApplicationException("Isi Unit From dan Unit Thru. Untuk satu unit, isi nomor unit yang sama di kedua field.");
                if (String.Compare(unitFrom, unitThru, StringComparison.OrdinalIgnoreCase) > 0)
                    throw new ApplicationException("Unit From tidak boleh lebih besar dari Unit Thru.");

                DataSet ds = new DataSet();
                using (SqlConnection cn = new SqlConnection(ConnectionString))
                using (SqlCommand cmd = new SqlCommand("dbo.usp_KTA_Billing_DummyStatement", cn))
                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 0;
                    AddTvpParameters(cmd);
                    cmd.Parameters.Add("@UpdateUser", SqlDbType.VarChar, 10).Value = SafeUser();
                    cmd.Parameters.Add("@UnitNoFrom", SqlDbType.VarChar, 10).Value = unitFrom;
                    cmd.Parameters.Add("@UnitNoThru", SqlDbType.VarChar, 10).Value = unitThru;
                    cmd.Parameters.Add("@GroupIDFrom", SqlDbType.VarChar, 10).Value = "";
                    cmd.Parameters.Add("@GroupIDThru", SqlDbType.VarChar, 10).Value = "";
                    cmd.Parameters.Add("@ByWhat", SqlDbType.VarChar, 1).Value = "1";
                    cn.Open();
                    da.Fill(ds);
                }

                if (ds.Tables.Count == 0 || ds.Tables[0].Rows.Count == 0)
                {
                    gvDummyStatement.DataSource = null;
                    gvDummyStatement.DataBind();
                    lblDummyStatement.Text = "Tidak ada statement untuk rentang unit tersebut. Periksa SendEmail='Y' dan data billing.";
                    lblDummyStatement.CssClass = "text-warning small d-block mt-2";
                    return;
                }

                ds.DataSetName = "sp_KTA_rpt_BillingStatement";
                ds.Tables[0].TableName = "sp_KTA_rpt_BillingStatement";
                try
                {
                    byte[] pdf = ExportCrystalStatement(ds, unitFrom, unitThru);
                    Response.Clear();
                    Response.ContentType = "application/pdf";
                    Response.AddHeader("Content-Disposition", "inline; filename=BillingStatement-DUMMY-" +
                        Regex.Replace(unitFrom, "[^A-Za-z0-9_-]", "") + ".pdf");
                    Response.BinaryWrite(pdf);
                    Response.Flush();
                    Response.End();
                    return;
                }
                catch (Exception crystalEx)
                {
                    // Keep the generated report dataset visible when the SAP Crystal runtime
                    // or a compatible report binding is not installed on the target IIS server.
                    gvDummyStatement.DataSource = ds.Tables[0];
                    gvDummyStatement.DataBind();
                    lblDummyStatement.Text = "Dataset dummy berhasil dibuat (" + ds.Tables[0].Rows.Count +
                        " baris) dan transaksi sudah di-rollback. PDF Crystal belum dapat dirender: " +
                        Server.HtmlEncode(crystalEx.GetBaseException().Message);
                    lblDummyStatement.CssClass = "text-warning small d-block mt-2";
                }
            }
            catch (ThreadAbortException)
            {
                // Response.End() intentionally terminates the request after streaming the PDF.
            }
            catch (Exception ex)
            {
                gvDummyStatement.DataSource = null;
                gvDummyStatement.DataBind();
                ShowError(ex);
            }
        }

        protected void btnProcess_Click(object sender, EventArgs e)
        {
            ValidateAndRefresh();
            if (!CanPreview()) return;
            try
            {
                using (SqlConnection cn = new SqlConnection(ConnectionString))
                using (SqlCommand cmd = new SqlCommand("dbo.usp_KTA_Billing_Post", cn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 0;
                    AddTvpParameters(cmd);
                    cmd.Parameters.Add("@CompanyID", SqlDbType.VarChar, 10).Value = txtCompany.Text.Trim();
                    cmd.Parameters.Add("@LocationID", SqlDbType.VarChar, 3).Value = txtLocation.Text.Trim();
                    cmd.Parameters.Add("@UpdateUser", SqlDbType.VarChar, 10).Value = SafeUser();
                    cn.Open();
                    cmd.ExecuteNonQuery();
                }
                lblStatus.Text = "Billing selesai. KTA_T_ProcessBilling sudah berstatus 1 untuk periode CAR aktif.";
                lblStatus.CssClass = "alert alert-success d-block mb-0";
                SetProcessState(true);
                LoadPreview();
            }
            catch (Exception ex)
            {
                ShowError(ex);
                SetProcessState(false);
            }
        }

        private void UploadCsvMember()
        {
            if (!fuMember.HasFile) { ShowError("Billing_member.csv belum dipilih."); return; }
            try
            {
                DataTable src = ReadDelimited(fuMember.PostedFile.InputStream, ',');
                RequireHeaders(src, "unitno", "date_trans", "description", "amount");
                DataTable dt = new DataTable();
                dt.Columns.Add("RowNo", typeof(int));
                dt.Columns.Add("UnitNo", typeof(string));
                dt.Columns.Add("DateTrans", typeof(DateTime));
                dt.Columns.Add("Description", typeof(string));
                dt.Columns.Add("Amount", typeof(decimal));
                int row = 0;
                foreach (DataRow r in src.Rows)
                {
                    if (IsBlankRow(r)) continue;
                    row++;
                    DataRow n = dt.NewRow();
                    n["RowNo"] = row;
                    n["UnitNo"] = Text(r, "unitno");
                    n["DateTrans"] = ParseDate(Text(r, "date_trans"));
                    n["Description"] = Text(r, "description");
                    n["Amount"] = ParseMoney(Text(r, "amount"));
                    dt.Rows.Add(n);
                }
                Session[MemberKey] = dt;
                lblMember.Text = row + " rows uploaded.";
                lblMember.CssClass = "text-success small d-block mt-1";
                ValidateAndRefresh();
            }
            catch (Exception ex) { ShowError(ex); }
        }

        private void UploadCsvMonthly()
        {
            if (!fuMonthly.HasFile) { ShowError("Billing Report Monthly belum dipilih."); return; }
            try
            {
                byte[] monthlyBytes = ReadAllBytes(fuMonthly.PostedFile.InputStream);
                DataTable src;
                using (MemoryStream ms = new MemoryStream(monthlyBytes))
                    src = ReadDelimited(ms, '\t');
                if (!HasHeader(src, "driver account"))
                {
                    using (MemoryStream ms = new MemoryStream(monthlyBytes))
                        src = ReadDelimited(ms, ',');
                }
                RequireHeaders(src, "driver account", "subtotal amount", "period");
                DataTable dt = new DataTable();
                dt.Columns.Add("RowNo", typeof(int));
                dt.Columns.Add("DriverAccount", typeof(string));
                dt.Columns.Add("SubtotalAmount", typeof(decimal));
                dt.Columns.Add("SubtotalPowerUsage", typeof(decimal));
                dt.Columns.Add("PeriodFrom", typeof(DateTime));
                dt.Columns.Add("PeriodThru", typeof(DateTime));
                int row = 0;
                foreach (DataRow r in src.Rows)
                {
                    if (IsBlankRow(r)) continue;
                    DateTime from, thru;
                    ParseRangeFromRow(r, out from, out thru);
                    row++;
                    DataRow n = dt.NewRow();
                    n["RowNo"] = row;
                    n["DriverAccount"] = Text(r, "driver account");
                    n["SubtotalAmount"] = ParseMoney(Text(r, "subtotal amount"));
                    n["SubtotalPowerUsage"] = ParseDecimal(Text(r, "subtotal power usage (kwh)"));
                    n["PeriodFrom"] = from; n["PeriodThru"] = thru;
                    dt.Rows.Add(n);
                }
                Session[MonthlyKey] = dt;
                chkMonthlyNA.Checked = false;
                SetOptionalUnavailable("Monthly", false);
                lblMonthly.Text = row + " rows uploaded.";
                lblMonthly.CssClass = "text-success small d-block mt-1";
                ValidateAndRefresh();
            }
            catch (Exception ex) { ShowError(ex); }
        }

        private void UploadCardXlsx()
        {
            if (!fuCard.HasFile) { ShowError("Card Payment Statistics.xlsx belum dipilih."); return; }
            try
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("BatchID", typeof(Guid));
                dt.Columns.Add("RowNo", typeof(int));
                dt.Columns.Add("ChargingCardName", typeof(string));
                dt.Columns.Add("ChargingCardNo", typeof(string));
                dt.Columns.Add("Currency", typeof(string));
                dt.Columns.Add("TotalPowerCost", typeof(decimal));
                dt.Columns.Add("TotalPowerConsumption", typeof(decimal));
                dt.Columns.Add("PeriodFrom", typeof(DateTime));
                dt.Columns.Add("PeriodThru", typeof(DateTime));

                using (ExcelPackage pkg = new ExcelPackage(fuCard.PostedFile.InputStream))
                {
                    ExcelWorksheet ws = pkg.Workbook.Worksheets.FirstOrDefault();
                    if (ws == null) throw new ApplicationException("Workbook tidak memiliki worksheet.");
                    Dictionary<string,int> headers = GetExcelHeaders(ws);
                    int cardNameCol = FindHeader(headers, "charging card name", "card name", "chargingcardname");
                    int cardNoCol = FindHeader(headers, "charging card no", "charging card number", "card no", "chargingcardno");
                    int currencyCol = FindHeader(headers, "currency", "currency id");
                    int costCol = FindHeader(headers, "total power cost", "totalpowercost");
                    int consumptionCol = FindHeader(headers, "total power consumption", "total power usage", "totalpowerconsumption");
                    int periodCol = FindHeader(headers, "statistical period", "period", "statisticalperiod");
                    if (cardNameCol < 0 || costCol < 0 || periodCol < 0)
                        throw new ApplicationException("Header Card Payment Statistics tidak sesuai. Dibutuhkan Charging Card Name, Total Power Cost dan Statistical Period.");

                    int row = 0;
                    for (int r = 2; r <= ws.Dimension.End.Row; r++)
                    {
                        string name = Cell(ws, r, cardNameCol);
                        if (String.IsNullOrWhiteSpace(name)) continue;
                        DateTime from, thru;
                        ParseRange(Cell(ws, r, periodCol), out from, out thru);
                        row++;
                        DataRow n = dt.NewRow();
                        n["RowNo"] = row;
                        n["ChargingCardName"] = name;
                        n["ChargingCardNo"] = cardNoCol < 0 ? "" : Cell(ws, r, cardNoCol);
                        n["Currency"] = currencyCol < 0 ? "" : Cell(ws, r, currencyCol);
                        n["TotalPowerCost"] = ParseMoney(Cell(ws, r, costCol));
                        n["TotalPowerConsumption"] = consumptionCol < 0 ? 0 : ParseDecimal(Cell(ws, r, consumptionCol));
                        n["PeriodFrom"] = from; n["PeriodThru"] = thru;
                        dt.Rows.Add(n);
                    }
                }
                Session[CardKey] = dt;
                chkCardNA.Checked = false;
                SetOptionalUnavailable("Card", false);
                lblCard.Text = dt.Rows.Count + " rows uploaded.";
                lblCard.CssClass = "text-success small d-block mt-1";
                ValidateAndRefresh();
            }
            catch (Exception ex) { ShowError(ex); }
        }

        private void SetOptionalUnavailable(string source, bool value)
        {
            if (source == "Card")
            {
                Session[CardNAKey] = value;
                if (value) Session[CardKey] = EmptyCard();
                else if (Session[CardKey] == null) Session[CardKey] = EmptyCard();
            }
            else
            {
                Session[MonthlyNAKey] = value;
                if (value) Session[MonthlyKey] = EmptyMonthly();
                else if (Session[MonthlyKey] == null) Session[MonthlyKey] = EmptyMonthly();
            }
        }

        private void ValidateAndRefresh()
        {
            try
            {
                using (SqlConnection cn = new SqlConnection(ConnectionString))
                using (SqlCommand cmd = new SqlCommand("dbo.usp_KTA_Billing_Validate", cn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    AddTvpParameters(cmd);
                    cn.Open();
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        if (dr.Read())
                        {
                            int status = Convert.ToInt32(dr["ValidationStatus"]);
                            string msg = dr["ValidationMessage"] == DBNull.Value ? "" : dr["ValidationMessage"].ToString();
                            if (dr["ActivePeriod"] != DBNull.Value)
                                txtActivePeriod.Text = Convert.ToDateTime(dr["ActivePeriod"]).ToString("yyyy-MM");
                            lblStatus.Text = status == 1 ? "Validation OK. Semua source sesuai periode KAR aktif." : msg;
                            lblStatus.CssClass = status == 1 ? "alert alert-success d-block mb-0" : "alert alert-danger d-block mb-0";
                            SetProcessState(false);
                        }
                    }
                }
            }
            catch (Exception ex) { ShowError(ex); }
        }

        private bool CanPreview()
        {
            if (BatchID == Guid.Empty || MemberData.Rows.Count == 0) return false;
            return true; // The Validate SP is called immediately before Preview/Process.
        }

        private void LoadPreview()
        {
            DataSet ds = new DataSet();
            using (SqlConnection cn = new SqlConnection(ConnectionString))
            using (SqlCommand cmd = new SqlCommand("dbo.usp_KTA_Billing_Preview", cn))
            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                AddTvpParameters(cmd);
                da.Fill(ds);
            }
            // Result 0 is validation status; results 1-3 are summary, detail and availability.
            gvSummary.DataSource = ds.Tables.Count > 1 ? ds.Tables[1] : null;
            gvSummary.DataBind();
            gvDetail.DataSource = ds.Tables.Count > 2 ? ds.Tables[2] : null;
            gvDetail.DataBind();
            gvSource.DataSource = ds.Tables.Count > 3 ? ds.Tables[3] : null;
            gvSource.DataBind();
        }

        private void SetProcessState(bool completed)
        {
            bool enabled = !completed && BatchID != Guid.Empty && MemberData.Rows.Count > 0;
            btnProcess.Enabled = enabled;
            btnDummyStatementRun.Enabled = !completed;
            btnPreview.Enabled = !completed;
            btnValidate.Enabled = !completed;
            btnUploadMember.Enabled = !completed;
            btnUploadCard.Enabled = !completed;
            btnUploadMonthly.Enabled = !completed;
            chkCardNA.Enabled = !completed;
            chkMonthlyNA.Enabled = !completed;
        }

        private bool CanPreviewSafe()
        {
            try { return CanPreview(); } catch { return false; }
        }

        private void ClearPreview()
        {
            gvSummary.DataSource = null; gvSummary.DataBind();
            gvDetail.DataSource = null; gvDetail.DataBind();
            gvSource.DataSource = null; gvSource.DataBind();
            gvDummyStatement.DataSource = null; gvDummyStatement.DataBind();
            lblDummyStatement.Text = "";
            lblMember.Text = lblCard.Text = lblMonthly.Text = "";
            chkCardNA.Checked = false; chkMonthlyNA.Checked = false;
        }

        private void AddTvpParameters(SqlCommand cmd)
        {
            cmd.Parameters.Add("@CompanyID", SqlDbType.VarChar, 10).Value = txtCompany.Text.Trim();
            cmd.Parameters.Add("@LocationID", SqlDbType.VarChar, 3).Value = txtLocation.Text.Trim();
            SqlParameter m = cmd.Parameters.AddWithValue("@BillingMember", MemberData);
            m.SqlDbType = SqlDbType.Structured; m.TypeName = "dbo.Type_KTA_BillingMember";
            SqlParameter c = cmd.Parameters.AddWithValue("@BillingCard", CardData);
            c.SqlDbType = SqlDbType.Structured; c.TypeName = "dbo.Type_KTA_BillingCard";
            SqlParameter mo = cmd.Parameters.AddWithValue("@BillingMonthly", MonthlyData);
            mo.SqlDbType = SqlDbType.Structured; mo.TypeName = "dbo.Type_KTA_BillingMonthly";
            cmd.Parameters.Add("@CardUnavailableConfirmed", SqlDbType.Bit).Value = Convert.ToBoolean(Session[CardNAKey] ?? false);
            cmd.Parameters.Add("@MonthlyUnavailableConfirmed", SqlDbType.Bit).Value = Convert.ToBoolean(Session[MonthlyNAKey] ?? false);
        }

        private static DataTable EmptyMember()
        {
            DataTable t = new DataTable();
            t.Columns.Add("RowNo", typeof(int)); t.Columns.Add("UnitNo", typeof(string));
            t.Columns.Add("DateTrans", typeof(DateTime)); t.Columns.Add("Description", typeof(string));
            t.Columns.Add("Amount", typeof(decimal)); return t;
        }

        private static DataTable EmptyCard()
        {
            DataTable t = new DataTable();
            t.Columns.Add("RowNo", typeof(int)); t.Columns.Add("ChargingCardName", typeof(string));
            t.Columns.Add("ChargingCardNo", typeof(string)); t.Columns.Add("Currency", typeof(string));
            t.Columns.Add("TotalPowerCost", typeof(decimal)); t.Columns.Add("TotalPowerConsumption", typeof(decimal));
            t.Columns.Add("PeriodFrom", typeof(DateTime)); t.Columns.Add("PeriodThru", typeof(DateTime)); return t;
        }

        private static DataTable EmptyMonthly()
        {
            DataTable t = new DataTable();
            t.Columns.Add("RowNo", typeof(int)); t.Columns.Add("DriverAccount", typeof(string));
            t.Columns.Add("SubtotalAmount", typeof(decimal)); t.Columns.Add("SubtotalPowerUsage", typeof(decimal));
            t.Columns.Add("PeriodFrom", typeof(DateTime)); t.Columns.Add("PeriodThru", typeof(DateTime)); return t;
        }

        private byte[] ExportCrystalStatement(DataSet data, string unitFrom, string unitThru)
        {
            string reportPath = Server.MapPath("~/Reports/rptKTA_BillingStatement_New.rpt");
            if (!File.Exists(reportPath))
                throw new ApplicationException("File rptKTA_BillingStatement_New.rpt tidak ditemukan di folder Reports.");

            Type reportType = Type.GetType(
                "CrystalDecisions.CrystalReports.Engine.ReportDocument, CrystalDecisions.CrystalReports.Engine", false);
            if (reportType == null)
                throw new ApplicationException("SAP Crystal Reports runtime belum terpasang/terdaftar di IIS.");

            Type exportType = Type.GetType(
                "CrystalDecisions.Shared.ExportFormatType, CrystalDecisions.Shared", false);
            if (exportType == null)
                throw new ApplicationException("Assembly CrystalDecisions.Shared tidak ditemukan.");

            object report = Activator.CreateInstance(reportType);
            try
            {
                reportType.GetMethod("Load", new Type[] { typeof(string) }).Invoke(report, new object[] { reportPath });
                reportType.GetMethod("SetDataSource", new Type[] { typeof(DataSet) }).Invoke(report, new object[] { data });

                // Populate the original stored-procedure parameter fields if they are present in the RPT.
                MethodInfo setParameter = reportType.GetMethod("SetParameterValue", new Type[] { typeof(string), typeof(object) });
                if (setParameter != null)
                {
                    setParameter.Invoke(report, new object[] { "CompanyID", txtCompany.Text.Trim() });
                    setParameter.Invoke(report, new object[] { "LocationID", txtLocation.Text.Trim() });
                    setParameter.Invoke(report, new object[] { "UnitNoFrom", unitFrom });
                    setParameter.Invoke(report, new object[] { "UnitNoThru", unitThru });
                    setParameter.Invoke(report, new object[] { "GroupIDFrom", "" });
                    setParameter.Invoke(report, new object[] { "GroupIDThru", "" });
                    setParameter.Invoke(report, new object[] { "Period", txtActivePeriod.Text.Replace("-", "") });
                    setParameter.Invoke(report, new object[] { "ByWhat", "1" });
                }

                object pdfFormat = Enum.Parse(exportType, "PortableDocFormat");
                MethodInfo export = reportType.GetMethod("ExportToStream", new Type[] { exportType });
                if (export == null) throw new ApplicationException("Crystal ExportToStream(PortableDocFormat) tidak tersedia.");

                using (Stream crystalStream = (Stream)export.Invoke(report, new object[] { pdfFormat }))
                using (MemoryStream output = new MemoryStream())
                {
                    crystalStream.CopyTo(output);
                    return output.ToArray();
                }
            }
            catch (TargetInvocationException ex)
            {
                throw new ApplicationException("Crystal report gagal dirender: " +
                    (ex.InnerException == null ? ex.Message : ex.InnerException.Message), ex);
            }
            finally
            {
                MethodInfo close = reportType.GetMethod("Close", Type.EmptyTypes);
                MethodInfo dispose = reportType.GetMethod("Dispose", Type.EmptyTypes);
                if (close != null) close.Invoke(report, null);
                if (dispose != null) dispose.Invoke(report, null);
            }
        }

        private static DataTable ReadDelimited(Stream stream, char delimiter)
        {
            if (stream.CanSeek) stream.Position = 0;
            using (StreamReader sr = new StreamReader(stream, Encoding.UTF8, true))
            {
                string header = sr.ReadLine();
                if (header == null) throw new ApplicationException("File kosong.");
                List<string> headers = ParseCsvLine(header, delimiter);
                DataTable dt = new DataTable();
                int headerIndex = 0;
                foreach (string h in headers)
                {
                    string columnName = Normalize(h);
                    if (String.IsNullOrWhiteSpace(columnName)) columnName = "__column" + headerIndex;
                    string uniqueName = columnName;
                    int suffix = 1;
                    while (dt.Columns.Contains(uniqueName)) uniqueName = columnName + "_" + suffix++;
                    dt.Columns.Add(uniqueName);
                    headerIndex++;
                }
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    if (String.IsNullOrWhiteSpace(line)) continue;
                    List<string> values = ParseCsvLine(line, delimiter);
                    DataRow r = dt.NewRow();
                    for (int i = 0; i < dt.Columns.Count; i++) r[i] = i < values.Count ? values[i] : "";
                    dt.Rows.Add(r);
                }
                return dt;
            }
        }

        private static List<string> ParseCsvLine(string line, char delimiter)
        {
            List<string> result = new List<string>();
            StringBuilder b = new StringBuilder();
            bool quoted = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    if (quoted && i + 1 < line.Length && line[i + 1] == '"') { b.Append('"'); i++; }
                    else quoted = !quoted;
                }
                else if (c == delimiter && !quoted) { result.Add(b.ToString().Trim()); b.Clear(); }
                else b.Append(c);
            }
            result.Add(b.ToString().Trim());
            return result;
        }

        private static string Normalize(string s)
        {
            return (s ?? "").Trim().Trim('"').ToLowerInvariant();
        }

        private static bool HasHeader(DataTable dt, string header)
        {
            return dt.Columns.Contains(Normalize(header));
        }

        private static void RequireHeaders(DataTable dt, params string[] headers)
        {
            foreach (string h in headers)
                if (!HasHeader(dt, h)) throw new ApplicationException("Kolom wajib tidak ditemukan: " + h);
        }

        private static bool IsBlankRow(DataRow r)
        {
            return r.ItemArray.All(x => x == null || x == DBNull.Value || String.IsNullOrWhiteSpace(x.ToString()));
        }

        private static string Text(DataRow r, string name) { return Convert.ToString(r[Normalize(name)]).Trim(); }

        private static DateTime ParseDate(string value)
        {
            DateTime d;
            string[] formats = { "dd-MM-yyyy","dd/MM/yyyy","yyyy-MM-dd","MM/dd/yyyy","dd-MMM-yyyy","d-M-yyyy","d/M/yyyy" };
            if (DateTime.TryParseExact(value, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out d)) return d;
            if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out d)) return d;
            if (DateTime.TryParse(value, new CultureInfo("id-ID"), DateTimeStyles.None, out d)) return d;
            throw new ApplicationException("Tanggal tidak valid: " + value);
        }

        private static decimal ParseMoney(string value)
        {
            if (String.IsNullOrWhiteSpace(value)) return 0;
            string s = value.Trim().Replace("Rp", "").Replace("IDR", "").Replace(" ", "");
            bool negative = s.StartsWith("(") && s.EndsWith(")");
            s = s.Trim('(',')');
            int comma = s.LastIndexOf(',');
            int dot = s.LastIndexOf('.');
            if (comma >= 0 && dot >= 0)
            {
                if (comma > dot) s = s.Replace(".", "").Replace(',', '.');
                else s = s.Replace(",", "");
            }
            else if (comma >= 0)
            {
                if (s.Length - comma - 1 <= 2) s = s.Replace(',', '.');
                else s = s.Replace(",", "");
            }
            decimal d;
            if (!Decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                throw new ApplicationException("Amount tidak valid: " + value);
            return negative ? -d : d;
        }

        private static decimal ParseDecimal(string value)
        {
            if (String.IsNullOrWhiteSpace(value)) return 0;
            return ParseMoney(value);
        }

        private static void ParseRangeFromRow(DataRow row, out DateTime from, out DateTime thru)
        {
            if (row.Table.Columns.Contains("period"))
            {
                try
                {
                    ParseRange(Convert.ToString(row["period"]), out from, out thru);
                    return;
                }
                catch { }
            }

            foreach (DataColumn column in row.Table.Columns)
            {
                if (String.Equals(column.ColumnName, "period", StringComparison.OrdinalIgnoreCase)) continue;
                try
                {
                    ParseRange(Convert.ToString(row[column]), out from, out thru);
                    return;
                }
                catch { }
            }
            throw new ApplicationException("Kolom periode From-Thru tidak ditemukan pada baris: " +
                String.Join(" | ", row.ItemArray.Select(x => Convert.ToString(x)).ToArray()));
        }

        private static void ParseRange(string value, out DateTime from, out DateTime thru)
        {
            string s = (value ?? "").Trim().Trim('"').Replace('–','-').Replace('—','-');
            Match m = Regex.Match(s, @"(\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]\d{4})\s*(?:_|\s+-\s+|\s+to\s+|\s+THRU\s+)(\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]\d{4})", RegexOptions.IgnoreCase);
            if (!m.Success)
                throw new ApplicationException("Statistical/From-Thru period tidak valid: " + value);
            from = ParseDate(m.Groups[1].Value);
            thru = ParseDate(m.Groups[2].Value);
            if (from > thru) throw new ApplicationException("From period lebih besar dari Thru period: " + value);
        }

        private static byte[] ReadAllBytes(Stream stream)
        {
            if (stream.CanSeek) stream.Position = 0;
            using (MemoryStream ms = new MemoryStream())
            {
                stream.CopyTo(ms);
                return ms.ToArray();
            }
        }

        private static Dictionary<string,int> GetExcelHeaders(ExcelWorksheet ws)
        {
            Dictionary<string,int> h = new Dictionary<string,int>();
            for (int c = ws.Dimension.Start.Column; c <= ws.Dimension.End.Column; c++)
            {
                string v = Normalize(Cell(ws,1,c));
                if (!h.ContainsKey(v)) h.Add(v,c);
            }
            return h;
        }

        private static int FindHeader(Dictionary<string,int> headers, params string[] aliases)
        {
            foreach (string a in aliases)
            {
                string n = Normalize(a);
                if (headers.ContainsKey(n)) return headers[n];
            }
            return -1;
        }

        private static string Cell(ExcelWorksheet ws, int row, int col)
        {
            return ws.Cells[row,col].Text == null ? "" : ws.Cells[row,col].Text.Trim();
        }

        private string SafeUser()
        {
            string u = ResolveUser();
            return String.IsNullOrWhiteSpace(u) ? "WEB" : u.Substring(0, Math.Min(10,u.Length));
        }

        private void ShowError(Exception ex)
        {
            lblStatus.Text = Server.HtmlEncode(ex.Message);
            lblStatus.CssClass = "alert alert-danger d-block mb-0";
        }
    }
}