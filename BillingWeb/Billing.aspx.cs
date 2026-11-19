using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.UI;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace BillingWeb
{
    public partial class Billing : Page
    {
        private const string BatchViewStateKey = "KTA_Billing_BatchID";
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
            using (SqlConnection cn = new SqlConnection(ConnectionString))
            using (SqlCommand cmd = new SqlCommand("dbo.usp_KTA_Billing_CreateBatch", cn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@CompanyID", SqlDbType.VarChar, 10).Value = txtCompany.Text.Trim();
                cmd.Parameters.Add("@LocationID", SqlDbType.VarChar, 3).Value = txtLocation.Text.Trim();
                cmd.Parameters.Add("@UpdateUser", SqlDbType.VarChar, 10).Value = ResolveUser().Substring(0, Math.Min(10, ResolveUser().Length));
                SqlParameter pBatch = cmd.Parameters.Add("@BatchID", SqlDbType.UniqueIdentifier);
                pBatch.Direction = ParameterDirection.Output;
                SqlParameter pPeriod = cmd.Parameters.Add("@ActivePeriod", SqlDbType.SmallDateTime);
                pPeriod.Direction = ParameterDirection.Output;
                cn.Open();
                cmd.ExecuteNonQuery();
                BatchID = (Guid)pBatch.Value;
                txtBatchID.Text = BatchID.ToString();
                txtActivePeriod.Text = ((DateTime)pPeriod.Value).ToString("yyyy-MM");
            }
            lblStatus.Text = "Batch baru siap. Upload Billing_member dan file optional.";
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
                    cmd.Parameters.Add("@BatchID", SqlDbType.UniqueIdentifier).Value = BatchID;
                    cmd.Parameters.Add("@CompanyID", SqlDbType.VarChar, 10).Value = txtCompany.Text.Trim();
                    cmd.Parameters.Add("@LocationID", SqlDbType.VarChar, 3).Value = txtLocation.Text.Trim();
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

                gvDummyStatement.DataSource = ds.Tables[0];
                gvDummyStatement.DataBind();
                lblDummyStatement.Text = "Dummy statement berhasil dibuat untuk " + ds.Tables[0].Rows.Count +
                    " baris report. Seluruh perubahan billing sudah di-rollback.";
                lblDummyStatement.CssClass = "text-success small d-block mt-2";
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
                    cmd.Parameters.Add("@BatchID", SqlDbType.UniqueIdentifier).Value = BatchID;
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
                dt.Columns.Add("BatchID", typeof(Guid));
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
                    n["BatchID"] = BatchID; n["RowNo"] = row;
                    n["UnitNo"] = Text(r, "unitno");
                    n["DateTrans"] = ParseDate(Text(r, "date_trans"));
                    n["Description"] = Text(r, "description");
                    n["Amount"] = ParseMoney(Text(r, "amount"));
                    dt.Rows.Add(n);
                }
                ReplaceStage("dbo.KTA_W_BillingMember", dt);
                SetUploaded("Member");
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
                dt.Columns.Add("BatchID", typeof(Guid));
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
                    ParseRange(Text(r, "period"), out from, out thru);
                    row++;
                    DataRow n = dt.NewRow();
                    n["BatchID"] = BatchID; n["RowNo"] = row;
                    n["DriverAccount"] = Text(r, "driver account");
                    n["SubtotalAmount"] = ParseMoney(Text(r, "subtotal amount"));
                    n["SubtotalPowerUsage"] = ParseDecimal(Text(r, "subtotal power usage (kwh)"));
                    n["PeriodFrom"] = from; n["PeriodThru"] = thru;
                    dt.Rows.Add(n);
                }
                ReplaceStage("dbo.KTA_W_BillingMonthly", dt);
                SetUploaded("Monthly");
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
                        n["BatchID"] = BatchID; n["RowNo"] = row;
                        n["ChargingCardName"] = name;
                        n["ChargingCardNo"] = cardNoCol < 0 ? "" : Cell(ws, r, cardNoCol);
                        n["Currency"] = currencyCol < 0 ? "" : Cell(ws, r, currencyCol);
                        n["TotalPowerCost"] = ParseMoney(Cell(ws, r, costCol));
                        n["TotalPowerConsumption"] = consumptionCol < 0 ? 0 : ParseDecimal(Cell(ws, r, consumptionCol));
                        n["PeriodFrom"] = from; n["PeriodThru"] = thru;
                        dt.Rows.Add(n);
                    }
                }
                ReplaceStage("dbo.KTA_W_BillingCard", dt);
                SetUploaded("Card");
                chkCardNA.Checked = false;
                SetOptionalUnavailable("Card", false);
                lblCard.Text = dt.Rows.Count + " rows uploaded.";
                lblCard.CssClass = "text-success small d-block mt-1";
                ValidateAndRefresh();
            }
            catch (Exception ex) { ShowError(ex); }
        }

        private void ReplaceStage(string tableName, DataTable dt)
        {
            using (SqlConnection cn = new SqlConnection(ConnectionString))
            {
                cn.Open();
                using (SqlCommand cmd = new SqlCommand("DELETE FROM " + tableName + " WHERE BatchID=@BatchID", cn))
                {
                    cmd.Parameters.Add("@BatchID", SqlDbType.UniqueIdentifier).Value = BatchID;
                    cmd.ExecuteNonQuery();
                }
                using (SqlBulkCopy bc = new SqlBulkCopy(cn))
                {
                    bc.DestinationTableName = tableName;
                    bc.BatchSize = 1000;
                    foreach (DataColumn c in dt.Columns) bc.ColumnMappings.Add(c.ColumnName, c.ColumnName);
                    bc.WriteToServer(dt);
                }
            }
        }

        private void SetUploaded(string source)
        {
            string field = source == "Member" ? "MemberUploaded" : (source == "Card" ? "CardUploaded" : "MonthlyUploaded");
            string naField = source == "Card" ? "CardUnavailableConfirmed" : (source == "Monthly" ? "MonthlyUnavailableConfirmed" : "");
            string sql = "UPDATE dbo.KTA_W_BillingBatch SET " + field + "=1";
            if (!String.IsNullOrWhiteSpace(naField)) sql += ", " + naField + "=0";
            sql += " WHERE BatchID=@BatchID";
            using (SqlConnection cn = new SqlConnection(ConnectionString))
            using (SqlCommand cmd = new SqlCommand(sql, cn))
            {
                cmd.Parameters.Add("@BatchID", SqlDbType.UniqueIdentifier).Value = BatchID;
                cn.Open(); cmd.ExecuteNonQuery();
            }
        }

        private void SetOptionalUnavailable(string source, bool value)
        {
            if (BatchID == Guid.Empty) return;
            string field = source == "Card" ? "CardUnavailableConfirmed" : "MonthlyUnavailableConfirmed";
            string uploaded = source == "Card" ? "CardUploaded" : "MonthlyUploaded";
            using (SqlConnection cn = new SqlConnection(ConnectionString))
            using (SqlCommand cmd = new SqlCommand("UPDATE dbo.KTA_W_BillingBatch SET " + field + "=@Value, " + uploaded + "=CASE WHEN @Value=1 THEN 0 ELSE " + uploaded + " END WHERE BatchID=@BatchID", cn))
            {
                cmd.Parameters.Add("@Value", SqlDbType.Bit).Value = value;
                cmd.Parameters.Add("@BatchID", SqlDbType.UniqueIdentifier).Value = BatchID;
                cn.Open(); cmd.ExecuteNonQuery();
            }
        }

        private void ValidateAndRefresh()
        {
            try
            {
                using (SqlConnection cn = new SqlConnection(ConnectionString))
                using (SqlCommand cmd = new SqlCommand("dbo.usp_KTA_Billing_ValidateBatch", cn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add("@BatchID", SqlDbType.UniqueIdentifier).Value = BatchID;
                    cn.Open();
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        if (dr.Read())
                        {
                            int status = Convert.ToInt32(dr["ValidationStatus"]);
                            string msg = dr["ValidationMessage"] == DBNull.Value ? "" : dr["ValidationMessage"].ToString();
                            DateTime p = Convert.ToDateTime(dr["ActivePeriod"]);
                            txtActivePeriod.Text = p.ToString("yyyy-MM");
                            if (status == 1)
                            {
                                lblStatus.Text = "Validation OK. Semua source sesuai periode KAR aktif.";
                                lblStatus.CssClass = "alert alert-success d-block mb-0";
                                SetProcessState(false);
                            }
                            else
                            {
                                lblStatus.Text = msg;
                                lblStatus.CssClass = "alert alert-danger d-block mb-0";
                                SetProcessState(false);
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { ShowError(ex); }
        }

        private bool CanPreview()
        {
            using (SqlConnection cn = new SqlConnection(ConnectionString))
            using (SqlCommand cmd = new SqlCommand("SELECT ValidationStatus FROM dbo.KTA_W_BillingBatch WHERE BatchID=@BatchID", cn))
            {
                cmd.Parameters.Add("@BatchID", SqlDbType.UniqueIdentifier).Value = BatchID;
                cn.Open();
                object v = cmd.ExecuteScalar();
                return v != null && Convert.ToInt32(v) == 1;
            }
        }

        private void LoadPreview()
        {
            DataSet ds = new DataSet();
            using (SqlConnection cn = new SqlConnection(ConnectionString))
            using (SqlCommand cmd = new SqlCommand("dbo.usp_KTA_Billing_Preview", cn))
            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@BatchID", SqlDbType.UniqueIdentifier).Value = BatchID;
                da.Fill(ds);
            }
            gvSummary.DataSource = ds.Tables.Count > 0 ? ds.Tables[0] : null;
            gvSummary.DataBind();
            gvDetail.DataSource = ds.Tables.Count > 1 ? ds.Tables[1] : null;
            gvDetail.DataBind();
            gvSource.DataSource = ds.Tables.Count > 2 ? ds.Tables[2] : null;
            gvSource.DataBind();
        }

        private void SetProcessState(bool completed)
        {
            bool enabled = !completed && BatchID != Guid.Empty && CanPreviewSafe();
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
            lblMember.Text = lblCard.Text = lblMonthly.Text = "";
            chkCardNA.Checked = false; chkMonthlyNA.Checked = false;
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
                foreach (string h in headers) dt.Columns.Add(Normalize(h));
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