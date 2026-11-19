<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Billing.aspx.cs" Inherits="BillingWeb.Billing" MasterPageFile="~/MasterPages/Site_BS_5.Master" %>

<asp:Content ID="MainContent" ContentPlaceHolderID="MainContent" runat="server">
<div class="toolbar d-flex flex-wrap align-items-center gap-2">
    <asp:Button ID="btnNew" runat="server" Text="New Batch" CssClass="btn btn-outline-secondary btn-sm" OnClick="btnNew_Click" />
    <asp:Button ID="btnValidate" runat="server" Text="Validate" CssClass="btn btn-outline-primary btn-sm" OnClick="btnValidate_Click" />
    <asp:Button ID="btnPreview" runat="server" Text="Preview" CssClass="btn btn-outline-primary btn-sm" OnClick="btnPreview_Click" />
    <asp:Button ID="btnDummyStatement" runat="server" Text="Generate Dummy Statement" CssClass="btn btn-outline-dark btn-sm" OnClick="btnDummyStatement_Click" OnClientClick="return confirm('Generate dummy statement? All temporary billing writes will be rolled back.');" />
    <asp:Button ID="btnProcess" runat="server" Text="Process Billing" CssClass="btn btn-primary btn-sm" OnClick="btnProcess_Click"
        OnClientClick="return confirm('Process billing untuk periode aktif? Setelah berhasil billing tidak dapat diproses ulang untuk periode yang sama.');" />
</div>

<div class="card mb-3">
    <div class="card-body">
        <div class="row g-3">
            <div class="col-md-3"><label class="form-label">Company</label><asp:TextBox ID="txtCompany" runat="server" CssClass="form-control form-control-sm" ReadOnly="true" /></div>
            <div class="col-md-2"><label class="form-label">Location</label><asp:TextBox ID="txtLocation" runat="server" CssClass="form-control form-control-sm" ReadOnly="true" /></div>
            <div class="col-md-3"><label class="form-label">Active KAR Period</label><asp:TextBox ID="txtActivePeriod" runat="server" CssClass="form-control form-control-sm" ReadOnly="true" /></div>
            <div class="col-md-4"><label class="form-label">Batch ID</label><asp:TextBox ID="txtBatchID" runat="server" CssClass="form-control form-control-sm" ReadOnly="true" /></div>
        </div>
        <div class="mt-3 status-box"><asp:Label ID="lblStatus" runat="server" CssClass="alert alert-info d-block mb-0" /></div>
    </div>
</div>

<div class="card mb-3">
    <div class="card-header fw-semibold">Source Files</div>
    <div class="card-body">
        <div class="row g-3">
            <div class="col-lg-4">
                <label class="form-label fw-semibold">1. Billing_member.csv <span class="text-danger">*</span></label>
                <div class="input-group input-group-sm">
                    <asp:FileUpload ID="fuMember" runat="server" CssClass="form-control" />
                    <asp:Button ID="btnUploadMember" runat="server" Text="Upload" CssClass="btn btn-outline-primary" OnClick="btnUploadMember_Click" />
                </div>
                <div class="small text-muted mt-1">Kolom: unitno, date_trans, description, amount</div>
                <asp:Label ID="lblMember" runat="server" CssClass="small d-block mt-1" />
            </div>
            <div class="col-lg-4">
                <label class="form-label fw-semibold">2. Card Payment Statistics.xlsx</label>
                <div class="input-group input-group-sm">
                    <asp:FileUpload ID="fuCard" runat="server" CssClass="form-control" />
                    <asp:Button ID="btnUploadCard" runat="server" Text="Upload" CssClass="btn btn-outline-primary" OnClick="btnUploadCard_Click" />
                </div>
                <div class="form-check mt-2">
                    <asp:CheckBox ID="chkCardNA" runat="server" CssClass="form-check-input" Text=" File tidak tersedia / sudah dikonfirmasi tidak tersedia" AutoPostBack="true" OnCheckedChanged="chkCardNA_CheckedChanged" />
                </div>
                <asp:Label ID="lblCard" runat="server" CssClass="small d-block mt-1" />
            </div>
            <div class="col-lg-4">
                <label class="form-label fw-semibold">3. Billing Report Monthly.csv / .tsv</label>
                <div class="input-group input-group-sm">
                    <asp:FileUpload ID="fuMonthly" runat="server" CssClass="form-control" />
                    <asp:Button ID="btnUploadMonthly" runat="server" Text="Upload" CssClass="btn btn-outline-primary" OnClick="btnUploadMonthly_Click" />
                </div>
                <div class="form-check mt-2">
                    <asp:CheckBox ID="chkMonthlyNA" runat="server" CssClass="form-check-input" Text=" File tidak tersedia / sudah dikonfirmasi tidak tersedia" AutoPostBack="true" OnCheckedChanged="chkMonthlyNA_CheckedChanged" />
                </div>
                <asp:Label ID="lblMonthly" runat="server" CssClass="small d-block mt-1" />
            </div>
        </div>
    </div>
</div>

<div class="card mb-3">
    <div class="card-header fw-semibold">Dummy Billing Statement</div>
    <div class="card-body">
        <div class="row g-2 align-items-end">
            <div class="col-md-3"><label class="form-label">Unit From</label><asp:TextBox ID="txtDummyUnitFrom" runat="server" CssClass="form-control form-control-sm" /></div>
            <div class="col-md-3"><label class="form-label">Unit Thru</label><asp:TextBox ID="txtDummyUnitThru" runat="server" CssClass="form-control form-control-sm" /></div>
            <div class="col-md-3"><asp:Button ID="btnDummyStatementRun" runat="server" Text="Generate Dummy Statement" CssClass="btn btn-outline-dark btn-sm" OnClick="btnDummyStatement_Click" OnClientClick="return confirm('Generate dummy statement? Temporary billing writes will be rolled back.');" /></div>
        </div>
        <asp:Label ID="lblDummyStatement" runat="server" CssClass="small d-block mt-2" />
        <div class="table-responsive mt-2">
            <asp:GridView ID="gvDummyStatement" runat="server" CssClass="table table-sm table-bordered table-hover mb-0" AutoGenerateColumns="true" EmptyDataText="Belum ada dummy statement." />
        </div>
        <div class="small text-muted mt-2">Dummy menggunakan nomor sementara berawalan D, tidak mengonsumsi nomor KBI, dan seluruh perubahan database di-rollback. Data yang ditampilkan mengikuti dataset SP billing statement legacy.</div>
    </div>
</div>

<div class="card mb-3">
    <div class="card-header fw-semibold">Validation / Control Total</div>
    <div class="card-body">
        <asp:GridView ID="gvSummary" runat="server" CssClass="table table-sm table-bordered table-hover mb-0" AutoGenerateColumns="true" EmptyDataText="Belum ada preview." />
    </div>
</div>

<div class="card mb-3">
    <div class="card-header fw-semibold">Unit / Source Detail</div>
    <div class="card-body table-responsive">
        <asp:GridView ID="gvDetail" runat="server" CssClass="table table-sm table-bordered table-hover mb-0" AutoGenerateColumns="true" EmptyDataText="Belum ada preview." />
    </div>
</div>

<div class="card mb-3">
    <div class="card-header fw-semibold">Source Availability</div>
    <div class="card-body table-responsive">
        <asp:GridView ID="gvSource" runat="server" CssClass="table table-sm table-bordered table-hover mb-0" AutoGenerateColumns="true" EmptyDataText="Belum ada preview." />
    </div>
</div>

<div class="alert alert-secondary small mb-0">
    Preview hanya membaca staging. Nomor Billing KBI baru dibuat saat tombol <strong>Process Billing</strong> dijalankan dan seluruh posting berada dalam satu transaction SQL.
</div>
</asp:Content>