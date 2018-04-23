Imports Microsoft.VisualBasic
Imports DevExpress.XtraBars.Ribbon
Imports DevExpress.Spreadsheet

Partial Public Class Form1
    Inherits RibbonForm

    Private dataSet As nwindDataSet
    Private adapter As nwindDataSetTableAdapters.InvoicesTableAdapter
    Private template As IWorkbook

	Public Sub New()
        InitializeComponent()

        dataSet = New nwindDataSet()
        adapter = New nwindDataSetTableAdapters.InvoicesTableAdapter()
        adapter.Fill(dataSet.Invoices)

        template = SpreadsheetControl1.Document

        template.MailMergeDataSource = dataSet
        template.MailMergeDataMember = "Invoices"
	End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SpreadsheetControl1.LoadDocument("Documents\MailMergeTemplate_GroupData.xlsx")
    End Sub
End Class
