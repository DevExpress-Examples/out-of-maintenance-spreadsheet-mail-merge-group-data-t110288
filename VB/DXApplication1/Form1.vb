Imports DevExpress.XtraBars.Ribbon
Imports DevExpress.Spreadsheet

Namespace DXApplication1
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
        End Sub

        Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            spreadsheetControl1.LoadDocument("Documents\MailMergeTemplate_GroupData.xlsx")
            template = spreadsheetControl1.Document
            template.MailMergeDataSource = dataSet
            template.MailMergeDataMember = "Invoices"
        End Sub
    End Class
End Namespace
