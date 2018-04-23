using DevExpress.XtraBars.Ribbon;
using DevExpress.Spreadsheet;

namespace DXApplication1 {
    public partial class Form1 : RibbonForm {

        nwindDataSet dataSet;
        nwindDataSetTableAdapters.InvoicesTableAdapter adapter;
        IWorkbook template;

        public Form1() {
            InitializeComponent();

            dataSet = new nwindDataSet();
            adapter = new nwindDataSetTableAdapters.InvoicesTableAdapter();
            adapter.Fill(dataSet.Invoices);

            template = spreadsheetControl1.Document;

            template.MailMergeDataSource = dataSet;
            template.MailMergeDataMember = "Invoices";
        }

        private void Form1_Load(object sender, System.EventArgs e) {
            spreadsheetControl1.LoadDocument("Documents\\MailMergeTemplate_GroupData.xlsx");
        }
    }
}
