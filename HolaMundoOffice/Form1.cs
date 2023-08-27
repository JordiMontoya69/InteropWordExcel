using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace HolaMundoOffice
{
    public partial class Form1 : Form
    {
        SaveFileDialog sfd;

        public Form1()
        {
            sfd = new SaveFileDialog();
            InitializeComponent();
        }

        private void btnWord_Click(object sender, EventArgs e)
        {
            if(sfd.ShowDialog() == DialogResult.OK)
            {
                Word.Application word = new Word.Application();
                Word.Document doc = word.Documents.Add();
                doc.Activate();
                word.Selection.TypeText(tbTexto.Text);
                doc.SaveAs2(sfd.FileName);
                word.Visible = true;
            }
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            if(sfd.ShowDialog() == DialogResult.OK)
            {
                Excel.Application excel = new Excel.Application();
                Excel.Workbook file = excel.Workbooks.Add();
                Excel.Worksheet hoja = file.Worksheets.Add();

                hoja.Cells.Range["A1"].Value = tbTexto.Text;
                file.SaveAs(sfd.FileName);
                excel.Visible = true;
            }
        }
    }
}
