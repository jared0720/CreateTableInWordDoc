using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace CreateWordDoc
{
    public partial class CreateWordDoc : Form
    {
        public CreateWordDoc()
        {
            InitializeComponent();

            PopulateDataTable();
            dgv_data.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            dgv_data.AllowUserToAddRows = false;      }

        private void PopulateDataTable()
        {
            DataTable dataTable1 = new DataTable();

            for (int i = 0; i < 3; i++)
            {
                dataTable1.Columns.Add("Column" + (i + 1).ToString(), typeof(string));
            }

            for (int i = 0; i < 5; i++)
            {
                DataRow dataRow = dataTable1.NewRow();

                dataRow[0] = "test" + i;
                dataRow[1] = "test" + i;
                dataRow[2] = "test" + i;
                dataTable1.Rows.Add(dataRow);
            }

            dgv_data.DataSource = dataTable1;
        }

        private void btn_export_Click(object sender, EventArgs e)
        {
            ExportToWord(dgv_data);
        }

        private void CreateAndPopulateTable(DataGridView dgv, Word.Document wordDoc)
        {
            object endOfDoc = "\\endofdoc";
            object missing = Type.Missing;
            int rowCount = dgv.Rows.Count;
            int colCount = dgv.Columns.Count;
            Word.Table table;

            Word.Range range = wordDoc.Bookmarks.get_Item(ref endOfDoc).Range;
            table = wordDoc.Tables.Add(range, 1, colCount, ref missing, ref missing);
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.AllowAutoFit = true;

            foreach (DataGridViewCell cell in dgv.Rows[0].Cells)
            {
                table.Cell(table.Rows.Count + 1, cell.ColumnIndex + 1).Range.Text = dgv.Columns[cell.ColumnIndex].Name;
            }
            table.Rows.Add();
                           
            foreach (DataGridViewRow row in dgv.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    table.Cell(table.Rows.Count + 1, cell.ColumnIndex + 1).Range.Text = cell.Value.ToString();
                }
                table.Rows.Add();
            }
            
            table.Rows[rowCount + 2].Delete();
            wordDoc.Paragraphs.Add();
        }

        private void ExportToWord(DataGridView dgv)
        {
            int rowCount = dgv.Rows.Count;
            int colCount = dgv.Columns.Count;
            object missing = Type.Missing;
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = new Word.Document();
            wordApp.Visible = true;
            wordApp.WindowState = Word.WdWindowState.wdWindowStateNormal;
            wordDoc = wordApp.Documents.Add(ref missing, ref missing, ref missing, ref missing);

            CreateAndPopulateTable(dgv_data, wordDoc);

            SaveDocument(wordApp);
        }

        private void SaveDocument(Word.Application wordApp)
        {
            object fileName = "C:\\Users\\jared\\Desktop\\test.pdf";
            object missing = Type.Missing;
            wordApp.ActiveDocument.SaveAs2(ref fileName, ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing, ref missing,
                                            ref missing);
        }
    }
}