using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Data.Sqlite;
using System.Data;

namespace Conv.Net
{
    public partial class Form1 : Form
    {
        DataTable tableLoad = new DataTable();
        public Form1()
        {
            InitializeComponent();

            this.dataGridView.SetDoubleBuffered(true);
        }

        void Refresh()
        {
            dataGridView.DataSource = tableLoad;
        }

        public static DataTable ReadExcelFile(string filePath)
        {
            DataTable dataTable = new DataTable();
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                if (sheet != null)
                {
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    Worksheet worksheet = worksheetPart.Worksheet;
                    SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData != null)
                    {
                        Row firstRow = sheetData.Descendants<Row>().FirstOrDefault();
                        if (firstRow != null)
                        {
                            foreach (Cell cell in firstRow.Descendants<Cell>())
                            {
                                string columnName = GetCellValue(cell, workbookPart);
                                dataTable.Columns.Add(columnName);
                            }
                            foreach (Row row in sheetData.Descendants<Row>().Skip(1))
                            {
                                DataRow dataRow = dataTable.NewRow();
                                int columnIndex = 0;
                                foreach (Cell cell in row.Descendants<Cell>())
                                {
                                    string cellValue = GetCellValue(cell, workbookPart);
                                    dataRow[columnIndex] = cellValue;
                                    columnIndex++;
                                }
                                dataTable.Rows.Add(dataRow);
                            }
                        }
                    }
                }
            }


            for (int i = dataTable.Rows.Count - 1; i >= 0; i--)
            {
                bool hasValue = false;
                foreach (DataColumn column in dataTable.Columns)
                {
                    if (!string.IsNullOrEmpty(dataTable.Rows[i][column].ToString()))
                    {
                        hasValue = true;
                        break;
                    }
                }
                if (!hasValue)
                {
                    dataTable.Rows.RemoveAt(i);
                }
            }

            DB.ExecuteQuery
                (
                    "DELETE FROM dataconvert"
                );

            foreach (DataRow row1 in dataTable.Rows)
            {
                var semestr = "";
                var disciplina = "";
                var fin = "";
                var raspred = "";
                var napravl = "";
                var groupp = "";
                var time_all = "";
                var podgrupp = "";
                var students = "";
                var potok = "";
                var lek = "";
                var prakt = "";
                var lab = "";
                var kursovoi = "";
                var ucheb_praktika = "";

                if (dataTable.Columns.Contains("�������"))
                {
                    semestr = (row1["�������"]).ToString();
                }

                if (dataTable.Columns.Contains("����������"))
                {
                    disciplina = (row1["����������"]).ToString();
                }

                if (dataTable.Columns.Contains("���"))
                {
                    fin = (row1["���"]).ToString();
                }

                if (dataTable.Columns.Contains("�����������"))
                {
                    napravl = (row1["�����������"]).ToString();
                }

                if (dataTable.Columns.Contains("������"))
                {
                    groupp = (row1["������"]).ToString();
                }

                if (dataTable.Columns.Contains("�����\r\n �����"))
                {
                    time_all = (row1["�����\r\n �����"]).ToString();
                }

                if (dataTable.Columns.Contains("�����"))
                {
                    podgrupp = (row1["�����"]).ToString();
                }

                if (dataTable.Columns.Contains("���-��\r\n ���������"))
                {
                    students = (row1["���-��\r\n ���������"]).ToString();
                }

                if (dataTable.Columns.Contains("������"))
                {
                    potok = (row1["������"]).ToString();
                }

                if (dataTable.Columns.Contains("���"))
                {
                    lek = (row1["���"]).ToString();
                }

                if (dataTable.Columns.Contains("��/\r\n ���"))
                {
                    prakt = (row1["��/\r\n ���"]).ToString();
                }

                if (dataTable.Columns.Contains("���"))
                {
                    lab = (row1["���"]).ToString();
                }

                if (dataTable.Columns.Contains("��/��"))
                {
                    kursovoi = (row1["��/��"]).ToString();
                }

                if (dataTable.Columns.Contains("��.��"))
                {
                    ucheb_praktika = (row1["��.��"]).ToString();
                }

                try
                {
                    DB.ExecuteQuery
                    (
                        "INSERT INTO dataconvert (semestr, disciplina, fin, raspred, napravl, groupp, " +
                        "time_all, podgrupp, students, potok, lek, prakt, lab, kursovoi, ucheb_praktika) " +
                        "VALUES (@semestr, @disciplina, @fin, @raspred, @napravl, @groupp, " +
                        "@time_all, @podgrupp, @students, @potok, @lek, @prakt, @lab, @kursovoi, @ucheb_praktika)",
                        new SqliteParameter("@semestr", semestr),
                        new SqliteParameter("@disciplina", disciplina),
                        new SqliteParameter("@fin", fin),
                        new SqliteParameter("@raspred", raspred),
                        new SqliteParameter("@napravl", napravl),
                        new SqliteParameter("@groupp", groupp),
                        new SqliteParameter("@time_all", time_all),
                        new SqliteParameter("@podgrupp", podgrupp),
                        new SqliteParameter("@students", students),
                        new SqliteParameter("@potok", potok),
                        new SqliteParameter("@lek", lek),
                        new SqliteParameter("@prakt", prakt),
                        new SqliteParameter("@lab", lab),
                        new SqliteParameter("@kursovoi", kursovoi),
                        new SqliteParameter("@ucheb_praktika", ucheb_praktika)
                    );
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

            }

            return dataTable;
        }

        private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            string value = cell.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
                if (sharedStringTablePart != null)
                {
                    value = sharedStringTablePart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
            }
            return value;
        }

        private void buttonOpenFile_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog open = new OpenFileDialog();
                open.Filter = "Exel Files|*.xls;*.xlsx;*.xlsm";
                if (open.ShowDialog() == DialogResult.OK)
                {
                    string excelFilePath = open.FileName;
                    var table = ReadExcelFile(open.FileName);
                    tableLoad = table;
                }

                Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void buttonConver_Click(object sender, EventArgs e)
        {

        }
    }
}