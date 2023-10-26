
using Epplus = OfficeOpenXml;
using System.Diagnostics;
using OfficeOpenXml.Style;
using System.Data;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml;

namespace Conv.Net
{
    public class ExcelEpp
    {
        private Epplus.ExcelPackage package;
        public bool toEmail = false;        
        public bool biathlon = false;
        public bool gender = false;

        private DataTable tableCopy = new DataTable();

        public ExcelEpp()
        {
            package = new Epplus.ExcelPackage();
            toEmail = false;
        }

        public bool ConvertToExcel(DataTable dataTable)
        {
            try
            {
                var sheet = package.Workbook.Worksheets.Add("Лист1");
                int row = 1, col = 1;
                var font = "Times New Roman";
                sheet.PrinterSettings.Orientation = eOrientation.Landscape;


                int sizeText = 10;
                int widthCol = 14;

                sheet.Cells[row, col].Value = "Основное распределение нагрузки преподавателей Колледжа ВлГУ на _________ семестр 20  -20   уч. года";
                sheet.Cells[row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[row, col, row, 11 ].Merge = true;
                row++;

                sheet.Cells[row, col].Value = "Преподаватель\r\n(должность,\r\nФ.И.О.)";
                sheet.Cells[row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(col).Width = 10;
                sheet.Cells[row, col].Style.WrapText = true;


                col++;

                sheet.Cells[row, col].Value = "Дисциплина";
                sheet.Cells[row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(col).Width = 10;
                sheet.Cells[row, col].Style.WrapText = true;

                col++;

                sheet.Cells[row, col].Value = "Кафедра";
                sheet.Cells[row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(col).Width = 8;
                sheet.Cells[row, col].Style.WrapText = true;

                col++;

                sheet.Cells[row, col].Value = "Тип\r\nзанятий\r\n";
                sheet.Cells[row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(col).Width = 8;
                sheet.Cells[row, col].Style.WrapText = true;

                col++;

                sheet.Cells[row, col].Value = "Группа (поток)";
                sheet.Cells[row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(col).Width = 22;
                sheet.Cells[row, col].Style.WrapText = true;

                col++;

                sheet.Cells[row, col].Value = "Под\r\nгруппа";
                sheet.Cells[row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(col).Width = 10;
                sheet.Cells[row, col].Style.WrapText = true;

                col++;

                sheet.Cells[row, col].Value = "Состав";
                sheet.Cells[row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(col).Width = 10;
                sheet.Cells[row, col].Style.WrapText = true;

                col++;

                sheet.Cells[row, col].Value = "Кол.\r\nчасов";
                sheet.Cells[row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(col).Width = 10;
                sheet.Cells[row, col].Style.WrapText = true;

                col++;

                sheet.Cells[row, col].Value = "Группи\r\nровка";
                sheet.Cells[row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(col).Width = 10;
                sheet.Cells[row, col].Style.WrapText = true;

                col++;

                sheet.Cells[row, col].Value = "№\r\nауд.";
                sheet.Cells[row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(col).Width = 10;
                sheet.Cells[row, col].Style.WrapText = true;

                col++;

                sheet.Cells[row, col].Value = "Число\r\nауд.";
                sheet.Cells[row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(col).Width = 10;
                sheet.Cells[row, col].Style.WrapText = true;

                row++;
                col++;






                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }              

        public void SaveEppFile(string fileNameOfProtocol)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Файлы Excel (*.xlsx)|*.xlsx";
                saveFileDialog.Title = "Сохранить файл Excel";
                saveFileDialog.FileName = fileNameOfProtocol;
                
                string defaultPath = "C:\\!Raspredelenie\\";

                if (!Directory.Exists(defaultPath))
                {
                    Directory.CreateDirectory(defaultPath);
                }
                saveFileDialog.InitialDirectory = defaultPath;
                
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (!string.IsNullOrEmpty(saveFileDialog.FileName))
                    {
                        // Сохранение документа
                        package.SaveAs(new FileInfo(saveFileDialog.FileName));


                        ButtonTwoForm form = new ButtonTwoForm("Открыть проводник и показать этот файл в папке", "Отправить по email");
                        form.ShowDialog();

                        if (form.left)
                        {
                            Process.Start("explorer.exe", "/select," + saveFileDialog.FileName);
                        }

                        if (form.right)
                        {
                            toEmail = true;
                        }

                        if (form.DialogResult == DialogResult.Cancel)
                        {
                            MessageBox.Show("распределение сформировоно и сохранено по пути:" + saveFileDialog.FileName, "Выполнено", MessageBoxButtons.OK);
                        }

                    }
                    else { return; }
                };    
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка:" + ex.Message, "Ошибка", MessageBoxButtons.OK);                
            }
        }

        public bool ExportToExcelDataTable (System.Data.DataTable dataTable)
        {
            if (dataTable ==null ||  dataTable.Rows.Count == 0) { return false; }

            try
            {
                var sheet = package.Workbook.Worksheets.Add("лист 1");
                //int row = 1, col = 1;

                for(int row = 1; row < dataTable.Rows.Count + 1; row++)
                {
                    for(int col = 1; col < dataTable.Columns.Count + 1; col++)
                    {
                        sheet.Cells[row, col].Value = dataTable.Rows[row - 1][col - 1].ToString();
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка:" + ex.Message, "Ошибка", MessageBoxButtons.OK);
                return false;
            }
        }
        public bool saveAs(String file)
        {
            try
            {
                file = file.Replace("\\\\", "\\");
                if (!File.Exists(file))
                {
                    package.SaveAs(new FileInfo(file));
                    //MessageBox.Show("протокол был сформировон и сохранён по пути:\\n" + file, "Выполнено", MessageBoxButtons.OK);
                }
                else
                {
                    file = file.Insert(file.Length - 5, "(1)");
                    this.saveAs(file, 1);
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка:" + ex.Message, "Ошибка", MessageBoxButtons.OK);
                return false;
            }
        }
        private bool saveAs(String file, int i)
        {
            try
            {
                if (!File.Exists(file))
                {
                    package.SaveAs(new FileInfo(file));
                    //MessageBox.Show("протокол был сформировон и сохранён по пути:" + file, "Выполнено", MessageBoxButtons.OK);
                }
                else
                {
                    file = file.Replace("(" + i + ")", "(" + (i + 1) + ")");
                    i++;
                    this.saveAs(file, i);
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка:" + ex.Message, "Ошибка", MessageBoxButtons.OK);
                return false;
            }
        }
    }
}
