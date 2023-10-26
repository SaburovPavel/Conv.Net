
using Epplus = OfficeOpenXml;
using System.Diagnostics;
using OfficeOpenXml.Style;

namespace Conv.Net
{
    public class ExcelEpp
    {
        private Epplus.ExcelPackage package;
        public bool toEmail = false;        
        public bool biathlon = false;
        public bool gender = false;

        private System.Data.DataTable tableCopy = new System.Data.DataTable();

        public ExcelEpp()
        {
            package = new Epplus.ExcelPackage();
            toEmail = false;
        }

        public bool ShablonSartUsers()
        {
            try
            {
                var sheet = package.Workbook.Worksheets.Add("Лист1");
                int row = 1, col = 1;

                int sizeText = 10;
                int widthCol = 14;

                
                sheet.Cells[row, col].Value = "№"; sheet.Column(col).Width = 4; col++;
                sheet.Cells[row, col].Value = "familia"; sheet.Column(col).Width = widthCol; col++;
                sheet.Cells[row, col].Value = "name"; sheet.Column(col).Width = widthCol; col++;
                sheet.Cells[row, col].Value = "datebirth"; sheet.Column(col).Width = widthCol; col++;
                sheet.Cells[row, col].Value = "gender"; sheet.Column(col).Width = widthCol; col++;
                sheet.Cells[row, col].Value = "startnum"; sheet.Column(col).Width = widthCol; col++;
                sheet.Cells[row, col].Value = "town"; sheet.Column(col).Width = widthCol; col++;
                sheet.Cells[row, col].Value = "group"; sheet.Column(col).Width = widthCol; col++;
                sheet.Cells[row, col].Value = "chip_number"; sheet.Column(col).Width = widthCol; col++;
                sheet.Cells[row, col].Value = "email"; sheet.Column(col).Width = widthCol; col++;
                sheet.Cells[row, col].Value = "timeStartUser"; sheet.Column(col).Width = widthCol; col++;
                sheet.Cells[row, col].Value = "idx"; sheet.Column(col).Width = widthCol; col++;
                sheet.Cells[row, col].Value = "server"; sheet.Column(col).Width = widthCol; col++;
                sheet.Cells[row, col].Value = "SN2"; sheet.Column(col).Width = widthCol; col++;

                sheet.Cells[row, 1, row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, 1, row, col].Style.Font.Name = "Courier New";                
                sheet.Cells[row, 1, row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[row, 1, row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells[row, 1, row, col].Style.Font.Bold = true;

                for (int i = 1; i < 100; i++)
                {
                    sheet.Cells[i, 1].Value = i; 
                }

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
                
                string defaultPath = "C:\\!ProtocolsDirectory\\";

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
                            MessageBox.Show("протокол был сформировон и сохранён по пути:" + saveFileDialog.FileName, "Выполнено", MessageBoxButtons.OK);
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
