
using Epplus = OfficeOpenXml;
using System.Diagnostics;
using OfficeOpenXml.Style;
using System.Data;
using OfficeOpenXml;
using Microsoft.Data.Sqlite;
using System.Text.RegularExpressions;

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
        
        private string CheckTitleGroup(string group, string potok)
        {
            var groupp = group;
            if (potok != "")
            {
                groupp = string.Join("", Regex.Matches(potok, @"\(([^()]+)\)").Cast<Match>().Select(m => m.Groups[1].Value));
            }
            else
            {
                groupp = Regex.Replace(groupp, @"\([^()]*\)", "");
            }
            return groupp;
        }

        public bool ConvertToExcel(DataTable dataTable)
        {
            try
            {
                var semestr = "Осенний";

                ButtonTwoForm buttonTwoForm = new ButtonTwoForm("Осенний", "Весенний");
                if (buttonTwoForm.ShowDialog() == DialogResult.OK)
                {
                    if (buttonTwoForm.right)
                    {
                        semestr = "Весенний";
                    }
                }

                var sheet = package.Workbook.Worksheets.Add("Лист1");
                int row = 1, col = 1;
                var font = "Times New Roman";
                sheet.PrinterSettings.Orientation = eOrientation.Landscape;
                sheet.PrinterSettings.LeftMargin = 0.4m; // Отступ слева 0,5 дюйма
                sheet.PrinterSettings.RightMargin = 0.4m; // Отступ справа 0,5 дюйма


                int sizeText = 10;
                int widthCol = 14;

                sheet.Cells[row, col].Value = $"Основное распределение нагрузки преподавателей Колледжа ВлГУ на {semestr} семестр 20  -20   уч. года";
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
                sheet.Column(col).Width = 15;
                sheet.Cells[row, col].Style.WrapText = true;
                sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                col++;

                sheet.Cells[row, col].Value = "Дисциплина";
                sheet.Cells[row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(col).Width = 11;
                sheet.Cells[row, col].Style.WrapText = true;
                sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                col++;

                sheet.Cells[row, col].Value = "Кафедра";
                sheet.Cells[row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(col).Width = 8;
                sheet.Cells[row, col].Style.WrapText = true;
                sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                col++;

                sheet.Cells[row, col].Value = "Тип\r\nзанятий\r\n";
                sheet.Cells[row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(col).Width = 8;
                sheet.Cells[row, col].Style.WrapText = true;
                sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                col++;

                sheet.Cells[row, col].Value = "Группа (поток)";
                sheet.Cells[row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(col).Width = 22;
                sheet.Cells[row, col].Style.WrapText = true;
                sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                col++;

                sheet.Cells[row, col].Value = "Под\r\nгруппа";
                sheet.Cells[row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(col).Width = 10;
                sheet.Cells[row, col].Style.WrapText = true;
                sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                col++;

                sheet.Cells[row, col].Value = "Состав";
                sheet.Cells[row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(col).Width = 10;
                sheet.Cells[row, col].Style.WrapText = true;
                sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                col++;

                sheet.Cells[row, col].Value = "Кол.\r\nчасов";
                sheet.Cells[row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(col).Width = 10;
                sheet.Cells[row, col].Style.WrapText = true;
                sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                col++;

                sheet.Cells[row, col].Value = "Группи\r\nровка";
                sheet.Cells[row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(col).Width = 10;
                sheet.Cells[row, col].Style.WrapText = true;
                sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                col++;

                sheet.Cells[row, col].Value = "№\r\nауд.";
                sheet.Cells[row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(col).Width = 10;
                sheet.Cells[row, col].Style.WrapText = true;
                sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                col++;

                sheet.Cells[row, col].Value = "Число\r\nауд.";
                sheet.Cells[row, col].Style.Font.Size = sizeText;
                sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(col).Width = 10;
                sheet.Cells[row, col].Style.WrapText = true;
                sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Cells[row, 1, row, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sheet.Cells[row, 1, row, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells[row, 1, row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                sheet.Cells[row, 1, row, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                row++;

                
                var tableSpec = DB.ExecuteQuery
                    (
                        "SELECT DISTINCT napravl from dataconvert WHERE semestr=@semestr",
                        new SqliteParameter("@semestr", semestr)
                    );

                

                if(tableSpec != null && tableSpec.Rows.Count > 0)
                {
                    foreach (DataRow dataRow in tableSpec.Rows )
                    {
                        var napravl = dataRow[0].ToString();

                        if(napravl != "")
                        {
                            var tableUnicDiscipl = DB.ExecuteQuery
                            (
                                "SELECT DISTINCT disciplina FROM dataconvert WHERE napravl=@napravl AND semestr=@semestr",
                                new SqliteParameter("@semestr", semestr),
                                new SqliteParameter("@napravl", napravl)
                            );

                            col = 1;

                            sheet.Cells[row, col].Value = napravl;
                            sheet.Cells[row, col].Style.Font.Size = sizeText;
                            sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                            sheet.Cells[row, col].Style.Font.Bold = true;
                            sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            sheet.Cells[row, col, row, 11].Merge = true;

                            sheet.Cells[row, 1, row, 11].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            sheet.Cells[row, 1, row, 11].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            sheet.Cells[row, 1, row, 11].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            sheet.Cells[row, 1, row, 11].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                            row++;

                            if (tableUnicDiscipl != null && tableUnicDiscipl.Rows.Count > 0)
                            {
                                foreach(DataRow rowDiscipl  in tableUnicDiscipl.Rows)
                                {
                                    var disciplina = rowDiscipl[0].ToString();

                                    if (disciplina != "")
                                    {
                                        var tableDiscipllekDistinkt = DB.ExecuteQuery
                                            (
                                                "SELECT DISTINCT raspred FROM dataconvert " +
                                                "WHERE disciplina=@disciplina AND napravl=@napravl AND semestr=@semestr",
                                                new SqliteParameter("@semestr", semestr),
                                                new SqliteParameter("@napravl", napravl),
                                                new SqliteParameter("@disciplina", disciplina)
                                            );

                                        if (tableDiscipllekDistinkt.Rows.Count > 0)
                                        {
                                            foreach (DataRow rowRaspr in tableDiscipllekDistinkt.Rows)
                                            {
                                                var raspred = rowRaspr["raspred"].ToString();

                                                var tableDiscipllek = DB.ExecuteQuery
                                                    (
                                                        "SELECT raspred, disciplina,kaf, groupp,podgrupp,sum(students),potok, SUM(lek) FROM dataconvert " +
                                                        "WHERE disciplina=@disciplina AND napravl=@napravl AND semestr=@semestr AND raspred=@raspred",
                                                        new SqliteParameter("@semestr", semestr),
                                                        new SqliteParameter("@napravl", napravl),
                                                        new SqliteParameter("@disciplina", disciplina),
                                                        new SqliteParameter("@raspred", raspred)
                                                    );

                                                if (tableDiscipllek.Rows[0]["SUM(lek)"].ToString() != "0")
                                                {

                                                    col = 1;
                                                    var groupp = tableDiscipllek.Rows[0]["groupp"].ToString();
                                                    var potok = tableDiscipllek.Rows[0]["potok"].ToString();

                                                    groupp = CheckTitleGroup(groupp, potok);

                                                    sheet.Cells[row, col].Value = tableDiscipllek.Rows[0]["raspred"].ToString(); ;
                                                    sheet.Cells[row, col].Style.Font.Size = sizeText;
                                                    sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                                    sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                    sheet.Cells[row, col].Style.WrapText = true;
                                                    sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                                    col++;

                                                    sheet.Cells[row, col].Value = tableDiscipllek.Rows[0]["disciplina"].ToString();
                                                    sheet.Cells[row, col].Style.Font.Size = sizeText;
                                                    sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                                    sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                    sheet.Cells[row, col].Style.WrapText = true;
                                                    sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                                    col++;

                                                    sheet.Cells[row, col].Value = tableDiscipllek.Rows[0]["kaf"].ToString();
                                                    sheet.Cells[row, col].Style.Font.Size = sizeText;
                                                    sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                                    sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                    sheet.Cells[row, col].Style.WrapText = true;
                                                    sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                                    col++;

                                                    sheet.Cells[row, col].Value = "лк";
                                                    sheet.Cells[row, col].Style.Font.Size = sizeText;
                                                    sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                                    sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                    sheet.Cells[row, col].Style.WrapText = true;
                                                    sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                                    col++;

                                                    sheet.Cells[row, col].Value = groupp;
                                                    sheet.Cells[row, col].Style.Font.Size = sizeText;
                                                    sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                                    sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                    sheet.Cells[row, col].Style.WrapText = true;
                                                    sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                                    col++;
                                                    col++;

                                                    sheet.Cells[row, col].Value = tableDiscipllek.Rows[0]["sum(students)"].ToString();
                                                    sheet.Cells[row, col].Style.Font.Size = sizeText;
                                                    sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                                    sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                    sheet.Cells[row, col].Style.WrapText = true;
                                                    sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                                    col++;

                                                    sheet.Cells[row, col].Value = tableDiscipllek.Rows[0]["SUM(lek)"].ToString();
                                                    sheet.Cells[row, col].Style.Font.Size = sizeText;
                                                    sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                                    sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                    sheet.Cells[row, col].Style.WrapText = true;
                                                    sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                                    sheet.Cells[row, 1, row, 11].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                                    sheet.Cells[row, 1, row, 11].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                                    sheet.Cells[row, 1, row, 11].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                                    sheet.Cells[row, 1, row, 11].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                                                    row++;
                                                }
                                            }
                                        }
                                        

                                        var tableDisciplpr = DB.ExecuteQuery
                                            (
                                                "SELECT raspred, disciplina,kaf, groupp,podgrupp,sum(students),potok, SUM(prakt) FROM dataconvert " +
                                                "WHERE disciplina=@disciplina AND napravl=@napravl AND semestr=@semestr GROUP BY raspred",
                                                new SqliteParameter("@semestr", semestr),
                                                new SqliteParameter("@napravl", napravl),
                                                new SqliteParameter("@disciplina", disciplina)
                                            );

                                        if (tableDisciplpr.Rows[0]["SUM(prakt)"].ToString() != "0")
                                        {

                                            col = 1;
                                            var groupp = tableDisciplpr.Rows[0]["groupp"].ToString();
                                            string pattern = @"\((\d+)\)";

                                            int sum = 0;

                                            foreach (Match match in Regex.Matches(groupp, pattern))
                                            {
                                                int value;
                                                if (int.TryParse(match.Groups[1].Value, out value))
                                                {
                                                    sum += value;
                                                }
                                            }

                                            groupp = CheckTitleGroup(groupp, "");

                                            sheet.Cells[row, col].Value = tableDisciplpr.Rows[0]["raspred"].ToString(); ;
                                            sheet.Cells[row, col].Style.Font.Size = sizeText;
                                            sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                            sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            sheet.Cells[row, col].Style.WrapText = true;
                                            sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                            col++;

                                            sheet.Cells[row, col].Value = tableDisciplpr.Rows[0]["disciplina"].ToString();
                                            sheet.Cells[row, col].Style.Font.Size = sizeText;
                                            sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                            sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            sheet.Cells[row, col].Style.WrapText = true;
                                            sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                            col++;

                                            sheet.Cells[row, col].Value = tableDisciplpr.Rows[0]["kaf"].ToString();
                                            sheet.Cells[row, col].Style.Font.Size = sizeText;
                                            sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                            sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            sheet.Cells[row, col].Style.WrapText = true;
                                            sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                            col++;

                                            sheet.Cells[row, col].Value = "пр";
                                            sheet.Cells[row, col].Style.Font.Size = sizeText;
                                            sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                            sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            sheet.Cells[row, col].Style.WrapText = true;
                                            sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                            col++;

                                            sheet.Cells[row, col].Value = groupp;
                                            sheet.Cells[row, col].Style.Font.Size = sizeText;
                                            sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                            sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            sheet.Cells[row, col].Style.WrapText = true;
                                            sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                            col++;

                                            sheet.Cells[row, col].Value = tableDisciplpr.Rows[0]["podgrupp"].ToString();
                                            sheet.Cells[row, col].Style.Font.Size = sizeText;
                                            sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                            sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            sheet.Cells[row, col].Style.WrapText = true;
                                            sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                            col++;

                                            sheet.Cells[row, col].Value = sum;
                                            sheet.Cells[row, col].Style.Font.Size = sizeText;
                                            sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                            sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            sheet.Cells[row, col].Style.WrapText = true;
                                            sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                            col++;

                                            sheet.Cells[row, col].Value = tableDisciplpr.Rows[0]["SUM(prakt)"].ToString();
                                            sheet.Cells[row, col].Style.Font.Size = sizeText;
                                            sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                            sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            sheet.Cells[row, col].Style.WrapText = true;
                                            sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                            sheet.Cells[row, 1, row, 11].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                            sheet.Cells[row, 1, row, 11].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                            sheet.Cells[row, 1, row, 11].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                            sheet.Cells[row, 1, row, 11].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                                            row++;


                                        }

                                        var tableDiscipllab = DB.ExecuteQuery
                                            (
                                                "SELECT raspred, disciplina,kaf, groupp,podgrupp,sum(students),potok, SUM(lab) FROM dataconvert " +
                                                "WHERE disciplina=@disciplina AND napravl=@napravl AND semestr=@semestr GROUP BY raspred",
                                                new SqliteParameter("@semestr", semestr),
                                                new SqliteParameter("@napravl", napravl),
                                                new SqliteParameter("@disciplina", disciplina)
                                            );

                                        if (tableDiscipllab.Rows[0]["SUM(lab)"].ToString() != "0")
                                        {

                                            col = 1;
                                            var groupp = tableDiscipllab.Rows[0]["groupp"].ToString();
                                            string pattern = @"\((\d+)\)";

                                            int sum = 0;

                                            foreach (Match match in Regex.Matches(groupp, pattern))
                                            {
                                                int value;
                                                if (int.TryParse(match.Groups[1].Value, out value))
                                                {
                                                    sum += value;
                                                }
                                            }
                                            groupp = groupp = CheckTitleGroup(groupp, "");

                                            sheet.Cells[row, col].Value = tableDiscipllab.Rows[0]["raspred"].ToString(); ;
                                            sheet.Cells[row, col].Style.Font.Size = sizeText;
                                            sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                            sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            sheet.Cells[row, col].Style.WrapText = true;
                                            sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                            col++;

                                            sheet.Cells[row, col].Value = tableDiscipllab.Rows[0]["disciplina"].ToString();
                                            sheet.Cells[row, col].Style.Font.Size = sizeText;
                                            sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                            sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            sheet.Cells[row, col].Style.WrapText = true;
                                            sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                            col++;

                                            sheet.Cells[row, col].Value = tableDiscipllab.Rows[0]["kaf"].ToString();
                                            sheet.Cells[row, col].Style.Font.Size = sizeText;
                                            sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                            sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            sheet.Cells[row, col].Style.WrapText = true;
                                            sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                            col++;

                                            sheet.Cells[row, col].Value = "лб";
                                            sheet.Cells[row, col].Style.Font.Size = sizeText;
                                            sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                            sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            sheet.Cells[row, col].Style.WrapText = true;
                                            sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                            col++;

                                            sheet.Cells[row, col].Value = groupp;
                                            sheet.Cells[row, col].Style.Font.Size = sizeText;
                                            sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                            sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            sheet.Cells[row, col].Style.WrapText = true;
                                            sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                            col++;

                                            sheet.Cells[row, col].Value = tableDiscipllab.Rows[0]["podgrupp"].ToString();
                                            sheet.Cells[row, col].Style.Font.Size = sizeText;
                                            sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                            sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            sheet.Cells[row, col].Style.WrapText = true;
                                            sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                            col++;

                                            sheet.Cells[row, col].Value = sum;
                                            sheet.Cells[row, col].Style.Font.Size = sizeText;
                                            sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                            sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            sheet.Cells[row, col].Style.WrapText = true;
                                            sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                            col++;

                                            sheet.Cells[row, col].Value = tableDiscipllab.Rows[0]["SUM(lab)"].ToString();
                                            sheet.Cells[row, col].Style.Font.Size = sizeText;
                                            sheet.Cells[row, col].Style.Font.Name = "Times New Roman";
                                            sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            sheet.Cells[row, col].Style.WrapText = true;
                                            sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                            sheet.Cells[row, 1, row, 11].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                            sheet.Cells[row, 1, row, 11].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                            sheet.Cells[row, 1, row, 11].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                            sheet.Cells[row, 1, row, 11].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                                            row++;


                                        }
                                    }
                                }
                            }
                        }
                        
                    }
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
