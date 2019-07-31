using OfficeOpenXml;
using System;
using System.IO;
using System.Text;

namespace HofCreator
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"C:\Users\pablo\OneDrive\1-Documentos\2 - Projetos\3 - Map Palhocity\Linhas Palhocity.xlsx";
            FileInfo file = new FileInfo(path);
            ExcelPackage package = new ExcelPackage(file);
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Palhocity"];
            int rows = worksheet.Dimension.Rows;

            StringBuilder sd = new StringBuilder();

            sd.AppendFormat("{0}\n", "===========================================================");
            sd.AppendFormat("{0}\n", "TERMINI BUS");
            sd.AppendFormat("{0}\n", "===========================================================");

            for (int i = 2; i <= rows; i++)
            {
                if (worksheet.Cells[$"A{i}"].Value == null) break;
                sd.AppendFormat("{0}\n", "[addterminus]");
                sd.AppendFormat("{0}\n", worksheet.Cells[$"J{i}"].Value == null ? "" : worksheet.Cells[$"J{i}"].Value.ToString());
                sd.AppendFormat("{0}\n", worksheet.Cells[$"C{i}"].Value == null ? "" : worksheet.Cells[$"C{i}"].Value.ToString());
                sd.AppendFormat("{0}\n", worksheet.Cells[$"G{i}"].Value == null ? "" : worksheet.Cells[$"G{i}"].Value.ToString());
                sd.AppendFormat("{0}\n", worksheet.Cells[$"D{i}"].Value == null ? "" : worksheet.Cells[$"D{i}"].Value.ToString());
                sd.AppendFormat("{0}\n", worksheet.Cells[$"E{i}"].Value == null ? "" : worksheet.Cells[$"E{i}"].Value.ToString());
                sd.AppendFormat("{0}\n", worksheet.Cells[$"F{i}"].Value == null ? "" : worksheet.Cells[$"F{i}"].Value.ToString());
                sd.AppendFormat("{0}\n", worksheet.Cells[$"I{i}"].Value == null ? "" : worksheet.Cells[$"I{i}"].Value.ToString());
                sd.AppendFormat("{0}\n", worksheet.Cells[$"H{i}"].Value == null ? "" : worksheet.Cells[$"H{i}"].Value.ToString());
                sd.AppendFormat("{0}\n", "-----------------------------------------------------------");

            }
            File.WriteAllText(@"C:\Users\pablo\OneDrive\1-Documentos\2 - Projetos\3 - Map Palhocity\Palhocity.hof", sd.ToString());
        }
    }
}
