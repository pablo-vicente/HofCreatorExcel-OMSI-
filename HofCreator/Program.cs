using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace HofCreator
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Insira o caminho para o Arquivo:");
            string path = Console.ReadLine();
            Console.WriteLine("Insira o caminho para salvar o HOF:");
            string path2 = Console.ReadLine();
            path2 = @"C:\Users\pablo\OneDrive\1-Documentos\2 - Projetos\3 - Map Palhocity";


            path = @"C:\Users\pablo\OneDrive\1-Documentos\2 - Projetos\3 - Map Palhocity\Linhas Palhocity.xlsx";
            FileInfo file = new FileInfo(path);
            ExcelPackage package = new ExcelPackage(file);

            StringBuilder sb = new StringBuilder();
            sb.Append(AdicionarInformacoesMapa(package));
            sb.Append(AdicionarTerminus(package));
            sb.Append(AdicionarBusStops(package));
            sb.Append(AdicionarTripsRoutas(package));
            File.WriteAllText(Path.Combine(path2, "Palhocity.hof"), sb.ToString(), Encoding.Default);
            Process.Start("explorer.exe", path2);
        }

        public static StringBuilder AdicionarInformacoesMapa(ExcelPackage package)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
            int rows = worksheet.Dimension.Rows;

            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("{0}\r\n", "#####TEMPLATE FILE#####");
            sb.AppendFormat("{0}\r\n\n\n", "SD200_SD202_Template.xml");
            sb.AppendFormat("{0}\r\n", "[name]");
            sb.AppendFormat("{0}\r\n\n\n", worksheet.Cells["C3"].Value.ToString());
            sb.AppendFormat("{0}\r\n", "[servicetrip]");
            sb.AppendFormat("{0}\r\n\n\n", worksheet.Cells["C4"].Value.ToString());
            sb.AppendFormat("{0}\r\n", "stringcount_terminus");
            sb.AppendFormat("{0}\r\n\n\n", "6");
            sb.AppendFormat("{0}\r\n", "stringcount_busstop");
            sb.AppendFormat("{0}\r\n\n\n", "4");

            return sb;
        }

        public static StringBuilder AdicionarTerminus(ExcelPackage package)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            int rows = worksheet.Dimension.Rows;

            StringBuilder sd = new StringBuilder();

            sd.AppendFormat("{0}\r\n", "===========================================================");
            sd.AppendFormat("{0}\r\n", "======================= TERMINI BUS =======================");
            sd.AppendFormat("{0}\r\n\n", "===========================================================");

            for (int i = 2; i <= rows; i++)
            {
                var linebus = worksheet.Cells[$"A{i}"].Value;
                if (linebus == null) break; ;

                sd.AppendFormat("{0}\r\n", linebus.ToString().Equals("0") ? "[addterminus_allexit]" : "[addterminus]");

                sd.AppendFormat("{0}\r\n", worksheet.Cells[$"J{i}"].Value == null ? "" : worksheet.Cells[$"J{i}"].Value.ToString());
                sd.AppendFormat("{0}\r\n", worksheet.Cells[$"C{i}"].Value == null ? "" : worksheet.Cells[$"C{i}"].Value.ToString());
                sd.AppendFormat("{0}\r\n", worksheet.Cells[$"G{i}"].Value == null ? "" : worksheet.Cells[$"G{i}"].Value.ToString());
                sd.AppendFormat("{0}\r\n", worksheet.Cells[$"D{i}"].Value == null ? "" : worksheet.Cells[$"D{i}"].Value.ToString());
                sd.AppendFormat("{0}\r\n", worksheet.Cells[$"E{i}"].Value == null ? "" : worksheet.Cells[$"E{i}"].Value.ToString());
                sd.AppendFormat("{0}\r\n", worksheet.Cells[$"F{i}"].Value == null ? "" : worksheet.Cells[$"F{i}"].Value.ToString());
                sd.AppendFormat("{0}\r\n", worksheet.Cells[$"I{i}"].Value == null ? "" : worksheet.Cells[$"I{i}"].Value.ToString());
                sd.AppendFormat("{0}\r\n", worksheet.Cells[$"H{i}"].Value == null ? "" : worksheet.Cells[$"H{i}"].Value.ToString());
                sd.AppendFormat("{0}\r\n", "-----------------------------------------------------------");
            }
            return sd;
        }

        public static StringBuilder AdicionarBusStops(ExcelPackage package)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            int rows = worksheet.Dimension.Rows;

            StringBuilder sd = new StringBuilder();

            sd.AppendFormat("\r\n{0}\r\n", "===========================================================");
            sd.AppendFormat("{0}\r\n", "======================== BUS STOPS ========================");
            sd.AppendFormat("{0}\r\n\n", "===========================================================");

            List<string> busStops = new List<string>();

            for (int i = 2; i <= rows; i++)
            {
                var linebus = worksheet.Cells[$"A{i}"].Value;
                if (linebus == null) break; ;

                if (!linebus.ToString().Equals("0"))
                {
                    string firtBusStop = worksheet.Cells[$"P{i}"].Value == null ? "" : worksheet.Cells[$"P{i}"].Value.ToString();
                    string lastBusStop = worksheet.Cells[$"Q{i}"].Value == null ? "" : worksheet.Cells[$"Q{i}"].Value.ToString();

                    busStops.Add(firtBusStop);
                    busStops.Add(lastBusStop);
                }
            }
            foreach (string busstop in busStops.Distinct().OrderBy(x => x.ToString()))
            {
                sd.AppendFormat("{0}\r\n", "[addbusstop]");
                for (int j = 1; j <= 5; j++)
                {
                    sd.AppendFormat("{0}\r\n", busstop);
                }
                sd.AppendFormat("{0}\r\n", "-----------------------------------------------------------");
            }
            return sd;
        }

        public static StringBuilder AdicionarTripsRoutas(ExcelPackage package)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            int rows = worksheet.Dimension.Rows;

            StringBuilder sd = new StringBuilder();

            sd.AppendFormat("\r\n{0}\r\n", "===========================================================");
            sd.AppendFormat("{0}\r\n", "========================= ROUTES ==========================");
            sd.AppendFormat("{0}\r\n\n", "===========================================================");

            for (int i = 2; i <= rows; i++)
            {
                var linebus = worksheet.Cells[$"A{i}"].Value;
                if (linebus == null) break; ;

                if (!linebus.ToString().Equals("0"))
                {
                    string firtBusStop = worksheet.Cells[$"P{i}"].Value == null ? "" : worksheet.Cells[$"P{i}"].Value.ToString();
                    string lastBusStop = worksheet.Cells[$"Q{i}"].Value == null ? "" : worksheet.Cells[$"Q{i}"].Value.ToString();
                    var nLine = linebus.ToString().Replace("-", "");
                    var nRoute = $"{0}{worksheet.Cells[$"B{i}"].Value.ToString()}";

                    sd.AppendFormat("{0}\r\n", "-----------------------------------------------------------");
                    sd.AppendFormat("{0}\r\n", $"Line {nLine} Route {nRoute} : {firtBusStop} => {lastBusStop}");
                    sd.AppendFormat("{0}\r\n", "-----------------------------------------------------------");
                    sd.AppendFormat("{0}\r\n", "[infosystem_trip]");
                    sd.AppendFormat("{0}\r\n", $"{nLine}{nRoute}");
                    sd.AppendFormat("{0}\r\n", $"{firtBusStop} => {lastBusStop} 1");
                    sd.AppendFormat("{0}\r\n", worksheet.Cells[$"J{i}"].Value == null ? "" : worksheet.Cells[$"J{i}"].Value.ToString());
                    sd.AppendFormat("{0}\r\n\n", "via");
                    sd.AppendFormat("{0}\r\n", "[infosystem_busstop_list]");
                    sd.AppendFormat("{0}\r\n", "2");
                    sd.AppendFormat("{0}\r\n", firtBusStop);
                    sd.AppendFormat("{0}\r\n", lastBusStop);
                }
            }
            return sd;
        }
    }
}
