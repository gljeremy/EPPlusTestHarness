using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;

namespace SpreadsheetHarness
{
    class Program
    {
        public class PirateDto
        {
            public int Id { get; set; }
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public int Plunders { get; set; }
            public int Ships { get; set; }
        }


        static void Main(string[] args)
        {
            var pirateDtos = GetPirateDtos().ToList();

            BuildExcelEmpty(pirateDtos);

            //BuildExcelManual(pirateDtos);
            
            //BuildExcelDataTable(pirateDtos);

            //BuildExcelCollection(pirateDtos);
        }

        private static void BuildExcelEmpty(List<PirateDto> pirateDtos)
        {
            DateTime beforeTime, afterTime;

            using (var package = new ExcelPackage(new FileInfo(@"C:\Data\GID.xlsm"), true))
            {
                beforeTime = DateTime.Now;

                

                byte[] afterBytes = package.GetAsByteArray();

                afterTime = DateTime.Now;

                File.WriteAllBytes(@"C:\Data\AfterEmpty.xlsm", afterBytes);
            }

            Console.WriteLine("Total for empty: [" + (afterTime - beforeTime) + "]");
        }

        private static void BuildExcelCollection(IEnumerable<PirateDto> pirateDtos)
        {
            DateTime beforeTime, afterTime;

            using (var package = new ExcelPackage(new FileInfo(@"C:\Data\Book1.xlsm"), false))
            {
                beforeTime = DateTime.Now;

                package.Workbook.Worksheets["Sheet2"].Cells["A1"].LoadFromCollection(pirateDtos);

                byte[] afterBytes = package.GetAsByteArray();

                afterTime = DateTime.Now;

                File.WriteAllBytes(@"C:\Data\AfterDataCollection.xlsm", afterBytes);
            }

            Console.WriteLine("Total for collection: [" + (afterTime - beforeTime) + "]");
        }

        private static void BuildExcelDataTable(List<PirateDto> pirateDtos)
        {
            var dtPirates = pirateDtos.ToDataTable();

            DateTime beforeTime, afterTime;

            using (var package = new ExcelPackage(new FileInfo(@"C:\Data\Book1.xlsm"), false))
            {
                beforeTime = DateTime.Now;

                package.Workbook.Worksheets["Sheet2"].Cells["A1"].LoadFromDataTable(dtPirates, true);

                byte[] afterBytes = package.GetAsByteArray();

                afterTime = DateTime.Now;

                File.WriteAllBytes(@"C:\Data\AfterDataLoad.xlsm", afterBytes);
            }

            Console.WriteLine("Total for data table: [" + (afterTime - beforeTime) + "]");
        }

        private static void BuildExcelManual(List<PirateDto> pirateDtos)
        {
            DateTime beforeTime, afterTime;

            using (var package = new ExcelPackage(new FileInfo(@"C:\Data\Book1.xlsm"), true))
            {
                beforeTime = DateTime.Now;

                for (int i = 0; i < pirateDtos.Count(); i++)
                {
                    package.Workbook.Worksheets["Sheet2"].Cells[i + 1, 1].Value = pirateDtos[i].Id;
                    package.Workbook.Worksheets["Sheet2"].Cells[i + 1, 2].Value = pirateDtos[i].FirstName;
                    package.Workbook.Worksheets["Sheet2"].Cells[i + 1, 3].Value = pirateDtos[i].LastName;
                    package.Workbook.Worksheets["Sheet2"].Cells[i + 1, 4].Value = pirateDtos[i].Plunders;
                    package.Workbook.Worksheets["Sheet2"].Cells[i + 1, 5].Value = pirateDtos[i].Ships;
                    package.Workbook.Worksheets["Sheet2"].Cells[i + 1, 6].CreateArrayFormula("SUM(B2:C5)");
                }

                byte[] afterBytes = package.GetAsByteArray();

                afterTime = DateTime.Now;

                File.WriteAllBytes(@"C:\Data\AfterManualLoad.xlsm", afterBytes);
            }

            Console.WriteLine("Total for manual: [" + (afterTime - beforeTime) + "]");
        }

        static IEnumerable<PirateDto> GetPirateDtos()
        {
            foreach (var id in Enumerable.Range(1, 1000000))
                yield return new PirateDto() { Id = id, FirstName = "Mr", LastName = "Pirate", Plunders = 1, Ships = 2 };
        }
    }

    public static class DataTableExtensions
    {
        public static DataTable ToDataTable<T>(this IList<T> data)
        {
            PropertyDescriptorCollection properties =
                TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;
        }
    }
}
