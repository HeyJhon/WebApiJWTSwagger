using System;
using System.Collections.Generic;
using OfficeOpenXml;
namespace WebApplicationAPISW.Tools
{
    public static class GeneraExcelEPPlus
    {
        public static byte[] GenerarExcel<T>(this List<T> data, string nameSheetReport)
        {
            ArgumentValidator.NotNull("Lista de datos", data);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var excel = new ExcelPackage())
            {
                var workSheet = excel.Workbook.Worksheets.Add(nameSheetReport);
                workSheet.Cells[1, 1].LoadFromCollection(data, PrintHeaders: true, TableStyle: OfficeOpenXml.Table.TableStyles.Medium2);
                workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();
                return excel.GetAsByteArray();
            }

        }

        public static byte[] GenerarExcel<T>(this List<T> data, string nameSheetReport, List<Tuple<int, string>> formatos)
        {
            ArgumentValidator.NotNull("Lista de datos", data);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var excel = new ExcelPackage())
            {
                var workSheet = excel.Workbook.Worksheets.Add(nameSheetReport);
                workSheet.Cells[1, 1].LoadFromCollection(data, PrintHeaders: true, TableStyle: OfficeOpenXml.Table.TableStyles.Medium2);
                workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();

                AplicaFormatos(workSheet, data.Count, formatos);

                return excel.GetAsByteArray();
            }

        }

        private static void AplicaFormatos(ExcelWorksheet workSheet, int dataCount, List<Tuple<int, string>> formatos)
        {
            foreach (var formato in formatos)
            {
                int fila = 2;
                int columna = formato.Item1;
                string formatoCelda = formato.Item2;
                for (int i = 0; i < dataCount; i++)
                {
                    workSheet.Cells[fila, columna].Style.Numberformat.Format = formatoCelda;
                    fila++;
                }
            }
        }
        internal static class ArgumentValidator
        {
            public static void NotNull(string name, [ValidatedNotNull] object value)
            {
                if (value == null)
                {
                    throw new ArgumentNullException(name);
                }
            }
        }

        [AttributeUsage(AttributeTargets.All, Inherited = false, AllowMultiple = true)]
        internal sealed class ValidatedNotNullAttribute: Attribute { }
    }
}