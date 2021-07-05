using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;

namespace WebApplicationAPISW.Tools
{
    public class GenerarExcelOpenXml
    {

        public GenerarExcelOpenXml()
        {
            this.Cabecera = new Dictionary<string, string>();
        }
        protected Dictionary<string, string> Cabecera { get; set; }
        public byte[] GenerarExcel<T>(List<T> datos, string nombreReporte, string propiedadesValidas = "", char separadores = ',')
        {
            Type typeObj = typeof(T);
            if (!string.IsNullOrEmpty(propiedadesValidas))
                this.ObtenerCabecerasDeCadena(typeObj, propiedadesValidas, separadores);


            try
            {
                using (MemoryStream stream = new MemoryStream())
                {
                    using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
                    {
                        // Add a WorkbookPart to the document.
                        WorkbookPart workbookPart = document.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();
                        // Add a WorksheetPart to the WorkbookPart.
                        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                        SheetData sheetData = new SheetData();
                        worksheetPart.Worksheet = new Worksheet(sheetData);
                        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                        Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = nombreReporte };

                        sheets.Append(sheet);


                        if (!this.Cabecera.Any())
                        {
                            PropertyInfo[] properties = typeObj.GetProperties();
                            foreach (PropertyInfo propertyInfo in properties)
                            {
                                string cabecera = ObtenerNombre(propertyInfo);
                                if (!this.Cabecera.ContainsKey(propertyInfo.Name))
                                    this.Cabecera.Add(propertyInfo.Name, cabecera);
                            }
                        }


                        int rowIndex = 1;

                        Row rowHeader = new Row() { RowIndex = Convert.ToUInt32(rowIndex) };
                        int colIndex = 1;
                        foreach (KeyValuePair<string, string> item in this.Cabecera)
                        {
                            string colName = ColumnIndexToColumnLetter(colIndex) + rowIndex;
                            Cell header = new Cell() { CellReference = colName, CellValue = new CellValue(item.Value), DataType = CellValues.String };
                            rowHeader.Append(header);
                            colIndex++;
                        }
                        sheetData.Append(rowHeader);

                        rowIndex++;
                        if (datos != null)
                        {
                            foreach (T item in datos)
                            {
                                PropertyInfo[] properties = item.GetType().GetProperties();
                                Row row = new Row() { RowIndex = Convert.ToUInt32(rowIndex) };
                                colIndex = 1;
                                foreach (var property in properties)
                                {
                                    if (!this.Cabecera.ContainsKey(property.Name))
                                        continue;
                                    string colName = ColumnIndexToColumnLetter(colIndex) + rowIndex;
                                    string value = property.GetValue(item) != null ? property.GetValue(item).ToString() : "";
                                    Cell celda = new Cell() { CellReference = colName, CellValue = new CellValue(value), DataType = CellValues.String };
                                    row.Append(celda);
                                    colIndex++;
                                }
                                sheetData.Append(row);
                                rowIndex++;
                            }
                        }
                        workbookPart.Workbook.Save();
                    }

                    return stream.ToArray();
                }

            }
            catch (ArgumentNullException) { }
            catch (IOException) { }
            catch (ArgumentException) { }
            return null;
        }

        private static string ObtenerNombre(PropertyInfo propertyInfo)
        {
            Attribute attri = propertyInfo.GetCustomAttribute(typeof(DisplayAttribute), false);
            string cabecera = propertyInfo.Name;
            if (attri != null)
            {
                var displayName = (DisplayAttribute)attri;
                cabecera = displayName.Name;
            }

            return cabecera;
        }

        protected void ObtenerCabecerasDeCadena(Type typeObj, string propiedadesValidas = "", char separadores = ',')
        {
            try
            {
                this.Cabecera = new Dictionary<string, string>();
                if (!string.IsNullOrEmpty(propiedadesValidas))
                {

                    PropertyInfo[] properties = typeObj.GetProperties();
                    foreach (PropertyInfo propertyInfo in properties)
                    {
                        string cabecera = ObtenerNombre(propertyInfo);
                        if (!this.Cabecera.ContainsKey(propertyInfo.Name))
                            this.Cabecera.Add(propertyInfo.Name, cabecera);
                    }
                }
            }
            catch (ArgumentNullException) { }
            catch (ArgumentException) { }
        }

        private string ColumnIndexToColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }
    }

    public class RegistroExcel
    {
        public object Valor { get; set; }
    }
}