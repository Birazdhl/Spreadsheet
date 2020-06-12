using CreateExcelFile;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Quantum.Web.Framework.Document
{
    public static class ExcelHelper
    {
        /// <summary>
        /// Saves to server location (NOT to CLIENT). Use .xlsx as extension
        /// </summary>
        /// <typeparam name="T">Type of List</typeparam>
        /// <param name="records">List of Records</param>
        /// <param name="fileName">Full path (with file name) to where to save the file to (server path, NOT CLIENT). eg:e:\\excelfile.xlsx</param>
        public static void ListToExcel<T>(IList<T> records, string fileName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                WriteExcel(records, document, null, null);
            }
        }
        public static void ListToExcel<T>(IList<T> records, string fileName, string ExcludeColumn)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                WriteExcel(records, document, ExcludeColumn, null);
            }
        }
        public static void ListToExcel<T>(IList<T> records, string fileName, string ExcludeColumn, string wKsheetName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                WriteExcel(records, document, ExcludeColumn, wKsheetName);
            }
        }
        /// <summary>
        /// Directly downloads to the client browser
        /// </summary>
        /// <typeparam name="T">Type of List</typeparam>
        /// <param name="records">List of Records</param>
        /// <returns></returns>
        public static System.IO.MemoryStream ListToExcelStream<T>(IList<T> records)
        {
            System.IO.MemoryStream stream = new System.IO.MemoryStream();
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, true))
            {
                WriteExcel(records, document, null, null);
            }
            return stream;
        }

        #region " Private Heavyweight functions "
        private static void WriteExcel<T>(IList<T> records, SpreadsheetDocument spreadsheet, string ExcludeColumn, string WksheetName)
        {
            //  Create the Excel file contents.  This function is used when creating an Excel file either writing 
            //  to a file, or writing to a MemoryStream.
            spreadsheet.AddWorkbookPart();
            spreadsheet.WorkbookPart.Workbook = new Workbook();

            //  My thanks to James Miera for the following line of code (which prevents crashes in Excel 2010)
            spreadsheet.WorkbookPart.Workbook.Append(new BookViews(new WorkbookView()));

            //  If we don't add a "WorkbookStylesPart", OLEDB will refuse to connect to this .xlsx file !
            WorkbookStylesPart workbookStylesPart = spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>("rIdStyles");
            Stylesheet stylesheet = new Stylesheet();
            workbookStylesPart.Stylesheet = stylesheet;
            //  For each worksheet you want to create
            string workSheetID = "rId" + 1;
            string worksheetName = string.Empty;
            if (WksheetName == "Contract Import")
            {
                worksheetName = "Contract Import";
            }
            else
            {
                worksheetName = "Worksheet" + 1;
            }
            WorksheetPart newWorksheetPart = spreadsheet.WorkbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet();

            // create sheet data
            newWorksheetPart.Worksheet.AppendChild(new SheetData());

            // save worksheet
            WriteToWorkSheet(records, newWorksheetPart, ExcludeColumn);
            newWorksheetPart.Worksheet.Save();

            // create the worksheet to workbook relation
            // if (worksheetNumber == 1)
            spreadsheet.WorkbookPart.Workbook.AppendChild(new Sheets());

            spreadsheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().AppendChild(new Sheet()
            {
                Id = spreadsheet.WorkbookPart.GetIdOfPart(newWorksheetPart),
                SheetId = 1,
                Name = worksheetName
            });

            spreadsheet.WorkbookPart.Workbook.Save();
        }

        private static void WriteToWorkSheet<T>(IList<T> records, WorksheetPart worksheetPart, string ExcludeColumn)
        {
            var worksheet = worksheetPart.Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();


            var props = typeof(T).GetProperties();//GetMembers().Where(x => x.MemberType == MemberTypes.Property);

            if (ExcludeColumn != null)
            {
                var ExcludeColumnList = ExcludeColumn.Split(',');
                foreach (var item in ExcludeColumnList)
                {
                    props = props.Where(x => x.Name != item).ToArray();

                }
            }

            //  Create a Header Row in our Excel file, containing one header for each Column of data in our DataTable.
            //
            //  We'll also create an array, showing which type each column of data is (Text or Numeric), so when we come to write the actual
            //  cells of data, we'll know if to write Text values or Numeric cell values.
            int numberOfColumns = props.Count();
            bool[] IsNumericColumn = new bool[numberOfColumns];

            string[] excelColumnNames = new string[numberOfColumns];
            for (int n = 0; n < numberOfColumns; n++)
                excelColumnNames[n] = GetExcelColumnName(n);

            var colOptions = new List<ColOption>();
            //
            //  Create the Header row in our Excel Worksheet
            //
            uint rowIndex = 1;

            var headerRow = new Row { RowIndex = rowIndex, };  // add a row at the top of spreadsheet
            sheetData.Append(headerRow);
            int colIndex = 0;
            foreach (var prop in props)
            {
                var colOption = new ColOption();
                colOption.Type = GetCellValueType(prop);
                colOption.ExcelColumnName = excelColumnNames[colIndex];
                var browsableProp = prop.GetCustomAttribute(typeof(BrowsableAttribute), true) as BrowsableAttribute;
                // If the property is browsable then only show it in excel
                if (browsableProp == null || browsableProp.Browsable == true)
                {
                    string columnName = "";
                    // Check for different column Name
                    var displayNameProp = prop.GetCustomAttribute(typeof(DisplayNameAttribute), true) as DisplayNameAttribute;
                    if (displayNameProp != null)
                    {
                        columnName = displayNameProp.DisplayName;
                    }
                    else
                    {
                        columnName = prop.Name;
                    }
                    // Check for formatting
                    var formatStringProp = prop.GetCustomAttribute(typeof(DisplayFormatAttribute), true) as DisplayFormatAttribute;
                    if (formatStringProp != null)
                    {
                        colOption.DataFormatString = formatStringProp.DataFormatString;
                    }
                    AppendTextCell(excelColumnNames[colIndex] + "1", columnName, headerRow);
                    colIndex++;
                    colOption.Browsable = true;
                }
                else
                {
                    colOption.Browsable = false;
                }
                colOptions.Add(colOption);
            }
            //
            //  Now, step through each row of data in our DataTable...
            //
            foreach (var record in records)
            {
                // ...create a new row, and append a set of this row's data to it.
                ++rowIndex;
                var newExcelRow = new Row { RowIndex = rowIndex };  // add a row at the top of spreadsheet
                sheetData.Append(newExcelRow);
                colIndex = 0;
                foreach (var prop in props)
                {
                    var colOption = colOptions[colIndex];

                    if (colOption.Browsable)
                    {
                        object valueObj = prop.GetValue(record);
                        string cellValue = valueObj == null ? "" : (string.IsNullOrEmpty(colOption.DataFormatString) == true ? valueObj.ToString() : string.Format(colOption.DataFormatString, valueObj));
                        //string cellValue = "";
                        //  For text cells, just write the input data straight out to the Excel file.
                        AppendCell(colOption.ExcelColumnName + rowIndex.ToString(), cellValue, newExcelRow, colOption.Type);
                    }
                    colIndex++;
                }
            }
        }
        #endregion

        #region " Private Helper functions "
        private static string hexvaluesToRemove = @"[^\x09\x0A\x0D\x20-\xD7FF\xE000-\xFFFD\x10000-x10FFFF]";
        private static void AppendCell(string cellReference, string cellStringValue, Row excelRow, CellValues dataType)
        {
            //  Add a new Excel Cell to our Row 
            Cell cell = new Cell() { CellReference = cellReference, DataType = dataType };
            CellValue cellValue = new CellValue();

            if (string.IsNullOrEmpty(cellStringValue) == false)
            {
                cellStringValue = Regex.Replace(cellStringValue, hexvaluesToRemove, "");
            }
            cellValue.Text = cellStringValue;
            cell.Append(cellValue);
            excelRow.Append(cell);
        }

        private static void AppendTextCell(string cellReference, string cellStringValue, Row excelRow)
        {
            //  Add a new Excel Cell to our Row 
            Cell cell = new Cell() { CellReference = cellReference, DataType = CellValues.String };
            CellValue cellValue = new CellValue();
            if (string.IsNullOrEmpty(cellStringValue) == false)
            {
                cellStringValue = Regex.Replace(cellStringValue, hexvaluesToRemove, "");
            }
            cellValue.Text = cellStringValue;
            cell.Append(cellValue);
            excelRow.Append(cell);
        }
        private static Type[] numberTypes = new[] { typeof(int), typeof(Int16), typeof(long), typeof(int?), typeof(Int16?),
            typeof(long?), typeof(double), typeof(decimal), typeof(double?), typeof(decimal?), typeof(float), typeof(float?) };
        private static CellValues GetCellValueType(PropertyInfo property)
        {
            if (property.PropertyType == typeof(string))
            {
                return CellValues.String;
            }
            else if (property.PropertyType == typeof(bool) || property.PropertyType == typeof(bool?))
            {
                //  return CellValues.Boolean; // Shows error when opening excel file
                return CellValues.String;
            }
            else if (property.PropertyType == typeof(DateTime) || property.PropertyType == typeof(DateTime?))
            {
                return CellValues.String;
            }
            else if (numberTypes.Contains(property.PropertyType))
            {
                return CellValues.Number;
            }
            // If not found anything

            return CellValues.String;
        }
        private static string GetExcelColumnName(int columnIndex, int offset = 0)
        {
            //  Convert a zero-based column index into an Excel column reference  (A, B, C.. Y, Y, AA, AB, AC... AY, AZ, B1, B2..)
            //
            //  eg  GetExcelColumnName(0) should return "A"
            //      GetExcelColumnName(1) should return "B"
            //      GetExcelColumnName(25) should return "Z"
            //      GetExcelColumnName(26) should return "AA"
            //      GetExcelColumnName(27) should return "AB"
            //      ..etc..
            //
            if ((columnIndex - offset) < 26)
                return ((char)('A' + (columnIndex - offset))).ToString();

            char firstChar = (char)('A' + ((columnIndex - offset) / 26) - 1);
            char secondChar = (char)('A' + ((columnIndex - offset) % 26));

            return string.Format("{0}{1}", firstChar, secondChar);
        }
        private class ColOption
        {
            public CellValues Type = CellValues.String;
            public bool Browsable = true;
            public string ExcelColumnName = "";
            public string DataFormatString = "";
        }
        #endregion
        //sushin: need to replace Datatable.
        public static DataTable ReadExcelSheetDataTable(string fileName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                ////Read the first Sheet from Excel file.
                //Sheet sheet = spreadsheetDocument.WorkbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
                ////Get the Worksheet instance.
                //Worksheet worksheet = (spreadsheetDocument.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                ////Fetch all the rows present in the Worksheet.
                //IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();
                DataTable dt = new DataTable();
                //Loop through the Worksheet rows.
                foreach (Row row in rows)
                {
                    //Use the first row to add columns to DataTable.
                    if (row.RowIndex.Value == 1)
                    {
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            dt.Columns.Add(GetValue(spreadsheetDocument, cell));
                        }
                    }
                    else
                    {
                        //Add rows to DataTable.
                        dt.Rows.Add();
                        int i = 0;
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = GetValue(spreadsheetDocument, cell);
                            i++;
                        }
                    }
                }
                return dt;
            }
        }
        public static DataTable GetDataTableFromSpreadsheet(string MyExcelStream, bool ReadOnly, string SheetName, out string ErrorMessage)
        {
            DataTable dt = new DataTable();
            using (SpreadsheetDocument sDoc = SpreadsheetDocument.Open(MyExcelStream, ReadOnly))
            {
                WorkbookPart workbookPart = sDoc.WorkbookPart;
                IEnumerable<Sheet> sheets = sDoc.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                if (!string.IsNullOrEmpty(SheetName))
                {
                    sheets = sheets.Where(x => x.Name == SheetName);
                }

                ErrorMessage = "";

                if (sheets.Count() < 1)
                {
                    ErrorMessage = SheetName;
                    return dt;
                }

                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)sDoc.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;

                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();

                foreach (Cell cell in rows.ElementAt(0))
                {
                    dt.Columns.Add(GetCellValue(sDoc, cell));
                }

                foreach (Row row in rows) //this will also include your header row...
                {
                    var totalColumns = rows.ElementAt(0).Count();

                    DataRow tempRow = dt.NewRow();

                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                    {
                        Cell cell = row.Descendants<Cell>().ElementAt(i);
                        int actualCellIndex = CellReferenceToIndex(cell);
                        if (actualCellIndex < totalColumns)  //Rakshya 9283-6: this is done to read data of only expected no of columns.
                        {
                            tempRow[actualCellIndex] = GetCellValue(sDoc, cell);
                        }

                    }

                    dt.Rows.Add(tempRow);
                }
            }
            dt.Rows.RemoveAt(0);
            return dt;
        }
        private static int CellReferenceToIndex(Cell cell)
        {
            int index = 0;
            string reference = cell.CellReference.ToString().ToUpper();

            reference = reference.ToUpper();
            for (int ix = 0; ix < reference.Length && reference[ix] >= 'A'; ix++)
            {
                index = (index * 26) + ((int)reference[ix] - 64);
            }

            return index - 1;
        }
        public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            if (cell.CellValue == null)
            {
                return "";
            }
            string value = cell.CellValue.InnerXml;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText.Trim();
            }
            else
            {
                if (cell.StyleIndex == null)
                {

                    return value.Trim();
                }
                var cellformat = GetCellFormat(cell);

                ////Swoyuj: Cell Format commented for Date values >>>> Added Hardcoded check for Start Date And End Date in Equinox Contract Import in bsImportCOntroller
                /////https: //stackoverflow.com/questions/11781210/c-sharp-open-xml-2-0-numberformatid-range

                //if ((cellformat.NumberFormatId >= 14 && cellformat.NumberFormatId <= 22) ||
                //            (cellformat.NumberFormatId >= 165 && cellformat.NumberFormatId <= 180) ||
                //                cellformat.NumberFormatId == 278 || cellformat.NumberFormatId == 185 || cellformat.NumberFormatId == 196 ||
                //                cellformat.NumberFormatId == 217 || cellformat.NumberFormatId == 326) // Dates
                //{
                //    return DateTime.FromOADate(Convert.ToDouble(value)).ToShortDateString();
                //}

                return value.Trim();
            }
        }
        private static CellFormat GetCellFormat(Cell cell)
        {
            Worksheet workSheet = cell.Ancestors<Worksheet>().FirstOrDefault();
            SpreadsheetDocument doc = workSheet.WorksheetPart.OpenXmlPackage as SpreadsheetDocument;
            WorkbookPart workbookPart = doc.WorkbookPart;
            int styleIndex = (int)cell.StyleIndex.Value;
            CellFormat cellFormat = (CellFormat)workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ElementAt(styleIndex);
            return cellFormat;
        }
        public static DataTable ReadExcelSheetDataTable(Stream fileStream)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileStream, false))
            {
                ////Read the first Sheet from Excel file.
                //Sheet sheet = spreadsheetDocument.WorkbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
                ////Get the Worksheet instance.
                //Worksheet worksheet = (spreadsheetDocument.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                ////Fetch all the rows present in the Worksheet.
                //IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();
                DataTable dt = new DataTable();
                //Loop through the Worksheet rows.
                foreach (Row row in rows)
                {
                    //Use the first row to add columns to DataTable.
                    if (row.RowIndex.Value == 1)
                    {
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            dt.Columns.Add(GetValue(spreadsheetDocument, cell));
                        }
                    }
                    else
                    {
                        //Add rows to DataTable.
                        dt.Rows.Add();
                        int i = 0;
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = GetValue(spreadsheetDocument, cell);
                            i++;
                        }
                    }
                }
                return dt;
            }
        }
        private static string GetValue(SpreadsheetDocument doc, Cell cell)
        {
            string value = cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
            }
            return value;
        }
        public static List<object> ReadExcelSheet(string fileName)
        {
            var listColumnHeaders = new List<dynamic>();
            var listData = new List<dynamic>();
            var listReturnData = new List<dynamic>();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();
                //Loop through the Worksheet rows.
                foreach (Row row in rows)
                {
                    //Use the first row to add columns Header. 
                    if (row.RowIndex.Value == 1)
                    {
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            listColumnHeaders.Add(GetValue(spreadsheetDocument, cell));
                        }
                    }
                    else
                    {
                        try
                        {
                            if (listColumnHeaders != null)
                            {
                                int i = 0;
                                foreach (Cell cell in row.Descendants<Cell>())
                                {
                                    string KeyName = listColumnHeaders[i];
                                    string ValueName = GetValue(spreadsheetDocument, cell);
                                    var data = new { KeyName, ValueName };
                                    listData.Add(data);
                                    i++;
                                }
                                string newJson = JsonConvert.SerializeObject(listData);
                                listReturnData.Add(newJson);
                                listData.Clear();
                            }
                        }
                        catch (Exception e)
                        {

                        }
                    }
                }
                return listReturnData;
            }
        }
        public static bool CheckSequenceNumber(string sequenceNumber)
        {
            int n;
            bool isNumeric = int.TryParse(sequenceNumber, out n);
            if (isNumeric)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public static bool ValidateHeaders(string headerText)
        {
            bool IsValidate = false;

            switch (headerText.Trim())
            {
                case "SiteID":
                    IsValidate = true;
                    break;
                case "VendorSiteID":
                    IsValidate = true;
                    break;
                case "SequenceNumber":
                    IsValidate = true;
                    break;

            }


            return IsValidate;
        }
        #region Validate Excel
        public static bool ValidateRequiredField(System.Data.DataRow dr, out ArrayList ErrMsgList, out int ErroneousColumnIndex, int k)
        {
            ErrMsgList = new ArrayList();
            ErroneousColumnIndex = -1;
            if (dr["SiteID"].ToString().Trim().Length == 0 && dr["VendorSiteID"].ToString().Trim().Length == 0)
            {
                ErrMsgList.Add(" Either SiteID or VendorSiteID is required at row " + k);
                ErroneousColumnIndex = 1;
                return false;
            }
            else
            {
                return true;
            }
        }
        #endregion Validate Excel
        public static DataTable ToDataTable<T>(this IEnumerable<T> collection)
        {
            DataTable dt = new DataTable("DataTable");
            Type t = typeof(T);
            PropertyInfo[] pia = t.GetProperties();

            //Inspect the properties and create the columns in the DataTable
            foreach (PropertyInfo pi in pia)
            {
                Type ColumnType = pi.PropertyType;
                if ((ColumnType.IsGenericType))
                {
                    ColumnType = ColumnType.GetGenericArguments()[0];
                }
                dt.Columns.Add(pi.Name, ColumnType);
            }

            //Populate the data table
            foreach (T item in collection)
            {
                DataRow dr = dt.NewRow();
                dr.BeginEdit();
                foreach (PropertyInfo pi in pia)
                {
                    if (pi.GetValue(item, null) != null)
                    {
                        dr[pi.Name] = pi.GetValue(item, null);
                    }
                }
                dr.EndEdit();
                dt.Rows.Add(dr);
            }
            return dt;
        }

        public static void ListToExcelInMultipleSheets(DataSet ds, string Savepath)
        {
            using (var workbook = SpreadsheetDocument.Create(Savepath, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = workbook.AddWorkbookPart();
                workbook.WorkbookPart.Workbook = new Workbook();
                workbook.WorkbookPart.Workbook.Sheets = new Sheets();

                uint sheetId = 1;
                int count = 0;

                foreach (DataTable table in ds.Tables)
                {
                    var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    sheetPart.Worksheet = new Worksheet(sheetData);

                    Columns lstColumns = workbook.WorkbookPart.Workbook.GetFirstChild<Columns>();
                    if(ds.DataSetName == "ContractBillingExport" && table.Columns.Contains("ID"))
                    {
                        table.Columns.Remove("ID");
                    }
                    if (ds.DataSetName == "ContractBillingExport")
                    {
                        string tblName = table.TableName;
                        ChangetheColumnWidth(lstColumns, sheetPart, "ContractBillingExport",0);
                    }

                    Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                    string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                    if (sheets.Elements<Sheet>().Count() > 0)
                    {
                        sheetId =
                            sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                    }

                    if (count == 0 && ds.DataSetName == "ContractBillingExport")
                    {
                        WorkbookStylesPart wspContractImport = workbookPart.AddNewPart<WorkbookStylesPart>();
                        wspContractImport.Stylesheet = GenerateStylesheetDefault();
                        wspContractImport.Stylesheet.Save();
                        count = count + 1;
                    }

                    Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                    sheets.Append(sheet);

                    Row headerRow = new Row();

                    List<String> columns = new List<string>();
                    foreach (DataColumn column in table.Columns)
                    {
                        var  totalcol = table.Columns.Count;
                        columns.Add(column.ColumnName);

                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;
                        if (ds.DataSetName == "ContractBillingExport")
                        {
                        Cell cellStyleIndx = CellStyleIndex(column.ColumnName, "CalledFromHeader");
                           
                        cell.StyleIndex = cellStyleIndx.StyleIndex;
                        for (int i = 8; i<=totalcol; i++)
                        {   
                            uint number = (uint)(int)i;
                            ChangetheColumnWidth(lstColumns, sheetPart, "ContractBillingExportDynamic", number);
                        }
                        }
                       
                        if (column.ColumnName == "Column1" || column.ColumnName == "Column2" || column.ColumnName == "Column3" || column.ColumnName == "Column4" || column.ColumnName == "Column5" || column.ColumnName == "Column6" || column.ColumnName == "Column7" || column.ColumnName == "Column8" || column.ColumnName == "Column9" || column.ColumnName == "Column10" || column.ColumnName == "Column11" || column.ColumnName == "Column12" || column.ColumnName == "Column13" || column.ColumnName == "Column14" || column.ColumnName == "Column15" || column.ColumnName == "Column16")
                        {
                            cell.CellValue = new CellValue("");
                        }
                        else
                        {
                            cell.CellValue = new CellValue(column.ColumnName);
                        }
                        headerRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(headerRow);

                    foreach (DataRow dsrow in table.Rows)
                    {
                        Row newRow = new Row();
                        foreach (String col in columns)
                        {
                            Cell cell = new Cell();
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(dsrow[col].ToString());
                            if((!col.Contains("Contract#") && !col.Contains("Campaign Name") && !col.Contains("AGENCY") && !col.Contains("AGENCY COMPANY #") && !col.Contains("ADVERTISER") && !col.Contains("ADVERTISER COMPANY #") && !col.Contains("DIVISION") && !col.Contains("OFFICE CODE") && !col.Contains("Market") && !col.Contains("Description") && !col.Contains("Vendor Unit ID") && !col.Contains("Vendor Name") && !col.Contains("Start Date") && !col.Contains("End Date") && !col.Contains("Line Item")) && ds.DataSetName == "ContractBillingExport")
                            {
                                Cell cellStyleIndx = CellStyleIndex("toberightalignedforContractBillingExport", "CalledFromHeader");
                                cell.StyleIndex = cellStyleIndx.StyleIndex;
                            }                         
                            newRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(newRow);
                    }
                }
                workbook.WorkbookPart.Workbook.Save();
            }
        }
        /// <summary>
        /// This is being added by sanjay for taskid 10212-51
        /// Here you can add different color to different sheet
        /// Please do changes on this method for appending multiple sheet and this should be dynamic with the column style
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="Savepath"></param>

        public static void AppendMultipleSheetWhileDoingExportToExcel(DataSet ds, string Savepath)
        {
        
            using (var workbook = SpreadsheetDocument.Create(Savepath, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = workbook.AddWorkbookPart();
                workbook.WorkbookPart.Workbook = new Workbook();
                workbook.WorkbookPart.Workbook.Sheets = new Sheets();
                uint sheetId = 1;
                int count = 0;

                foreach (DataTable table in ds.Tables)
                {
                    var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    sheetPart.Worksheet = new Worksheet(sheetData);

                    Columns lstColumns = workbook.WorkbookPart.Workbook.GetFirstChild<Columns>();
                    if (table.TableName == "Pulldown Column Data" || table.TableName == "AE Checklist" || table.TableName == "New Organization Request Form" || table.TableName == "Revision Log" || table.TableName == "Revision Report" || table.TableName =="Data List" || table.TableName== "Formatted Invoice")
                    {
                        string tblName = table.TableName;
                        ChangetheColumnWidth(lstColumns, sheetPart, tblName,0);
                    }

                    Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                    string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);
                    uint rowIndex = 1;

                    int numberofCols = table.Columns.Count;
                    bool[] IsNumericColumn = new bool[numberofCols];
                    string[] excelColumnNames = new string[numberofCols];
                    for (int n = 0; n < numberofCols; n++)
                        excelColumnNames[n] = GetExcelColumnName(n);

                    if (count == 0)
                    {
                        WorkbookStylesPart wspContractImport = workbookPart.AddNewPart<WorkbookStylesPart>();
                        wspContractImport.Stylesheet = GenerateStylesheetDefault();
                        wspContractImport.Stylesheet.Save();
                        count = count + 1;
                    }

                    if (sheets.Elements<Sheet>().Count() > 0)
                    {
                        sheetId =
                            sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                    }

                    Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                    sheets.Append(sheet);
                    Row headerRow = new Row();
                    List<String> columns = new List<string>();
                    int colIndex = 0;
                    foreach (DataColumn column in table.Columns)
                    {
                        // Below cell.StyleIndex always start from 1 you can add font as well as cellFormat or Color fill but adding it on Font and CellFormat it always start from index 0 but for cell.StyleIndex it start from 1 index 
                        columns.Add(column.ColumnName);
                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(column.ColumnName);
                        cell.CellReference = excelColumnNames[colIndex] + "1";

                        Cell cellStyleIndx = CellStyleIndex(column.ColumnName, "CalledFromHeader");
                        //cell.CellValue = new CellValue(cellStyleIndx.CellValue.InnerText);
                        cell.StyleIndex = cellStyleIndx.StyleIndex;
                        headerRow.AppendChild(cell);

                        colIndex++;
                    }
                    sheetData.AppendChild(headerRow);

                    foreach (DataRow dsrow in table.Rows)
                    {
                        int innerColIndex = 0;
                        rowIndex++;
                        Row newRow = new Row();
                        foreach (String col in columns)
                        {
                            Stylesheet stylesheet1 = new Stylesheet();
                            Cell cell = new Cell();
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(dsrow[col].ToString());

                            cell.CellReference = excelColumnNames[innerColIndex] + rowIndex.ToString();
                            if (table.TableName == "Pulldown Column Data" || table.TableName == "Revision Log" || table.TableName == "New Organization Request Form" || table.TableName == "AE Checklist")
                            {
                                Cell cellStyleIndx = CellStyleIndex(dsrow[col].ToString(), string.Empty);
                                cell.StyleIndex = cellStyleIndx.StyleIndex;
                            }
                            if(table.TableName == "Revision Report")
                            {
                             

                                Cell cellStyleIndx = CellStyleIndex(dsrow[col].ToString(), "RevisionReport");
                               
                                cell.StyleIndex = cellStyleIndx.StyleIndex;

                            }
                            if (table.TableName == "Formatted Invoice")
                            {
                              
                                Cell cellStyleIndx = CellStyleIndex(dsrow[col].ToString(), "FormattedInvoice");
                                cell.StyleIndex = cellStyleIndx.StyleIndex;
                           

                            }
                            if (dsrow[col].ToString() == "BorderTop" || dsrow[col].ToString() == "BottomRightBorder" || dsrow[col].ToString() == "BorderRight" || dsrow[col].ToString()== "BorderBottom" || dsrow[col].ToString() == "TopRightBorder")
                            {
                                cell.CellValue = new CellValue("");
                            }
                            newRow.AppendChild(cell);
                            innerColIndex++;
                        }

                        sheetData.AppendChild(newRow);              
                    }
                }
                workbook.WorkbookPart.Workbook.Save();
            }
        }

        public static void ExportToExcelWithImageHeader(DataSet ds, string Savepath, string sImagePath)
        {
            using (var workbook = SpreadsheetDocument.Create(Savepath, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = workbook.AddWorkbookPart();
                workbook.WorkbookPart.Workbook = new Workbook();
                MergeCells mergeCells = new MergeCells();
                FileVersion fv = new FileVersion();
                fv.ApplicationName = "Microsoft Office Excel";
                workbook.WorkbookPart.Workbook.Sheets = new Sheets();
                uint sheetId = 1;
                int count = 0;
                int head = 7;

                foreach (DataTable table in ds.Tables)
                {
                    var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    sheetPart.Worksheet = new Worksheet(sheetData);

                    Columns lstColumns = workbook.WorkbookPart.Workbook.GetFirstChild<Columns>();

                    Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                    string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);
                    uint rowIndex = 1;

                    int numberofCols = table.Columns.Count;
                    bool[] IsNumericColumn = new bool[numberofCols];
                    string[] excelColumnNames = new string[numberofCols];
                    for (int n = 0; n < numberofCols; n++)
                        excelColumnNames[n] = GetExcelColumnName(n);

                    if (count == 0)
                    {
                        WorkbookStylesPart wspContractImport = workbookPart.AddNewPart<WorkbookStylesPart>();
                        wspContractImport.Stylesheet = GenerateStylesheetDefault();
                        wspContractImport.Stylesheet.Save();
                        count = count + 1;
                    }

                    DrawingsPart dp = sheetPart.AddNewPart<DrawingsPart>();
                    ImagePart imgp = dp.AddImagePart(ImagePartType.Png, sheetPart.GetIdOfPart(dp));
                    using (FileStream fs = new FileStream(sImagePath, FileMode.Open))
                    {
                        imgp.FeedData(fs);
                    }

                    DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties nvdp = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties();
                    nvdp.Id = 1025;
                    nvdp.Name = "Picture 1";
                    nvdp.Description = "deltamedialogo";
                    DocumentFormat.OpenXml.Drawing.PictureLocks picLocks = new DocumentFormat.OpenXml.Drawing.PictureLocks();
                    picLocks.NoChangeAspect = true;
                    picLocks.NoChangeArrowheads = true;
                    DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureDrawingProperties nvpdp = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureDrawingProperties();
                    nvpdp.PictureLocks = picLocks;
                    DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureProperties nvpp = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureProperties();
                    nvpp.NonVisualDrawingProperties = nvdp;
                    nvpp.NonVisualPictureDrawingProperties = nvpdp;

                    DocumentFormat.OpenXml.Drawing.Stretch stretch = new DocumentFormat.OpenXml.Drawing.Stretch();
                    stretch.FillRectangle = new DocumentFormat.OpenXml.Drawing.FillRectangle();

                    DocumentFormat.OpenXml.Drawing.Spreadsheet.BlipFill blipFill = new DocumentFormat.OpenXml.Drawing.Spreadsheet.BlipFill();
                    DocumentFormat.OpenXml.Drawing.Blip blip = new DocumentFormat.OpenXml.Drawing.Blip();
                    blip.Embed = dp.GetIdOfPart(imgp);
                    blip.CompressionState = DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print;
                    blipFill.Blip = blip;
                    blipFill.SourceRectangle = new DocumentFormat.OpenXml.Drawing.SourceRectangle();
                    blipFill.Append(stretch);

                    DocumentFormat.OpenXml.Drawing.Transform2D t2d = new DocumentFormat.OpenXml.Drawing.Transform2D();
                    DocumentFormat.OpenXml.Drawing.Offset offset = new DocumentFormat.OpenXml.Drawing.Offset();
                    offset.X = 0;
                    offset.Y = 0;
                    t2d.Offset = offset;
                    System.Drawing.Bitmap bm = new System.Drawing.Bitmap(sImagePath);

                    DocumentFormat.OpenXml.Drawing.Extents extents = new DocumentFormat.OpenXml.Drawing.Extents();
                    extents.Cx = (long)bm.Width * (long)((float)914400 / bm.HorizontalResolution);
                    extents.Cy = (long)bm.Height * (long)((float)914400 / bm.VerticalResolution);
                    bm.Dispose();
                    t2d.Extents = extents;
                    DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties sp = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties();
                    sp.BlackWhiteMode = DocumentFormat.OpenXml.Drawing.BlackWhiteModeValues.Auto;
                    sp.Transform2D = t2d;
                    DocumentFormat.OpenXml.Drawing.PresetGeometry prstGeom = new DocumentFormat.OpenXml.Drawing.PresetGeometry();
                    prstGeom.Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle;
                    prstGeom.AdjustValueList = new DocumentFormat.OpenXml.Drawing.AdjustValueList();
                    sp.Append(prstGeom);
                    sp.Append(new DocumentFormat.OpenXml.Drawing.NoFill());

                    DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture picture = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture();
                    picture.NonVisualPictureProperties = nvpp;
                    picture.BlipFill = blipFill;
                    picture.ShapeProperties = sp;

                    DocumentFormat.OpenXml.Drawing.Spreadsheet.Position pos = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Position();
                    pos.X = 0;
                    pos.Y = 0;
                    DocumentFormat.OpenXml.Drawing.Spreadsheet.Extent ext = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Extent();
                    ext.Cx = extents.Cx;
                    ext.Cy = extents.Cy;
                    DocumentFormat.OpenXml.Drawing.Spreadsheet.AbsoluteAnchor anchor = new DocumentFormat.OpenXml.Drawing.Spreadsheet.AbsoluteAnchor();
                    anchor.Position = pos;
                    anchor.Extent = ext;
                    anchor.Append(picture);
                    anchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ClientData());

                    DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing wsd = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
                    wsd.Append(anchor);
                    Drawing drawing = new Drawing();
                    drawing.Id = dp.GetIdOfPart(imgp);

                    wsd.Save(dp);
                    Random rand = new Random();

                    sheetPart.Worksheet.Append(drawing);

                    if (sheets.Elements<Sheet>().Count() > 0)
                    {
                        sheetId =
                            sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                    }

                    Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                    sheets.Append(sheet);
                    Row headerRow = new Row();
                    List<String> columns = new List<string>();
                    int colIndex = 0;
                    foreach (DataColumn column in table.Columns)
                    {
                        // Below cell.StyleIndex always start from 1 you can add font as well as cellFormat or Color fill but adding it on Font and CellFormat it always start from index 0 but for cell.StyleIndex it start from 1 index 
                        columns.Add(column.ColumnName);
                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(column.ColumnName);
                        cell.CellReference = excelColumnNames[colIndex] + "1";

                        Cell cellStyleIndx = CellStyleIndex(column.ColumnName, "CalledFromHeader");
                        //cell.CellValue = new CellValue(cellStyleIndx.CellValue.InnerText);
                        cell.StyleIndex = cellStyleIndx.StyleIndex;
                        headerRow.AppendChild(cell);

                        colIndex++;
                    }
                    sheetData.AppendChild(headerRow);

                    foreach (DataRow dsrow in table.Rows)
                    {
                        int innerColIndex = 0;
                        rowIndex++;
                        Row newRow = new Row();
                        foreach (String col in columns)
                        {
                            Stylesheet stylesheet1 = new Stylesheet();
                            Cell cell = new Cell();
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(dsrow[col].ToString());

                            if (dsrow[col].ToString() == "POSTING ORDER : INSTRUCTIONS & CONFIRMATION OF RECEIPT OF MATERIALS")
                            {
                                ChangetheColumnWidthOfWorkOrderTable(lstColumns, sheetPart, "POSTING ORDER : INSTRUCTIONS & CONFIRMATION OF RECEIPT OF MATERIALS");
                            }
                            if (dsrow[col].ToString() == "TakeDown ORDER : INSTRUCTIONS & CONFIRMATION OF RECEIPT OF MATERIALS")
                            {
                                ChangetheColumnWidthOfWorkOrderTable(lstColumns, sheetPart, "TakeDown ORDER : INSTRUCTIONS & CONFIRMATION OF RECEIPT OF MATERIALS");
                            }

                            cell.CellReference = excelColumnNames[innerColIndex] + rowIndex.ToString();
                            if (table.TableName == "Work Order Report")
                            {
                                Cell cellStyleIndx = CellStyleIndex(dsrow[col].ToString(), string.Empty);
                                cell.StyleIndex = cellStyleIndx.StyleIndex;
                            }

                        if (table.TableName == "Work Order Report")
                            {
                                if (dsrow[col].ToString() == "POSTER: 10% MUST HAVE APPROACH AND CLOSE-UP SHOTS - PHOTO OF EACH CREATIVE")
                                {
                                    for (int i = 0; i < 3; i++)
                                    {
                                        if (i == 0)
                                        {
                                            MergeReportDataCell(dsrow[col].ToString(), head, cell.CellReference, excelColumnNames[innerColIndex], rowIndex, mergeCells);
                                        }
                                        else if (i == 1)
                                        {
                                            StringValue NewcellReference = excelColumnNames[innerColIndex] + (rowIndex + 1);
                                            MergeReportDataCell(dsrow[col].ToString(), head, NewcellReference, excelColumnNames[innerColIndex], rowIndex + 1, mergeCells);

                                        }
                                        else
                                        {
                                            StringValue NewcellReference = excelColumnNames[innerColIndex] + (rowIndex + 2);
                                            MergeReportDataCell(dsrow[col].ToString(), head, NewcellReference, excelColumnNames[innerColIndex], rowIndex + 2, mergeCells);
                                        }
                                    }
                                }

                                if (dsrow[col].ToString() != "POSTER: 10% MUST HAVE APPROACH AND CLOSE-UP SHOTS - PHOTO OF EACH CREATIVE")
                                {
                                    MergeReportDataCell(dsrow[col].ToString(), head, cell.CellReference, excelColumnNames[innerColIndex], rowIndex, mergeCells);
                                }
                                head = head + 1;
                            }

                            newRow.AppendChild(cell);
                            innerColIndex++;
                        }

                        sheetData.AppendChild(newRow);
                    }
                    if (table.TableName == "Work Order Report" && mergeCells != null)
                    {
                        sheetPart.Worksheet.InsertAfter(mergeCells, sheetPart.Worksheet.Elements<SheetData>().First());
                    }
                }
                workbook.WorkbookPart.Workbook.Save();
            }
        }
        public static void MergeReportDataCell(string cellNameWorkOrder,int head, StringValue cellReference, string columnIndex,uint rowIndex, MergeCells mergeCells)
        {

            if (cellNameWorkOrder == "POSTER: 10% MUST HAVE APPROACH AND CLOSE-UP SHOTS - PHOTO OF EACH CREATIVE" || cellNameWorkOrder == "BULLETINS: 100% CLOSE-UP AND APPROACH OF EACH UNIT" || cellNameWorkOrder == "PLEASE SIGN AND EMAIL BACK" || cellNameWorkOrder == "TO CONFIRM THAT THE OUTDOOR CO. ACKNOWLEDGES THE PROPER DESIGN TO BE POSTED ON THE DATE SHOWN ABOVE AND ALSO THE CREATIVE REMOVAL STATUS." 
                || cellNameWorkOrder == "Do not post prior to posting date without confirmation that it is okay to do so." || cellNameWorkOrder == "TO CONFIRM THAT THE OUTDOOR CO. ACKNOWLEDGES THE PROPER DESIGN  TO BE REMOVED AFTER THE TAKEDOWN DATE SHOWN ABOVE AND " || cellNameWorkOrder == "ALSO TO PROVIDE THE ACTUAL TAKEDOWN DATE. IF MATERIALS NEED TO BE REMOVED PRIOR TO TAKEDOWN DATE SHOWN ABOVE , CONTACT US IMMEDIATELY." 
                || cellNameWorkOrder == "Do not remove prior to takedown date without confirmation that it is okay to do so." || cellNameWorkOrder == "Authorized Signature" || cellNameWorkOrder == "Date"/* || cellNameWorkOrder == "ADD CREATIVE IMAGE HERE" */|| cellNameWorkOrder == "STANDARD PHOTO REQUEST:" || cellNameWorkOrder == "IF MATERIALS ARE NOT DELIVERED 5 DAYS PRIOR TO POSTING ," 
                || cellNameWorkOrder == "IF MATERIALS NEED TO BE REMOVED PRIOR TO TAKEDOWN DATE SHOWN ABOVE ," || cellNameWorkOrder == "If you have any questions, please feel free to contact us." || head < 14)
            {
                if (mergeCells == null)
                    mergeCells = new MergeCells();

                var cellAddress = cellReference;
                var cellAddressTwo = "I" + rowIndex.ToString();
                if(/*cellNameWorkOrder == "ADD CREATIVE IMAGE HERE" || */cellNameWorkOrder == "IF MATERIALS ARE NOT DELIVERED 5 DAYS PRIOR TO POSTING ," || cellNameWorkOrder == "IF MATERIALS NEED TO BE REMOVED PRIOR TO TAKEDOWN DATE SHOWN ABOVE ," || cellNameWorkOrder == "STANDARD PHOTO REQUEST:" || cellNameWorkOrder == "POSTER: 10% MUST HAVE APPROACH AND CLOSE-UP SHOTS - PHOTO OF EACH CREATIVE" || cellNameWorkOrder == "BULLETINS: 100% CLOSE-UP AND APPROACH OF EACH UNIT" || cellNameWorkOrder == "If you have any questions, please feel free to contact us.")
                {
                    cellAddress = cellReference;
                    cellAddressTwo = "E" + rowIndex.ToString();
                }
                if(cellNameWorkOrder == "IF MATERIALS ARE NOT DELIVERED 5 DAYS PRIOR TO POSTING ," || cellNameWorkOrder == "IF MATERIALS NEED TO BE REMOVED PRIOR TO TAKEDOWN DATE SHOWN ABOVE ,")
                {
                    cellAddress = cellReference; ;
                    cellAddressTwo = "D" + rowIndex.ToString();
                }

                if(cellNameWorkOrder == "Do not post prior to posting date without confirmation that it is okay to do so." || cellNameWorkOrder == "Do not remove prior to takedown date without confirmation that it is okay to do so.")
                {
                    cellAddress = cellReference;
                    cellAddressTwo = "C" + rowIndex.ToString();
                }
                
                if (head < 15)
                {
                    cellAddress = "A" + (head + 1).ToString();
                    cellAddressTwo = "D" + (head + 1).ToString();
                }

                    mergeCells.Append(new MergeCell() { Reference = new StringValue(cellAddress + ":" + cellAddressTwo) });
            }
        }
       
        public static Stylesheet GenerateStylesheetDefault()
        {
            
            Stylesheet styleSheet = null;

            Fonts fonts = new Fonts(
                new Font( // Index 0 - default
                     new FontSize() { Val = 11 },
                     new FontName() { Val = "Calibri" },
                     new FontFamilyNumbering() { Val = 2 }
                ),
                new Font( // Index 1 - header
                    new FontSize() { Val = 11 },
                    new Bold(),
                    new Color() { Rgb = "FFFFFF" }
                    ),
                new Font( // Index 2 - header
                    new FontSize() { Val = 11 },
                     new Bold() { Val = true },
                     new Italic() { Val = true }
                ),
                new Font( // Index 3 - header
                    new FontSize() { Val = 11 },
                     new Color() { Rgb = "000000" },
                     new FontName() { Val = "Calibri" },
                     new Bold() { Val = true }
                ),
                 new Font( // Index 4 - header
                    new FontSize() { Val = 11 },
                     new Bold() { Val = true },
                     new Underline() { Val = UnderlineValues.Single }
                ),
                new Font( // Index 5 - header
                    new FontSize() { Val = 11 },
                     new Color() { Rgb = "ff0202" },
                     new FontName() { Val = "Calibri" }
                )
                );

            Fills fills = new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - default
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value =TranslateForeground(System.Drawing.ColorTranslator.FromHtml("#16365C")) } })
                    { PatternType = PatternValues.Solid }), // Index 2 - header
                     new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = TranslateForeground(System.Drawing.ColorTranslator.FromHtml("#FF0000")) } })  //FF0000 light red
                     { PatternType = PatternValues.Solid }),
                      new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = TranslateForeground(System.Drawing.ColorTranslator.FromHtml("#FABF8F")) } })  // FABF8F brown
                      { PatternType = PatternValues.Solid }),
                       new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = TranslateForeground(System.Drawing.ColorTranslator.FromHtml("#00B0F0")) } })   // 00B0F0 light blue
                       { PatternType = PatternValues.Solid }),
                       new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = TranslateForeground(System.Drawing.ColorTranslator.FromHtml("#92D050")) } })   // 00B0F0 light Gray
                       { PatternType = PatternValues.Solid }),
                       new Fill(new PatternFill (new ForegroundColor { Rgb = new HexBinaryValue() { Value = TranslateForeground(System.Drawing.ColorTranslator.FromHtml("#FFFFFF")) } })   
                       { PatternType = PatternValues.None }),
                       new Fill(new PatternFill (new ForegroundColor { Rgb = new HexBinaryValue() { Value = TranslateForeground(System.Drawing.ColorTranslator.FromHtml("#FFFFFF")) } })   
                       { PatternType = PatternValues.None })

                );

            Borders borders = new Borders(
                    new Border(), // index 0 default
                    new Border( // index 1 black border
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder()),

                    // From Index 2 to Index 11 is for TaskID_11708-2-3 //ReportLauncher/WorkOrder/ExportReport

                    new Border( // index 2 black border
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thick },
                        new RightBorder(),
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues. Thick },
                        new BottomBorder(),
                        new DiagonalBorder()),
                    new Border( // index 3 black border
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thick },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thick },
                        new TopBorder(),
                        new BottomBorder(),
                        new DiagonalBorder()),
                   new Border( // index 4 black border
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thick },
                        new BottomBorder(),
                        new DiagonalBorder()),
                   new Border( // index 5 black border
                        new LeftBorder(),
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thick },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thick },
                        new BottomBorder(),
                        new DiagonalBorder()),
                   new Border( // index 6 black border
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thick },
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thick },
                        new DiagonalBorder()),
                   new Border( // index 7 black border
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thick },
                        new DiagonalBorder()),
                   new Border( // index 8 black border
                        new LeftBorder(),
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thick },
                        new TopBorder(),
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thick },
                        new DiagonalBorder()),
                   new Border( // index 9 black border
                        new LeftBorder(),
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thick },
                        new TopBorder(),
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thick },
                        new DiagonalBorder()),
                   new Border( // index 10 black border
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thick },
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(),
                        new DiagonalBorder()),
                   new Border( // index 11 black border
                        new LeftBorder(),
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thick },
                        new TopBorder(),
                        new BottomBorder(),
                        new DiagonalBorder()),
                    new Border( // index 12 black border
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thick },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thick },
                        new DiagonalBorder()
                       )

                );
            
            CellFormats cellFormats = new CellFormats(
                // This below index (1,2,3,4..) are used to display font fill border eg see above i.e (cell.StyleIndex = 8U) for Bold and Italic font and no any color fill
                    new CellFormat(), // default
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 0, ApplyBorder = true }, // body index 1
                    new CellFormat { FontId = 1, FillId = 1, BorderId = 0, ApplyFill = true }, // header index 2
                    new CellFormat { FontId = 1, FillId = 2, BorderId = 0, ApplyFill = true }, // header index 3 
                    new CellFormat { FontId = 1, FillId = 3, BorderId = 0, ApplyFill = true }, // header index 4
                    new CellFormat { FontId = 1, FillId = 4, BorderId = 0, ApplyFill = true }, // header index 5
                    new CellFormat { FontId = 1, FillId = 5, BorderId = 0, ApplyFill = true }, // header index 6
                    new CellFormat { FontId = 1, FillId = 6, BorderId = 0, ApplyFill = true }, // header index 7
                    new CellFormat { FontId = 2, FillId = 7, BorderId = 0, ApplyFill = true }, // header index 8
                    new CellFormat { FontId = 3, FillId = 8, BorderId = 0, ApplyFill = true }, // header index 9
                    new CellFormat { FontId = 4, FillId = 8, BorderId = 0, ApplyFill = true }, // header index 10
                    new CellFormat { FontId = 5, FillId = 8, BorderId = 0, ApplyFill = true }, // header index 11

                    // From Index 12 to Index 21 is for TaskID_11708-2-3 //ReportLauncher/WorkOrder/ExportReport

                    new CellFormat { FontId = 0, FillId = 0, BorderId = 2, ApplyFill = true }, // header index 12 
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 3, ApplyFill = true }, // header index 13 
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 4, ApplyFill = true }, // header index 14 
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 5, ApplyFill = true }, // header index 15 
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 6, ApplyFill = true }, // header index 16 
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 7, ApplyFill = true }, // header index 17 
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 8, ApplyFill = true }, // header index 18 
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 9, ApplyFill = true }, // header index 19
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 10, ApplyFill = true }, // header index 20
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 11, ApplyFill = true }, // header index 21
                    new CellFormat { FontId = 3, FillId = 8, BorderId = 4, ApplyFill = true }, // header index 22
                    new CellFormat { FontId = 3, FillId = 8, BorderId = 11, ApplyFill = true }, //header index 23
                  
                    new CellFormat( new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Center }) { FontId = 0, FillId = 0, BorderId = 0, ApplyAlignment = true }, // header index 24
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Center }) { FontId = 0, FillId = 8, BorderId = 0, NumberFormatId = 2, ApplyFill = true, ApplyNumberFormat = true }, // header index 25
                   // new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }) { FontId = 3, FillId = 8, BorderId = 12, ApplyAlignment = true }  //header index 26
                    new CellFormat { FontId = 3, FillId = 8, BorderId = 12, ApplyFill = true } //header index 26

                );

            styleSheet = new Stylesheet(fonts, fills, borders, cellFormats);

            return styleSheet;
        }
        public static Cell CellStyleIndex(string cellName,string columnCalledFrom)
        {
            Cell cell = new Cell();
            if (columnCalledFrom == "CalledFromHeader")
            {
                //if (cellName == "Column1" || cellName == "Column2" || cellName == "Column3" || cellName == "Column4" || cellName == "Column5" || cellName == "Column6" || cellName == "Column7" || cellName == "Column8" || cellName == "Column9" || cellName == "Column10" || cellName == "Column11" || cellName == "Column12" || cellName == "Column13" || cellName == "Column14" || cellName == "Column15" || cellName == "Column16")
                //{
                //    cell.CellValue = new CellValue("");
                //}
                //else
                //{
                //    cell.CellValue = new CellValue(cellName);
                //}
                if (cellName == "MarketName-DMA" || cellName == "Country" || cellName == "State")
                {
                    cell.StyleIndex = 7U;
                }
                else if (cellName == "A/E")
                {
                    cell.StyleIndex = 5U;
                }
                else if (cellName == "Clients/BillTos/Agencies")
                {
                    cell.StyleIndex = 4U;
                }
                else if (cellName == "VENDORS" || cellName == "Contract Type" || cellName == "Demographics" || cellName == "Billing Term" || cellName == "Line Type"
                         || cellName == "MM MediaType" || cellName == "Standard/Digital" || cellName == "MM Read" || cellName == "MM Facing")
                {
                    cell.StyleIndex = 6U;
                }
                else if (cellName == "Please use below checklist to confirm the necessary steps for contracting have been completed."
                      || cellName == "Please use below form to request new vendor or client records be added to the system.")
                {
                    cell.StyleIndex = 9U;
                }
                else if (cellName == " " || cellName == "  " || cellName == "   " || cellName == "    " || cellName == "     "
                        || cellName == "      " || cellName == "       " || cellName == "        " || cellName == "         "
                        || cellName == "          " || cellName == "PeriodTerm")
                {
                    cell.StyleIndex = 1U;
                }               
                else if (cellName == "toberightalignedforContractBillingExport")
                {
                    cell.StyleIndex = 25U;
                }

                else
                {
                    cell.StyleIndex = 3U;
                }
            }
            else if (columnCalledFrom == "RevisionReport")
            {
                if (cellName == "Revision Report" || cellName == "Revision Notes: " || cellName == "Invoice" || cellName == "Market" || cellName == "Date" || cellName == "Reference number" || cellName == "Campaign dates" || cellName == "Amount due")
                {
                    cell.StyleIndex = 9U;
                }
                else if (cellName.Contains("Contract#") || cellName.Contains("BorderTop") || cellName.Contains("Last saved date"))
                {
                    cell.StyleIndex = 22U;
                }
                else if (cellName.Contains("BorderBottom"))
                {
                    cell.StyleIndex = 17U;
                }
                else if(cellName.Contains("BorderRight") || cellName == "Display" || cellName == "Production" || cellName == "Install        " || cellName == "Ship" || cellName == "Commission")
                {
                    cell.StyleIndex = 21U;
                }
                else if(cellName.Contains("Type"))
                {
                    cell.StyleIndex = 23U;
                }
                else if(cellName.Contains("TopRightBorder"))
                {
                    cell.StyleIndex = 15U;
                }
                else if(cellName.Contains("BottomRightBorder"))
                {
                    cell.StyleIndex = 19U;
                }
                else if(cellName.Contains("$") || cellName.Contains("%"))
                {
                    cell.StyleIndex = 24U;
                }
            }
            else if(columnCalledFrom== "FormattedInvoice")
            {
                if (cellName == "DMA" || cellName == "MARKET" || cellName == "VENDOR" || cellName == "UNIT SIZE" || cellName == "MEDIA TYPE" || cellName == "Vendor UNIT NUMBER" || cellName == "LOCATION DESCRIPTION" || cellName == "FACING" || cellName=="RATE")
                {
                    cell.StyleIndex = 9U;
                }
                else if (cellName.Contains("BorderTop"))
                {
                    cell.StyleIndex = 22U;
                }
                else if (cellName.Contains("BorderBottom"))
                {
                    cell.StyleIndex = 17U;
                }
                else if (cellName.Contains("BorderTopBottom") || cellName.Contains("AMOUNT") || cellName.Contains("POSTING DATES / DESCRIPTION:"))
                {
                    cell.StyleIndex = 24U;
                }
                   
              
            }
            else 
            {
                
                 if (cellName == "BorderTop")
                {
                    cell.StyleIndex = 14U;
                }
                else if (cellName == "BorderTopRight")
                {
                    cell.StyleIndex = 15U;
                }
          
                else if (cellName == "BorderRight")
                {
                    cell.StyleIndex = 21U;
                }
           
                else if (cellName == "BorderBottom")
                {
                    cell.StyleIndex = 17U;
                }
                else if (cellName == "BorderBottomRight")
                {
                    cell.StyleIndex = 19U;
                }  

            }

            return cell;
        }

        /// <summary>
        /// Added By sanjay for Managing the Width of Cell 
        /// Min = 1, Max = 1 ==> Apply this to column 1 (A)
        /// Min = 2, Max = 2 ==> Apply this to column 2 (B)
        /// Width = 25 ==> Set the width to 25
        /// CustomWidth = true ==> Tell Excel to use the custom width
        /// </summary>
        /// <param name="cols"></param>
        /// <param name="sheetPart"></param>
        /// <param name="tableName"></param>
        public static void ChangetheColumnWidth(Columns cols, WorksheetPart sheetPart,string tableName, uint number)
        {
            Columns lstColumns = cols;
            Boolean needToInsertColumns = false;
            if (lstColumns == null)
            {
                lstColumns = new Columns();
                needToInsertColumns = true;
            }
            if (tableName == "Pulldown Column Data")
            {
                lstColumns.Append(new Column() { Min = 1, Max = 1, Width = 19, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 2, Max = 2, Width = 40, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 3, Max = 3, Width = 33, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 4, Max = 4, Width = 14, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 5, Max = 5, Width = 35, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 6, Max = 6, Width = 15, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 7, Max = 7, Width = 14, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 8, Max = 8, Width = 20, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 9, Max = 9, Width = 16, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 10, Max = 10, Width = 16, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 11, Max = 11, Width = 16, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 12, Max = 12, Width = 36, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 13, Max = 13, Width = 26, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 14, Max = 14, Width = 18, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 15, Max = 15, Width = 12, CustomWidth = true });
            }
            if (tableName == "AE Checklist")
            {
                lstColumns.Append(new Column() { Min = 1, Max = 1, Width = 144, CustomWidth = true });
            }
            if (tableName == "New Organization Request Form")
            {
                lstColumns.Append(new Column() { Min = 1, Max = 1, Width = 80, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 2, Max = 2, Width = 48, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 3, Max = 3, Width = 50, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 4, Max = 4, Width = 45, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 5, Max = 5, Width = 14, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 6, Max = 6, Width = 14, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 7, Max = 7, Width = 35, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 8, Max = 8, Width = 42, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 9, Max = 9, Width = 25, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 10, Max = 10, Width = 30, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 11, Max = 11, Width = 14, CustomWidth = true });
            }

            if (tableName == "Revision Log")
            {
                lstColumns.Append(new Column() { Min = 1, Max = 1, Width = 13, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 2, Max = 2, Width = 20, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 3, Max = 3, Width = 70, CustomWidth = true });
            }

            if (tableName == "Work Order Report")
            {
                lstColumns.Append(new Column() { Min = 1, Max = 1, Width = 43, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 2, Max = 2, Width = 14, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 3, Max = 3, Width = 17, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 4, Max = 4, Width = 41, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 5, Max = 5, Width = 36, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 6, Max = 6, Width = 16, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 7, Max = 7, Width = 16, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 8, Max = 8, Width = 30, CustomWidth = true });
            }
            if (tableName == "ContractBillingExport")
            {
                lstColumns.Append(new Column() { Min = 1, Max = 1, Width = 10, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 2, Max = 2, Width = 15, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 3, Max = 3, Width = 15, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 4, Max = 4, Width = 19, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 5, Max = 5, Width = 28, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 6, Max = 6, Width = 22, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 7, Max = 7, Width = 22, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 8, Max = 8, Width = 12, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 8, Max = 8, Width = 13, CustomWidth = true });
            }
            if (tableName == "Revision Report")
            {
                lstColumns.Append(new Column() { Min = 1, Max = 1, Width = 22, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 2, Max = 2, Width = 25, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 3, Max = 3, Width = 20, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 4, Max = 4, Width = 21, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 5, Max = 5, Width = 25, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 6, Max = 6, Width = 22, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 7, Max = 7, Width = 12, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 8, Max = 8, Width = 15, CustomWidth = true });
            }
            if (tableName == "ContractBillingExportDynamic")
            {
                lstColumns.Append(new Column() { Min = number, Max = number, Width = 25, CustomWidth = true });
            }
            if (tableName == "Data List")
            {
                lstColumns.Append(new Column() { Min = 1, Max = 1, Width = 48, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 2, Max = 2, Width = 26, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 3, Max = 3, Width = 16, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 4, Max = 4, Width = 20, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 5, Max = 5, Width = 17, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 6, Max = 6, Width = 15, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 7, Max = 7, Width = 20, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 8, Max = 8, Width = 15, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 9, Max = 9, Width = 37, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 10, Max = 10, Width = 27, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 11, Max = 11, Width = 14, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 12, Max = 12, Width = 20, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 13, Max = 13, Width = 11, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 14, Max = 14, Width = 16, CustomWidth = true });
            }
            if(tableName== "Formatted Invoice")
            {
                lstColumns.Append(new Column() { Min = 1, Max = 1, Width = 43, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 2, Max = 2, Width = 14, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 3, Max = 3, Width = 14, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 4, Max = 4, Width = 14, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 5, Max = 5, Width = 14, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 6, Max = 6, Width = 14, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 7, Max = 7, Width = 30, CustomWidth = true });
            }
            // Only insert the columns if you have to create a new columns element
            if (needToInsertColumns)
                sheetPart.Worksheet.InsertAt(lstColumns, 0);
        }

        public static void ChangetheColumnWidthOfWorkOrderTable(Columns cols, WorksheetPart sheetPart, string OrderType)
        {
            Columns lstColumns = cols;
            Boolean needToInsertColumns = false;
            if (lstColumns == null)
            {
                lstColumns = new Columns();
                needToInsertColumns = true;
            }

            if (OrderType == "TakeDown ORDER : INSTRUCTIONS & CONFIRMATION OF RECEIPT OF MATERIALS")
            {
                lstColumns.Append(new Column() { Min = 1, Max = 1, Width = 44, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 2, Max = 2, Width = 14, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 3, Max = 3, Width = 17, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 4, Max = 4, Width = 42, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 5, Max = 5, Width = 36, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 6, Max = 6, Width = 16, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 7, Max = 7, Width = 16, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 8, Max = 8, Width = 30, CustomWidth = true });
            }

            if (OrderType == "POSTING ORDER : INSTRUCTIONS & CONFIRMATION OF RECEIPT OF MATERIALS")
            {
                lstColumns.Append(new Column() { Min = 1, Max = 1, Width = 36, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 2, Max = 2, Width = 14, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 3, Max = 3, Width = 19, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 4, Max = 4, Width = 24, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 5, Max = 5, Width = 37, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 6, Max = 6, Width = 16, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 7, Max = 7, Width = 16, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 8, Max = 8, Width = 30, CustomWidth = true });
            }

            // Only insert the columns if we had to create a new columns element
            if (needToInsertColumns)
                sheetPart.Worksheet.InsertAt(lstColumns, 0);
        }
        private static HexBinaryValue TranslateForeground(System.Drawing.Color fillColor)
        {
            return new HexBinaryValue()
            {
                Value = System.Drawing.ColorTranslator.ToHtml(System.Drawing.Color.FromArgb(fillColor.A,fillColor.R,fillColor.G,fillColor.B)).Replace("#", "")
            };
        }
        public static void CreateExcelWithDynamicHeader(string fileName, List<string> headerFields, string sheetName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = sheetName };

                sheets.Append(sheet);

                workbookPart.Workbook.Save();

                double width = 6;
                // Constructing header
                Row headerRow = new Row();
                for (int i = 0; i < headerFields.Count; i++)
                {
                    if (width < headerFields[i].Length)
                    {
                        width = headerFields[i].Length;
                    }

                    Columns columns = new Columns();
                    columns.Append(new Column() { Min = 1, Max = (UInt32)headerFields.Count, Width = width, CustomWidth = true });

                    worksheetPart.Worksheet.Append(columns);
                    Cell cell = new Cell();
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(headerFields[i]);

                    headerRow.AppendChild(cell);
                }

                SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                sheetData.AppendChild(headerRow);

                workbookPart.Workbook.Save();
            }
        }
    }
}
