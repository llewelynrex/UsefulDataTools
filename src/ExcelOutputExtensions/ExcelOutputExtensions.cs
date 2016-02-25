using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading;
using NetOffice.ExcelApi;

namespace UsefulDataTools
{
    public static class ExcelOutputExtensions
    {
        public static void ToExcel<T>(this IEnumerable<T> input)
        {
            input.ToExcel(true, null, PostCreationActions.Open);
        }

        public static void ToExcel<T>(this IEnumerable<T> input, bool stringsTrimmed, PostCreationActions postCreationAction)
        {
            input.ToExcel(stringsTrimmed, null, postCreationAction);
        }

        public static void ToExcel<T>(this IEnumerable<T> input, string path, PostCreationActions postCreationAction)
        {
            input.ToExcel(true, path, postCreationAction);
        }

        public static void ToExcel<T>(this IEnumerable<T> input, bool stringsTrimmed, string path, PostCreationActions postCreationAction)
        {
            var type = typeof(T);
            var inputArray = input.ToArray();
            var rowCount = inputArray.Length;
            Exception exception = null;
            var dateFormat = Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern + " " + Thread.CurrentThread.CurrentCulture.DateTimeFormat.LongTimePattern;

            AbortIfPathIsEmptyAndSaveIsRequired(path, postCreationAction);

            using (var app = new Application())
            {
                var sheetsInNewWorkbook = app.SheetsInNewWorkbook;
                app.SheetsInNewWorkbook = 1;

                try
                {
                    using (var wb = app.Workbooks.Add())
                    {
                        var ws = (Worksheet)wb.Worksheets.Single();
                        ws.Name = type.Name;

                        if (type.IsSimpleType())
                        {
                            SimpleTypeProcessing(type, inputArray, rowCount, dateFormat, ws);
                        }
                        else
                        {
                            ComplexTypeProcessing(stringsTrimmed, type, inputArray, rowCount, dateFormat, ws);
                        }

                        ws.Columns.AutoFit();
                    }

                    switch (postCreationAction)
                    {
                        case PostCreationActions.Open:
                            app.Visible = true;
                            break;
                        case PostCreationActions.SaveAndView:
                            app.ActiveWorkbook.SaveAs(path);
                            app.Visible = true;
                            break;
                        case PostCreationActions.SaveAndClose:
                            app.ActiveWorkbook.SaveAs(path);
                            app.Quit();
                            app.Dispose();
                            break;
                        default:
                            throw new ArgumentOutOfRangeException(nameof(postCreationAction), postCreationAction, null);
                    }
                }
                catch (Exception ex)
                {
                    exception = ex;
                }
                finally
                {
                    app.SheetsInNewWorkbook = sheetsInNewWorkbook;
                    if (exception != null)
                    {
                        if (app.Workbooks.Any())
                        {
                            foreach (var workbook in app.Workbooks.Where(x => !x.IsDisposed))
                            {
                                workbook.Close(false, Missing.Value, Missing.Value);
                                workbook.Dispose();
                            }
                        }

                        if (app.IsDisposed) throw exception;
                        app.Quit();
                        app.Dispose();
                        throw exception;
                    }
                }
            }
        }

        private static void ComplexTypeProcessing<T>(bool stringsTrimmed, Type type, T[] inputArray, int rowCount, string dateFormat, Worksheet ws)
        {
            var properties = type.GetProperties();
            var fields = type.GetFields();
            var simpleProperties = properties.Where(p => p.PropertyType.IsSimpleType()).ToArray();
            var simpleFields = fields.Where(f => f.FieldType.IsSimpleType()).ToArray();
            var simplePropertyCount = simpleProperties.Length;
            var simpleFieldCount = simpleFields.Length;

            var headerStartCell = ws.Cells[1, 1];
            var headerEndCell = ws.Cells[1, simplePropertyCount + simpleFieldCount];
            var headerRange = ws.Range(headerStartCell, headerEndCell);

            var propertyHeaderArray = simpleProperties.Select(x => x.Name).ToArray();
            var fieldHeaderArray = simpleFields.Select(x => x.Name).ToArray();

            var headerList = new List<string>();
            headerList.AddRange(propertyHeaderArray);
            headerList.AddRange(fieldHeaderArray);

            var headerArray = headerList.ToArray();
            headerRange.Value2 = headerArray;

            SetPropertyColumnDataTypes(dateFormat, ws, simpleProperties);

            SetFieldColumnDataTypes(dateFormat, ws, simpleProperties.Length, simpleFields);

            var dataStartCell = ws.Cells[2, 1];
            var dataEndCell = ws.Cells[rowCount + 1, simplePropertyCount + simpleFieldCount];
            var dataRange = ws.Range(dataStartCell, dataEndCell);

            var dataArray = new object[rowCount, simplePropertyCount + simpleFieldCount];
            for (var row = 0; row < rowCount; row++)
            {
                SetPropertyDataPerColumn(stringsTrimmed, inputArray, simpleProperties, simplePropertyCount, dataArray, row);
                SetFieldDataPerColumn(stringsTrimmed, inputArray, simpleFields, simplePropertyCount, simpleFieldCount, dataArray, row);
            }
            dataRange.Value2 = dataArray;

            var fullRange = ws.Range(headerStartCell, dataEndCell);
            fullRange.Worksheet.ListObjects.Add(NetOffice.ExcelApi.Enums.XlListObjectSourceType.xlSrcRange, fullRange, Type.Missing, NetOffice.ExcelApi.Enums.XlYesNoGuess.xlYes, Type.Missing).Name = $"Table_{type.Name}";
        }

        private static void SetPropertyColumnDataTypes(string dateFormat, Worksheet ws, PropertyInfo[] simpleProperties)
        {
            for (var column = 0; column < simpleProperties.Length; column++)
            {
                var rangeColumn = GetExcelColumnName(column + 1);
                var range = ws.Range($"{rangeColumn}:{rangeColumn}");
                if (simpleProperties[column].PropertyType == typeof(string) || simpleProperties[column].PropertyType == typeof(char) || Nullable.GetUnderlyingType(simpleProperties[column].PropertyType) == typeof(char))
                    range.NumberFormat = "@";
                else if (simpleProperties[column].PropertyType == typeof(DateTime) || Nullable.GetUnderlyingType(simpleProperties[column].PropertyType) == typeof(DateTime))
                    range.NumberFormat = dateFormat;
            }
        }

        private static void SetFieldColumnDataTypes(string dateFormat, Worksheet ws, int simplePropertiesCount, FieldInfo[] simpleFields)
        {
            for (var column = simplePropertiesCount; column < simplePropertiesCount + simpleFields.Length; column++)
            {
                var rangeColumn = GetExcelColumnName(column + 1);
                var range = ws.Range($"{rangeColumn}:{rangeColumn}");
                if (simpleFields[column - simplePropertiesCount].FieldType == typeof(string) || simpleFields[column - simplePropertiesCount].FieldType == typeof(char) || Nullable.GetUnderlyingType(simpleFields[column - simplePropertiesCount].FieldType) == typeof(char))
                    range.NumberFormat = "@";
                else if (simpleFields[column - simplePropertiesCount].FieldType == typeof(DateTime) || Nullable.GetUnderlyingType(simpleFields[column - simplePropertiesCount].FieldType) == typeof(DateTime))
                    range.NumberFormat = dateFormat;
            }
        }

        private static void SetPropertyDataPerColumn<T>(bool stringsTrimmed, T[] inputArray, PropertyInfo[] simpleProperties, int simplePropertyCount, object[,] dataArray, int row)
        {
            for (var column = 0; column < simplePropertyCount; column++)
            {
                var value = simpleProperties[column].GetValue(inputArray[row]);
                if (value != null && (value is string || value is char) && stringsTrimmed)
                    dataArray[row, column] = value.ToTrimmedString();
                else
                    dataArray[row, column] = value;
            }
        }

        private static void SetFieldDataPerColumn<T>(bool stringsTrimmed, T[] inputArray, FieldInfo[] simpleFields, int simplePropertyCount, int simpleFieldCount, object[,] dataArray, int row)
        {
            for (var column = simplePropertyCount; column < simplePropertyCount + simpleFieldCount; column++)
            {
                var value = simpleFields[column - simplePropertyCount].GetValue(inputArray[row]);
                if (value != null && (value is string || value is char) && stringsTrimmed)
                    dataArray[row, column] = value.ToTrimmedString();
                else
                    dataArray[row, column] = value;
            }
        }

        private static void SimpleTypeProcessing<T>(Type type, T[] inputArray, int rowCount, string dateFormat, Worksheet ws)
        {
            var dataArray = new T[rowCount, 1];
            for (var i = 0; i < inputArray.Length; i++)
                dataArray[i, 0] = inputArray[i];

            var dataStartCell = ws.Cells[1, 1];
            var dataEndCell = ws.Cells[rowCount, 1];
            var dataRange = ws.Range(dataStartCell, dataEndCell);

            const string rangeColumn = "A";
            var range = ws.Range($"{rangeColumn}:{rangeColumn}");
            if (type == typeof(string) || type == typeof(char) || Nullable.GetUnderlyingType(type) == typeof(char))
                range.NumberFormat = "@";
            else if (type == typeof(DateTime) || Nullable.GetUnderlyingType(type) == typeof(DateTime))
                range.NumberFormat = dateFormat;

            dataRange.Value2 = dataArray;

            var fullRange = ws.Range(dataStartCell, dataEndCell);
            fullRange.Worksheet.ListObjects.Add(NetOffice.ExcelApi.Enums.XlListObjectSourceType.xlSrcRange, fullRange, Type.Missing, NetOffice.ExcelApi.Enums.XlYesNoGuess.xlNo, Type.Missing).Name = $"Table_{type.Name}";
        }

        private static void AbortIfPathIsEmptyAndSaveIsRequired(string path, PostCreationActions postCreationAction)
        {
            switch (postCreationAction)
            {
                case PostCreationActions.SaveAndView:
                case PostCreationActions.SaveAndClose:
                    if (string.IsNullOrEmpty(path))
                        throw new ArgumentException("The path cannot be null or empty.", nameof(path));
                    break;
                case PostCreationActions.Open:
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(postCreationAction), postCreationAction, null);
            }
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            var dividend = columnNumber;
            var columnName = string.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
    }
}
