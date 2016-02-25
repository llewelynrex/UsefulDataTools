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
            var type = typeof (T);
            var inputArray = input.ToArray();
            var rowCount = inputArray.Length;
            var exceptionThrown = false;
            Exception exception = null;
            var dateFormat = Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern + " " + Thread.CurrentThread.CurrentCulture.DateTimeFormat.LongTimePattern;

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

            using (var app = new Application())
            {
                var sheetsInNewWorkbook = app.SheetsInNewWorkbook;
                app.SheetsInNewWorkbook = 1;

                try
                {
                    using (var wb = app.Workbooks.Add())
                    {
                        try
                        {
                            var ws = (Worksheet)wb.Worksheets.Single();
                            ws.Name = type.Name;

                            if (type.IsSimpleType())
                            {
                                var dataArray = new T[rowCount, 1];
                                for (var i = 0; i < inputArray.Length; i++)
                                    dataArray[i, 0] = inputArray[i];

                                var dataStartCell = ws.Cells[1, 1];
                                var dataEndCell = ws.Cells[rowCount, 1];
                                var dataRange = ws.Range(dataStartCell, dataEndCell);

                                const string rangeColumn = "A";
                                var range = ws.Range($"{rangeColumn}:{rangeColumn}");
                                if (type == typeof (string) || type == typeof (char) || Nullable.GetUnderlyingType(type) == typeof (char))
                                    range.NumberFormat = "@";
                                else if (type == typeof (DateTime) || Nullable.GetUnderlyingType(type) == typeof (DateTime))
                                    range.NumberFormat = dateFormat;

                                dataRange.Value2 = dataArray;

                                var fullRange = ws.Range(dataStartCell, dataEndCell);
                                fullRange.Worksheet.ListObjects.Add(NetOffice.ExcelApi.Enums.XlListObjectSourceType.xlSrcRange, fullRange, Type.Missing, NetOffice.ExcelApi.Enums.XlYesNoGuess.xlNo, Type.Missing).Name = $"Table_{type.Name}";
                            }
                            else
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

                                for (var column = 0; column < simpleProperties.Length; column++)
                                {
                                    var rangeColumn = GetExcelColumnName(column + 1);
                                    var range = ws.Range($"{rangeColumn}:{rangeColumn}");
                                    if (simpleProperties[column].PropertyType == typeof (string) || simpleProperties[column].PropertyType == typeof (char) || Nullable.GetUnderlyingType(simpleProperties[column].PropertyType) == typeof (char))
                                        range.NumberFormat = "@";
                                    else if (simpleProperties[column].PropertyType == typeof (DateTime) || Nullable.GetUnderlyingType(simpleProperties[column].PropertyType) == typeof (DateTime))
                                        range.NumberFormat = dateFormat;
                                }

                                for (var column = simpleProperties.Length; column < simpleProperties.Length + simpleFields.Length; column++)
                                {
                                    var rangeColumn = GetExcelColumnName(column + 1);
                                    var range = ws.Range($"{rangeColumn}:{rangeColumn}");
                                    if (simpleFields[column - simpleProperties.Length].FieldType == typeof (string) || simpleFields[column - simpleProperties.Length].FieldType == typeof (char) || Nullable.GetUnderlyingType(simpleFields[column - simpleProperties.Length].FieldType) == typeof (char))
                                        range.NumberFormat = "@";
                                    else if (simpleFields[column - simpleProperties.Length].FieldType == typeof (DateTime) || Nullable.GetUnderlyingType(simpleFields[column - simpleProperties.Length].FieldType) == typeof (DateTime))
                                        range.NumberFormat = dateFormat;
                                }

                                var dataStartCell = ws.Cells[2, 1];
                                var dataEndCell = ws.Cells[rowCount + 1, simplePropertyCount + simpleFieldCount];
                                var dataRange = ws.Range(dataStartCell, dataEndCell);

                                var dataArray = new object[rowCount, simplePropertyCount + simpleFieldCount];
                                for (var row = 0; row < rowCount; row++)
                                {
                                    for (var column = 0; column < simplePropertyCount; column++)
                                    {
                                        var value = simpleProperties[column].GetValue(inputArray[row]);
                                        if (value != null && (value is string || value is char) && stringsTrimmed)
                                            dataArray[row, column] = value.ToTrimmedString();
                                        else
                                            dataArray[row, column] = value;
                                    }
                                    for (var column = simplePropertyCount; column < simplePropertyCount + simpleFieldCount; column++)
                                    {
                                        var value = simpleFields[column - simplePropertyCount].GetValue(inputArray[row]);
                                        if (value != null && (value is string || value is char) && stringsTrimmed)
                                            dataArray[row, column] = value.ToTrimmedString();
                                        else
                                            dataArray[row, column] = value;
                                    }
                                }
                                dataRange.Value2 = dataArray;

                                var fullRange = ws.Range(headerStartCell, dataEndCell);
                                fullRange.Worksheet.ListObjects.Add(NetOffice.ExcelApi.Enums.XlListObjectSourceType.xlSrcRange, fullRange, Type.Missing, NetOffice.ExcelApi.Enums.XlYesNoGuess.xlYes, Type.Missing).Name = $"Table_{type.Name}";
                            }

                            ws.Columns.AutoFit();
                        }
                        catch (Exception ex)
                        {
                            exceptionThrown = true;
                            exception = ex;
                        }
                        finally
                        {
                            if (exceptionThrown)
                            {
                                wb.Close(false, Missing.Value, Missing.Value);
                                if (!wb.IsDisposed)
                                    wb.Dispose();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    exceptionThrown = true;
                    exception = ex;
                }
                finally
                {
                    app.SheetsInNewWorkbook = sheetsInNewWorkbook;
                    if (exceptionThrown)
                    {
                        Workbook wb = null;
                        if (app.Workbooks.Any())
                            wb = app.Workbooks.Single();

                        if (wb != null)
                            if (!wb.IsDisposed)
                            {
                                wb.Close(false, Missing.Value, Missing.Value);
                                wb.Dispose();
                            }
                        if (!app.IsDisposed)
                        {
                            app.Quit();
                            app.Dispose();
                        }
                        if (exception != null) throw exception;
                    }
                    else
                    {
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
                }
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
