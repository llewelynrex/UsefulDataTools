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
        /// <summary>
        /// Evaluates, expands and exports an <see cref="IEnumerable{T}"/> into Excel and opens Excel to display the exported data. 
        /// If <see cref="T"/> is a complex type, all properties and fields are exported into columns. If <see cref="T"/> is a simple type, 
        /// the data is exported into a single column.
        /// </summary>
        /// <typeparam name="T">A generic type which is contained within the IEnumerable.</typeparam>
        /// <param name="input">An <see cref="IEnumerable{T}"/> which will be evaluated, expanded and exported into Excel.</param>
        public static void ToExcel<T>(this IEnumerable<T> input)
        {
            input.ToExcel(true, PostCreationActions.Open, null, "Data");
        }

        /// <summary>
        /// Evaluates, expands and exports an <see cref="IEnumerable{T}"/> into Excel and opens Excel to display the exported data. 
        /// If <see cref="T"/> is a complex type, all properties and fields are exported into columns. If <see cref="T"/> is a simple type, 
        /// the data is exported into a single column.
        /// </summary>
        /// <typeparam name="T">A generic type which is contained within the IEnumerable.</typeparam>
        /// <param name="input">An <see cref="IEnumerable{T}"/> which will be evaluated, expanded and exported into Excel.</param>
        /// <param name="trim">A parameter which determines whether strings should be trimmed before being written to Excel.</param>
        public static void ToExcel<T>(this IEnumerable<T> input, bool trim)
        {
            input.ToExcel(trim, PostCreationActions.Open, null, "Data");
        }

        /// <summary>
        /// Evaluates, expands and exports an <see cref="IEnumerable{T}"/> into Excel. If <see cref="T"/> is a complex type, 
        /// all properties and fields are exported into columns. If <see cref="T"/> is a simple type, the data is exported into a 
        /// single column.
        /// </summary>
        /// <typeparam name="T">A generic type which is contained within the IEnumerable.</typeparam>
        /// <param name="input">An <see cref="IEnumerable{T}"/> which will be evaluated, expanded and exported into Excel.</param>
        /// <param name="postCreationAction">Determines whether the Excel file will be opened with the data, opened and saved or just saved.</param>
        /// <param name="path">If the path parameter determines where the Excel file will be saved to if a save action is selected from the <see cref="PostCreationActions"/>.</param>
        public static void ToExcel<T>(this IEnumerable<T> input, PostCreationActions postCreationAction, string path)
        {
            input.ToExcel(true, postCreationAction, path, "Data");
        }

        /// <summary>
        /// Evaluates, expands and exports an <see cref="IEnumerable{T}"/> into Excel. If <see cref="T"/> is a complex type, 
        /// all properties and fields are exported into columns. If <see cref="T"/> is a simple type, the data is exported into a 
        /// single column.
        /// </summary>
        /// <typeparam name="T">A generic type which is contained within the IEnumerable.</typeparam>
        /// <param name="input">An <see cref="IEnumerable{T}"/> which will be evaluated, expanded and exported into Excel.</param>
        /// <param name="trim">A parameter which determines whether strings should be trimmed before being written to Excel.</param>
        /// <param name="postCreationAction">Determines whether the Excel file will be opened with the data, opened and saved or just saved.</param>
        /// <param name="path">If the path parameter determines where the Excel file will be saved to if a save action is selected from the <see cref="PostCreationActions"/>.</param>
        /// <param name="worksheetName">The name given to the newly generated worksheet.</param>
        public static void ToExcel<T>(this IEnumerable<T> input, bool trim, PostCreationActions postCreationAction, string path, string worksheetName)
        {
            var type = typeof(T);
            var inputArray = input.ToArray();
            var rowCount = inputArray.Length;
            Exception exception = null;
            var dateFormat = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern + " " + System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.LongTimePattern;

            AbortIfPathIsEmptyAndSaveIsRequired(path, postCreationAction);

            using (var app = new Application())
            {
                var sheetsInNewWorkbook = app.SheetsInNewWorkbook;
                app.SheetsInNewWorkbook = 1;

                try
                {
                    using (var wb = app.Workbooks.Add())
                    {
                        app.SheetsInNewWorkbook = sheetsInNewWorkbook;
                        using (var ws = (Worksheet)wb.Worksheets.Single())
                        {
                            ws.Name = worksheetName;

                            if (type.IsSimpleType())
                            {
                                SimpleTypeProcessing(type, inputArray, rowCount, dateFormat, ws);
                            }
                            else
                            {
                                ComplexTypeProcessing(trim, type, inputArray, rowCount, dateFormat, ws);
                            }

                            ws.Columns.AutoFit();
                        }
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

                        if (app.IsDisposed)
                            throw exception;
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

            var propertyHeaderArray = simpleProperties.Select(x => x.GetCustomAttribute<ColumnHeaderAttribute>() != null ? x.GetCustomAttribute<ColumnHeaderAttribute>().Header : x.Name).ToArray();
            var fieldHeaderArray = simpleFields.Select(x => x.GetCustomAttribute<ColumnHeaderAttribute>() != null ? x.GetCustomAttribute<ColumnHeaderAttribute>().Header : x.Name).ToArray();

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
                SetRangeNumberFormatBasedOnDataType(simpleProperties[column].PropertyType, range, dateFormat);
            }
        }

        private static void SetFieldColumnDataTypes(string dateFormat, Worksheet ws, int simplePropertiesCount, FieldInfo[] simpleFields)
        {
            for (var column = simplePropertiesCount; column < simplePropertiesCount + simpleFields.Length; column++)
            {
                var rangeColumn = GetExcelColumnName(column + 1);
                var range = ws.Range($"{rangeColumn}:{rangeColumn}");
                SetRangeNumberFormatBasedOnDataType(simpleFields[column - simplePropertiesCount].FieldType, range, dateFormat);
            }
        }

        private static void SetRangeNumberFormatBasedOnDataType(Type type, Range range, string dateFormat)
        {
            if (type == typeof(string) || type == typeof(char) || Nullable.GetUnderlyingType(type) == typeof(char))
                range.NumberFormat = "@";
            else if (type == typeof(DateTime) || Nullable.GetUnderlyingType(type) == typeof(DateTime))
                range.NumberFormat = dateFormat;
        }

        private static void SetPropertyDataPerColumn<T>(bool stringsTrimmed, T[] inputArray, PropertyInfo[] simpleProperties, int simplePropertyCount, object[,] dataArray, int row)
        {
            for (var column = 0; column < simplePropertyCount; column++)
            {
                inputArray[row].SetDataPerColumn(stringsTrimmed, column, row, dataArray, simpleProperties[column].GetValue);
            }
        }

        private static void SetFieldDataPerColumn<T>(bool stringsTrimmed, T[] inputArray, FieldInfo[] simpleFields, int simplePropertyCount, int simpleFieldCount, object[,] dataArray, int row)
        {
            for (var column = simplePropertyCount; column < simplePropertyCount + simpleFieldCount; column++)
            {
                inputArray[row].SetDataPerColumn(stringsTrimmed, column, row, dataArray, simpleFields[column - simplePropertyCount].GetValue);
            }
        }

        private static void SetDataPerColumn<T>(this T item, bool stringsTrimmed, int column, int row, object[,] dataArray, Func<object, object> getValueFunction)
        {
            var value = getValueFunction.Invoke(item);
            if (value != null && (value is string || value is char) && stringsTrimmed)
                dataArray[row, column] = value.ToTrimmedString();
            else
                dataArray[row, column] = value;
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
            SetRangeNumberFormatBasedOnDataType(type, range, dateFormat);

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
                        throw new ArgumentException($"The {nameof(path)} cannot be null or empty if the file is to be saved.", nameof(path));
                    break;
                case PostCreationActions.Open:
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(postCreationAction), postCreationAction, null);
            }
        }

        /// <summary>
        /// Returns the alphabetical column name when given the number of an excel column.
        /// </summary>
        /// <param name="columnNumber">The number to be converted to text</param>
        /// <returns><see cref="string"/></returns>
        public static string GetExcelColumnName(int columnNumber)
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
