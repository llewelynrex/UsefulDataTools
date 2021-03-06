using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading;
using NetOffice.ExcelApi.Enums;
using Excel = NetOffice.ExcelApi;

namespace UsefulDataTools
{
   public static class ExcelOutputExtensions
   {
      /// <summary>
      ///    Evaluates, expands and exports an <see cref="IEnumerable{T}" /> into Excel and opens Excel to display the exported
      ///    data. If <see cref="T" /> is a complex type, all properties and fields are exported into columns. If
      ///    <see cref="T" /> is a simple type, the data is exported into a single column.
      /// </summary>
      /// <typeparam name="T">A generic type which is contained within the IEnumerable.</typeparam>
      /// <param name="input">An <see cref="IEnumerable{T}" /> which will be evaluated, expanded and exported into Excel.</param>
      /// <returns>Returns the <see cref="IEnumerable{T}" /> for further processing if required.</returns>
      public static IEnumerable<T> ToExcel<T>(this IEnumerable<T> input)
      {
         ExcelOutputItem.CreateInstance(input).ToExcel();
         return input;
      }

      /// <summary>
      ///    Evaluates, expands and exports an <see cref="IEnumerable" /> of type <see cref="type" /> into Excel. If
      ///    <see cref="type" /> is a complex type, all properties and fields are exported into columns. If <see cref="type" />
      ///    is a simple type, the data is exported into a single column.
      /// </summary>
      /// <param name="input">An <see cref="IEnumerable" /> which will be evaluated, expanded and exported into Excel.</param>
      /// <param name="type">The type of the objects contained within the IEnumerable in <see cref="input" />.</param>
      /// <returns>Returns the <see cref="IEnumerable{T}" /> for further processing if required.</returns>
      public static IEnumerable ToExcel(IEnumerable input, Type type)
      {
         ExcelOutputItem.CreateInstance(input, type).ToExcel();
         return input;
      }

      /// <summary>
      ///    Evaluates, expands and exports an <see cref="IEnumerable{T}" /> into Excel and opens Excel to display the exported
      ///    data. If <see cref="T" /> is a complex type, all properties and fields are exported into columns. If
      ///    <see cref="T" /> is a simple type, the data is exported into a single column.
      /// </summary>
      /// <typeparam name="T">A generic type which is contained within the IEnumerable.</typeparam>
      /// <param name="input">An <see cref="IEnumerable{T}" /> which will be evaluated, expanded and exported into Excel.</param>
      /// <param name="trim">A parameter which determines whether strings should be trimmed before being written to Excel.</param>
      /// <returns>Returns the <see cref="IEnumerable{T}" /> for further processing if required.</returns>
      public static IEnumerable<T> ToExcel<T>(this IEnumerable<T> input, bool trim)
      {
         ExcelOutputItem.CreateInstance(input, trim).ToExcel();
         return input;
      }

      /// <summary>
      ///    Evaluates, expands and exports an <see cref="IEnumerable" /> of type <see cref="type" /> into Excel. If
      ///    <see cref="type" /> is a complex type, all properties and fields are exported into columns. If <see cref="type" />
      ///    is a simple type, the data is exported into a single column.
      /// </summary>
      /// <param name="input">An <see cref="IEnumerable" /> which will be evaluated, expanded and exported into Excel.</param>
      /// <param name="type">The type of the objects contained within the IEnumerable in <see cref="input" />.</param>
      /// <param name="trim">A parameter which determines whether strings should be trimmed before being written to Excel.</param>
      /// <returns>Returns the <see cref="IEnumerable{T}" /> for further processing if required.</returns>
      public static IEnumerable ToExcel(IEnumerable input, Type type, bool trim)
      {
         ExcelOutputItem.CreateInstance(input, type, trim).ToExcel();
         return input;
      }

      /// <summary>
      ///    Evaluates, expands and exports an <see cref="IEnumerable{T}" /> into Excel and opens Excel to display the exported
      ///    data. If <see cref="T" /> is a complex type, all properties and fields are exported into columns. If
      ///    <see cref="T" /> is a simple type, the data is exported into a single column.
      /// </summary>
      /// <typeparam name="T">A generic type which is contained within the IEnumerable.</typeparam>
      /// <param name="input">An <see cref="IEnumerable{T}" /> which will be evaluated, expanded and exported into Excel.</param>
      /// <param name="worksheetName">The name given to the newly generated worksheet.</param>
      /// <returns>Returns the <see cref="IEnumerable{T}" /> for further processing if required.</returns>
      public static IEnumerable<T> ToExcel<T>(this IEnumerable<T> input, string worksheetName)
      {
         ExcelOutputItem.CreateInstance(input, worksheetName).ToExcel();
         return input;
      }

      /// <summary>
      ///    Evaluates, expands and exports an <see cref="IEnumerable" /> of type <see cref="type" /> into Excel. If
      ///    <see cref="type" /> is a complex type, all properties and fields are exported into columns. If <see cref="type" />
      ///    is a simple type, the data is exported into a single column.
      /// </summary>
      /// <param name="input">An <see cref="IEnumerable" /> which will be evaluated, expanded and exported into Excel.</param>
      /// <param name="type">The type of the objects contained within the IEnumerable in <see cref="input" />.</param>
      /// <param name="worksheetName">The name given to the newly generated worksheet.</param>
      /// <returns>Returns the <see cref="IEnumerable{T}" /> for further processing if required.</returns>
      public static IEnumerable ToExcel(this IEnumerable input, Type type, string worksheetName)
      {
         ExcelOutputItem.CreateInstance(input, type, worksheetName).ToExcel();
         return input;
      }

      /// <summary>
      ///    Evaluates, expands and exports an <see cref="IEnumerable{T}" /> into Excel. If <see cref="T" /> is a complex type,
      ///    all properties and fields are exported into columns. If <see cref="T" /> is a simple type, the data is exported into
      ///    a single column.
      /// </summary>
      /// <typeparam name="T">A generic type which is contained within the IEnumerable.</typeparam>
      /// <param name="input">An <see cref="IEnumerable{T}" /> which will be evaluated, expanded and exported into Excel.</param>
      /// <param name="postCreationAction">
      ///    Determines whether the Excel file will be opened with the data, opened and saved or just saved.
      /// </param>
      /// <param name="path">
      ///    If the path parameter determines where the Excel file will be saved to if a save action is selected from the
      ///    <see cref="PostCreationActions" />.
      /// </param>
      /// <returns>Returns the <see cref="IEnumerable{T}" /> for further processing if required.</returns>
      public static IEnumerable<T> ToExcel<T>(this IEnumerable<T> input, PostCreationActions postCreationAction, string path)
      {
         ExcelOutputItem.CreateInstance(input, postCreationAction, path).ToExcel();
         return input;
      }

      /// <summary>
      ///    Evaluates, expands and exports an <see cref="IEnumerable" /> of type <see cref="type" /> into Excel. If
      ///    <see cref="type" /> is a complex type, all properties and fields are exported into columns. If <see cref="type" />
      ///    is a simple type, the data is exported into a single column.
      /// </summary>
      /// <param name="input">An <see cref="IEnumerable" /> which will be evaluated, expanded and exported into Excel.</param>
      /// <param name="type">The type of the objects contained within the IEnumerable in <see cref="input" />.</param>
      /// <param name="postCreationAction">
      ///    Determines whether the Excel file will be opened with the data, opened and saved or just saved.
      /// </param>
      /// <param name="path">
      ///    If the path parameter determines where the Excel file will be saved to if a save action is selected from the
      ///    <see cref="PostCreationActions" />.
      /// </param>
      /// <returns>Returns the <see cref="IEnumerable{T}" /> for further processing if required.</returns>
      public static IEnumerable ToExcel(this IEnumerable input, Type type, PostCreationActions postCreationAction, string path)
      {
         ExcelOutputItem.CreateInstance(input, type, postCreationAction, path).ToExcel();
         return input;
      }

      /// <summary>
      ///    Evaluates, expands and exports an <see cref="IEnumerable{T}" /> into Excel. If <see cref="T" /> is a complex type,
      ///    all properties and fields are exported into columns. If <see cref="T" /> is a simple type, the data is exported into
      ///    a single column.
      /// </summary>
      /// <typeparam name="T">A generic type which is contained within the IEnumerable.</typeparam>
      /// <param name="input">An <see cref="IEnumerable{T}" /> which will be evaluated, expanded and exported into Excel.</param>
      /// <param name="trim">A parameter which determines whether strings should be trimmed before being written to Excel.</param>
      /// <param name="postCreationAction">
      ///    Determines whether the Excel file will be opened with the data, opened and saved or just saved.
      /// </param>
      /// <param name="path">
      ///    If the path parameter determines where the Excel file will be saved to if a save action is selected from the
      ///    <see cref="PostCreationActions" />.
      /// </param>
      /// <param name="worksheetName">The name given to the newly generated worksheet.</param>
      /// <returns>Returns the <see cref="IEnumerable{T}" /> for further processing if required.</returns>
      public static IEnumerable<T> ToExcel<T>(this IEnumerable<T> input, string worksheetName, bool trim, PostCreationActions postCreationAction, string path)
      {
         ExcelOutputItem.CreateInstance(input, worksheetName, trim, postCreationAction, path).ToExcel();
         return input;
      }

      /// <summary>
      ///    Evaluates, expands and exports an <see cref="IEnumerable" /> of type <see cref="type" /> into Excel. If
      ///    <see cref="type" /> is a complex type, all properties and fields are exported into columns. If <see cref="type" />
      ///    is a simple type, the data is exported into a single column.
      /// </summary>
      /// <param name="input">An <see cref="IEnumerable" /> which will be evaluated, expanded and exported into Excel.</param>
      /// <param name="type">The type of the objects contained within the IEnumerable in <see cref="input" />.</param>
      /// <param name="trim">A parameter which determines whether strings should be trimmed before being written to Excel.</param>
      /// <param name="postCreationAction">
      ///    Determines whether the Excel file will be opened with the data, opened and saved or just saved.
      /// </param>
      /// <param name="path">
      ///    If the path parameter determines where the Excel file will be saved to if a save action is selected from the
      ///    <see cref="PostCreationActions" />.
      /// </param>
      /// <param name="worksheetName">The name given to the newly generated worksheet.</param>
      /// <returns>Returns the <see cref="IEnumerable{T}" /> for further processing if required.</returns>
      public static IEnumerable ToExcel(IEnumerable input, Type type, string worksheetName, bool trim, PostCreationActions postCreationAction, string path)
      {
         ExcelOutputItem.CreateInstance(input, type, worksheetName, trim, postCreationAction, path).ToExcel();
         return input;
      }

      /// <summary>
      ///    Evaluates, expands and exports a <see cref="DataTable" /> into Excel.
      /// </summary>
      /// <param name="input">An <see cref="IEnumerable" /> which will be evaluated, expanded and exported into Excel.</param>
      /// <returns>Returns the <see cref="DataTable" /> for further processing if required.</returns>
      public static DataTable ToExcel(this DataTable input)
      {
         ExcelOutputItem.CreateInstance(input).ToExcel();
         return input;
      }

      /// <summary>
      ///    Evaluates, expands and exports a <see cref="DataTable" /> into Excel.
      /// </summary>
      /// <param name="input">An <see cref="IEnumerable" /> which will be evaluated, expanded and exported into Excel.</param>
      /// <param name="worksheetName">The name given to the newly generated worksheet.</param>
      /// <returns>Returns the <see cref="DataTable" /> for further processing if required.</returns>
      public static DataTable ToExcel(this DataTable input, string worksheetName)
      {
         ExcelOutputItem.CreateInstance(input, worksheetName).ToExcel();
         return input;
      }

      /// <summary>
      ///    Evaluates, expands and exports a <see cref="DataTable" /> into Excel.
      /// </summary>
      /// <param name="input">An <see cref="IEnumerable" /> which will be evaluated, expanded and exported into Excel.</param>
      /// <param name="worksheetName">The name given to the newly generated worksheet.</param>
      /// <param name="trim">A parameter which determines whether strings should be trimmed before being written to Excel.</param>
      /// <returns>Returns the <see cref="DataTable" /> for further processing if required.</returns>
      public static DataTable ToExcel(this DataTable input, string worksheetName, bool trim)
      {
         ExcelOutputItem.CreateInstance(input, worksheetName, trim).ToExcel();
         return input;
      }

      /// <summary>
      ///    Evaluates, expands and exports a <see cref="DataTable" /> into Excel.
      /// </summary>
      /// <param name="input">An <see cref="IEnumerable" /> which will be evaluated, expanded and exported into Excel.</param>
      /// <param name="worksheetName">The name given to the newly generated worksheet.</param>
      /// <param name="postCreationAction">
      ///    Determines whether the Excel file will be opened with the data, opened and saved or just saved.
      /// </param>
      /// <param name="path">
      ///    If the path parameter determines where the Excel file will be saved to if a save action is selected from the
      ///    <see cref="PostCreationActions" />.
      /// </param>
      /// <returns>Returns the <see cref="DataTable" /> for further processing if required.</returns>
      public static DataTable ToExcel(this DataTable input, string worksheetName, PostCreationActions postCreationAction, string path)
      {
         ExcelOutputItem.CreateInstance(input, worksheetName, postCreationAction, path).ToExcel();
         return input;
      }

      /// <summary>
      ///    Evaluates, expands and exports a <see cref="DataTable" /> into Excel.
      /// </summary>
      /// <param name="input">An <see cref="IEnumerable" /> which will be evaluated, expanded and exported into Excel.</param>
      /// <param name="trim">A parameter which determines whether strings should be trimmed before being written to Excel.</param>
      /// <returns>Returns the <see cref="DataTable" /> for further processing if required.</returns>
      public static DataTable ToExcel(this DataTable input, bool trim)
      {
         ExcelOutputItem.CreateInstance(input, trim).ToExcel();
         return input;
      }

      /// <summary>
      ///    Evaluates, expands and exports a <see cref="DataTable" /> into Excel.
      /// </summary>
      /// <param name="input">An <see cref="IEnumerable" /> which will be evaluated, expanded and exported into Excel.</param>
      /// <param name="postCreationAction">
      ///    Determines whether the Excel file will be opened with the data, opened and saved or just saved.
      /// </param>
      /// <param name="path">
      ///    If the path parameter determines where the Excel file will be saved to if a save action is selected from the
      ///    <see cref="PostCreationActions" />.
      /// </param>
      /// <returns>Returns the <see cref="DataTable" /> for further processing if required.</returns>
      public static DataTable ToExcel(this DataTable input, PostCreationActions postCreationAction, string path)
      {
         ExcelOutputItem.CreateInstance(input, postCreationAction, path).ToExcel();
         return input;
      }

      /// <summary>
      ///    Evaluates, expands and exports a <see cref="DataTable" /> into Excel.
      /// </summary>
      /// <param name="input">An <see cref="IEnumerable" /> which will be evaluated, expanded and exported into Excel.</param>
      /// <param name="trim">A parameter which determines whether strings should be trimmed before being written to Excel.</param>
      /// <param name="postCreationAction">
      ///    Determines whether the Excel file will be opened with the data, opened and saved or just saved.
      /// </param>
      /// <param name="path">
      ///    If the path parameter determines where the Excel file will be saved to if a save action is selected from the
      ///    <see cref="PostCreationActions" />.
      /// </param>
      /// <returns>Returns the <see cref="DataTable" /> for further processing if required.</returns>
      public static DataTable ToExcel(this DataTable input, bool trim, PostCreationActions postCreationAction, string path)
      {
         ExcelOutputItem.CreateInstance(input, trim, postCreationAction, path);
         return input;
      }

      /// <summary>
      ///    Evaluates, expands and exports a <see cref="DataTable" /> into Excel.
      /// </summary>
      /// <param name="input">An <see cref="IEnumerable" /> which will be evaluated, expanded and exported into Excel.</param>
      /// <param name="worksheetName">The name given to the newly generated worksheet.</param>
      /// <param name="trim">A parameter which determines whether strings should be trimmed before being written to Excel.</param>
      /// <param name="postCreationAction">
      ///    Determines whether the Excel file will be opened with the data, opened and saved or just saved.
      /// </param>
      /// <param name="path">
      ///    If the path parameter determines where the Excel file will be saved to if a save action is selected from the
      ///    <see cref="PostCreationActions" />.
      /// </param>
      /// <returns>Returns the <see cref="DataTable" /> for further processing if required.</returns>
      public static DataTable ToExcel(this DataTable input, string worksheetName, bool trim, PostCreationActions postCreationAction, string path)
      {
         ExcelOutputItem.CreateInstance(input, worksheetName, trim, postCreationAction, path).ToExcel();
         return input;
      }

      /// <summary>
      ///    Evaluates, expands and exports a <see cref="DataSet" /> into Excel.
      /// </summary>
      /// <param name="input">An <see cref="DataSet" /> which will be evaluated, expanded and exported into Excel.</param>
      /// <returns>Returns the <see cref="DataSet" /> for further processing if required.</returns>
      public static DataSet ToExcel(this DataSet input)
      {
         return input.ToExcel(true, PostCreationActions.Open, string.Empty);
      }

      /// <summary>
      ///    Evaluates, expands and exports a <see cref="DataSet" /> into Excel.
      /// </summary>
      /// <param name="input">An <see cref="DataSet" /> which will be evaluated, expanded and exported into Excel.</param>
      /// <param name="trim">A parameter which determines whether strings should be trimmed before being written to Excel.</param>
      /// <returns>Returns the <see cref="DataSet" /> for further processing if required.</returns>
      public static DataSet ToExcel(this DataSet input, bool trim)
      {
         return input.ToExcel(trim, PostCreationActions.Open, string.Empty);
      }

      /// <summary>
      ///    Evaluates, expands and exports a <see cref="DataSet" /> into Excel.
      /// </summary>
      /// <param name="input">An <see cref="DataSet" /> which will be evaluated, expanded and exported into Excel.</param>
      /// <param name="postCreationAction">
      ///    Determines whether the Excel file will be opened with the data, opened and saved or just saved.
      /// </param>
      /// <param name="path">
      ///    If the path parameter determines where the Excel file will be saved to if a save action is selected from the
      ///    <see cref="PostCreationActions" />.
      /// </param>
      /// <returns>Returns the <see cref="DataSet" /> for further processing if required.</returns>
      public static DataSet ToExcel(this DataSet input, PostCreationActions postCreationAction, string path)
      {
         return input.ToExcel(true, PostCreationActions.Open, path);
      }

      /// <summary>
      ///    Evaluates, expands and exports a <see cref="DataSet" /> into Excel.
      /// </summary>
      /// <param name="input">An <see cref="DataSet" /> which will be evaluated, expanded and exported into Excel.</param>
      /// <param name="trim">A parameter which determines whether strings should be trimmed before being written to Excel.</param>
      /// <param name="postCreationAction">
      ///    Determines whether the Excel file will be opened with the data, opened and saved or just saved.
      /// </param>
      /// <param name="path">
      ///    If the path parameter determines where the Excel file will be saved to if a save action is selected from the
      ///    <see cref="PostCreationActions" />.
      /// </param>
      /// <returns>Returns the <see cref="DataSet" /> for further processing if required.</returns>
      public static DataSet ToExcel(this DataSet input, bool trim, PostCreationActions postCreationAction, string path)
      {
         var excelOutputCollection = new ExcelOutputCollection();
         var unnamedWorksheetCounter = 1;
         foreach (DataTable dataTable in input.Tables)
            if (string.IsNullOrWhiteSpace(dataTable.TableName))
            {
               var excelOutputItem = ExcelOutputItem.CreateInstance(dataTable, string.Concat(ExcelOutputConfiguration.DefaultWorksheetName, unnamedWorksheetCounter), trim, postCreationAction, path);
               excelOutputCollection.Add(excelOutputItem);
               unnamedWorksheetCounter++;
            }
            else
            {
               var excelOutputItem = ExcelOutputItem.CreateInstance(dataTable, trim, postCreationAction, path);
               excelOutputCollection.Add(excelOutputItem);
            }

         excelOutputCollection.ToExcel();
         return input;
      }

      /// <summary>
      ///    Exports an <see cref="ExcelOutputItem" />  into Excel.
      /// </summary>
      /// <param name="excelOutputItem">
      ///    An object containing all the necessary parameters to create an Excel export into a single worksheet.
      /// </param>
      public static void ToExcel(this ExcelOutputItem excelOutputItem)
      {
         var excelOutputCollection = new ExcelOutputCollection(excelOutputItem.PostCreationActions, excelOutputItem.Path) {excelOutputItem};
         excelOutputCollection.ToExcel();
      }

      /// <summary>
      ///    Exports an <see cref="ExcelOutputCollection" /> into Excel. A new sheet is created for each
      ///    <see cref="ExcelOutputItem" /> within the <see cref="ExcelOutputCollection" />.
      /// </summary>
      /// <param name="excelOutputCollection">
      ///    An object containing all the necessary parameters to create an Excel export into a multiple worksheets.
      /// </param>
      public static void ToExcel(this ExcelOutputCollection excelOutputCollection)
      {
         //HACK: Workaround for Excel on machines which are set up in the English language, but not an English region.
         var enusCultureInfo = CultureInfo.GetCultureInfo("en-US");
         Thread.CurrentThread.CurrentCulture = enusCultureInfo;

         AbortIfPathIsEmptyAndSaveIsRequired(excelOutputCollection.Path, excelOutputCollection.PostCreationAction);

         using (var app = new Excel.Application())
         {
            var sheetsInNewWorkbook = app.SheetsInNewWorkbook;
            app.SheetsInNewWorkbook = excelOutputCollection.Count;

            try
            {
               CreateWorkbook(app, excelOutputCollection);

               ExecutePostCreationActions(app, excelOutputCollection.PostCreationAction, excelOutputCollection.Path);
            }
            catch (Exception)
            {
               if (app.Workbooks.Any())
                  foreach (var workbook in app.Workbooks.Where(x => !x.IsDisposed))
                  {
                     workbook.Close(false, Missing.Value, Missing.Value);
                     workbook.Dispose();
                  }

               if (app.IsDisposed)
                  throw;
               app.SheetsInNewWorkbook = sheetsInNewWorkbook;
               app.Quit();
               app.Dispose();
               throw;
            }
            finally
            {
               Thread.CurrentThread.CurrentCulture = ExcelOutputConfiguration.DefaultCulture;
            }
         }
      }

      private static void ExecutePostCreationActions(Excel.Application app, PostCreationActions postCreationAction, string path)
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

      private static void CreateWorkbook(Excel.Application app, ExcelOutputCollection excelOutputCollection)
      {
         using (var wb = app.Workbooks.Add())
         {
            var count = 1;
            foreach (var item in excelOutputCollection)
            {
               if (item.Enumerable != null)
                  CreateWorksheetFromEnumerableItem(ref count, wb, item);
               if (item.DataTable != null)
                  CreateWorksheetFromDataTableItem(ref count, wb, item);
            }
         }
      }

      private static void CreateWorksheetFromEnumerableItem(ref int count, Excel.Workbook wb, ExcelOutputItem item)
      {
         using (var ws = (Excel.Worksheet) wb.Worksheets[count])
         {
            var enumerator = item.Enumerable.GetEnumerator();
            var arrayList = new ArrayList();
            while (enumerator.MoveNext())
               arrayList.Add(enumerator.Current);

            var inputArray = arrayList.ToArray();
            var rowCount = inputArray.Length;

            ws.Name = item.WorksheetName;

            if (item.Type.IsSimpleType())
               SimpleTypeProcessing(item.Type, inputArray, rowCount, ws);
            else
               ComplexTypeProcessing(item.Trim, item.Type, inputArray, rowCount, ws);

            ws.Columns.AutoFit();
         }

         count++;
      }

      private static void CreateWorksheetFromDataTableItem(ref int count, Excel.Workbook wb, ExcelOutputItem item)
      {
         using (var ws = (Excel.Worksheet) wb.Worksheets[count])
         {
            var enumerator = item.DataTable.Rows.GetEnumerator();
            var arrayList = new ArrayList();
            while (enumerator.MoveNext())
               arrayList.Add(enumerator.Current);

            var inputArray = arrayList.ToArray();
            var rowCount = inputArray.Length;

            ws.Name = item.WorksheetName;

            DataTableProcessing(item.Trim, item.WorksheetName, item.DataTable.Columns, inputArray, rowCount, ws);

            ws.Columns.AutoFit();
         }

         count++;
      }

      private static void ComplexTypeProcessing(bool stringsTrimmed, Type type, object[] inputArray, int rowCount, Excel.Worksheet ws)
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

         var headerList = new List<string>(simplePropertyCount + simpleFieldCount);
         headerList.AddRange(propertyHeaderArray);
         headerList.AddRange(fieldHeaderArray);

         var headerArray = headerList.ToArray();
         headerRange.Value2 = headerArray;

         SetPropertyColumnDataTypes(ws, simpleProperties);

         SetFieldColumnDataTypes(ws, simpleProperties.Length, simpleFields);

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
         fullRange.Worksheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, fullRange, Type.Missing, XlYesNoGuess.xlYes, Type.Missing).Name = $"Table_{type.Name}";
      }

      private static void DataTableProcessing(bool stringsTrimmed, string name, DataColumnCollection columns, object[] inputArray, int rowCount, Excel.Worksheet ws)
      {
         var columnsCount = columns.Count;

         var headerStartCell = ws.Cells[1, 1];
         var headerEndCell = ws.Cells[1, columnsCount];
         var headerRange = ws.Range(headerStartCell, headerEndCell);

         var headerList = new List<string>(columnsCount);
         foreach (DataColumn column in columns)
            headerList.Add(column.ColumnName);

         var headerArray = headerList.ToArray();
         headerRange.Value2 = headerArray;

         SetDataTableColumnDataTypes(ws, columns);

         var dataStartCell = ws.Cells[2, 1];
         var dataEndCell = ws.Cells[rowCount + 1, columnsCount];
         var dataRange = ws.Range(dataStartCell, dataEndCell);

         var dataArray = new object[rowCount, columnsCount];
         for (var row = 0; row < rowCount; row++) SetDataTableDataPerColumn(stringsTrimmed, inputArray, row, columnsCount, dataArray);

         dataRange.Value2 = dataArray;

         var fullRange = ws.Range(headerStartCell, dataEndCell);
         fullRange.Worksheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, fullRange, Type.Missing, XlYesNoGuess.xlYes, Type.Missing).Name = $"Table_{name}";
      }

      private static void SetPropertyColumnDataTypes(Excel.Worksheet ws, PropertyInfo[] simpleProperties)
      {
         for (var column = 0; column < simpleProperties.Length; column++)
         {
            var rangeColumn = GetExcelColumnName(column + 1);
            var range = ws.Range($"{rangeColumn}:{rangeColumn}");
            var columnDataTypeAttribute = simpleProperties[column].GetCustomAttribute<ColumnNumberFormatAttribute>();
            if (columnDataTypeAttribute == null)
               SetRangeNumberFormatBasedOnDataType(simpleProperties[column].PropertyType, range);
            else
               SetRangeNumberFormatBasedOnAttribute(range, columnDataTypeAttribute.NumberFormat);
         }
      }

      private static void SetFieldColumnDataTypes(Excel.Worksheet ws, int simplePropertiesCount, FieldInfo[] simpleFields)
      {
         for (var column = simplePropertiesCount; column < simplePropertiesCount + simpleFields.Length; column++)
         {
            var rangeColumn = GetExcelColumnName(column + 1);
            var range = ws.Range($"{rangeColumn}:{rangeColumn}");
            var columnDataTypeAttribute = simpleFields[column - simplePropertiesCount].GetCustomAttribute<ColumnNumberFormatAttribute>();
            if (columnDataTypeAttribute == null)
               SetRangeNumberFormatBasedOnDataType(simpleFields[column - simplePropertiesCount].FieldType, range);
            else
               SetRangeNumberFormatBasedOnAttribute(range, columnDataTypeAttribute.NumberFormat);
         }
      }

      private static void SetDataTableColumnDataTypes(Excel.Worksheet ws, DataColumnCollection columns)
      {
         for (var column = 0; column < columns.Count; column++)
         {
            var rangeColumn = GetExcelColumnName(column + 1);
            var range = ws.Range($"{rangeColumn}:{rangeColumn}");
            SetRangeNumberFormatBasedOnDataType(columns[column].DataType, range);
         }
      }

      private static void SetRangeNumberFormatBasedOnDataType(Type type, Excel.Range range)
      {
         if (type == typeof(string) || type == typeof(char) || Nullable.GetUnderlyingType(type) == typeof(char))
            range.NumberFormat = "@";
         else if (type == typeof(DateTime) || Nullable.GetUnderlyingType(type) == typeof(DateTime))
            range.NumberFormat = ExcelOutputConfiguration.DefaultDateTimeFormat;
      }

      private static void SetRangeNumberFormatBasedOnAttribute(Excel.Range range, string numberFormat)
      {
         range.NumberFormat = numberFormat;
      }

      private static void SetPropertyDataPerColumn(bool stringsTrimmed, object[] inputArray, PropertyInfo[] simpleProperties, int simplePropertyCount, object[,] dataArray, int row)
      {
         for (var column = 0; column < simplePropertyCount; column++) inputArray[row].SetDataPerColumn(stringsTrimmed, column, row, dataArray, simpleProperties[column].GetValue);
      }

      private static void SetFieldDataPerColumn(bool stringsTrimmed, object[] inputArray, FieldInfo[] simpleFields, int simplePropertyCount, int simpleFieldCount, object[,] dataArray, int row)
      {
         for (var column = simplePropertyCount; column < simplePropertyCount + simpleFieldCount; column++) inputArray[row].SetDataPerColumn(stringsTrimmed, column, row, dataArray, simpleFields[column - simplePropertyCount].GetValue);
      }

      private static void SetDataTableDataPerColumn(bool stringsTrimmed, object[] inputArray, int row, int columnCount, object[,] dataArray)
      {
         var dataRow = (DataRow) inputArray[row];
         for (var i = 0; i < columnCount; i++)
         {
            var column = i;
            inputArray[row].SetDataPerColumn(stringsTrimmed, i, row, dataArray, o => dataRow.ItemArray[column]);
         }
      }

      private static void SetDataPerColumn(this object item, bool stringsTrimmed, int column, int row, object[,] dataArray, Func<object, object> getValueFunction)
      {
         var value = getValueFunction(item);
         if (value != null && (value is string || value is char) && stringsTrimmed)
            dataArray[row, column] = value.ToTrimmedString();
         else
            dataArray[row, column] = value;
      }

      private static void SimpleTypeProcessing(Type type, object[] inputArray, int rowCount, Excel.Worksheet ws)
      {
         var dataArray = new object[rowCount, 1];
         for (var i = 0; i < inputArray.Length; i++)
            dataArray[i, 0] = inputArray[i];

         var dataStartCell = ws.Cells[1, 1];
         var dataEndCell = ws.Cells[rowCount, 1];
         var dataRange = ws.Range(dataStartCell, dataEndCell);

         const string rangeColumn = "A";
         var range = ws.Range($"{rangeColumn}:{rangeColumn}");
         SetRangeNumberFormatBasedOnDataType(type, range);

         dataRange.Value2 = dataArray;

         var fullRange = ws.Range(dataStartCell, dataEndCell);
         fullRange.Worksheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, fullRange, Type.Missing, XlYesNoGuess.xlNo, Type.Missing).Name = $"Table_{type.Name}";
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
      ///    Returns the alphabetical column name when given the number of an excel column.
      /// </summary>
      /// <param name="columnNumber">The number to be converted to text</param>
      /// <returns>
      ///    <see cref="string" />
      /// </returns>
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