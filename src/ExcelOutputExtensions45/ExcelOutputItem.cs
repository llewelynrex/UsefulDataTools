using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace UsefulDataTools
{
    public class ExcelOutputItem
    {
        private ExcelOutputItem() {}

        /// <summary>
        /// Instantiate an instance of the ExcelOutputItem with a particular <see cref="IEnumerable{T}"/> collection.
        /// This object allows for the creation of a single sheet, or when used in conjunction with <see cref="ExcelOutputCollection"/>,
        /// the creation of multiple sheets within a single Workbook.
        /// <typeparam name="T">A generic type which is contained within the IEnumerable.</typeparam>
        /// <param name="enumerable">An enumerable of type <typeparam name="T"/> which contains data to be exported to Excel.</param>
        /// </summary>
        public static ExcelOutputItem CreateInstance<T>(IEnumerable<T> enumerable)
        {
            return CreateInstance(enumerable, "Data", true, PostCreationActions.Open, null);
        }

        /// <summary>
        /// Instantiate an instance of the ExcelOutputItem with a particular <see cref="IEnumerable{T}"/> collection.
        /// This object allows for the creation of a single sheet, or when used in conjunction with <see cref="ExcelOutputCollection"/>,
        /// the creation of multiple sheets within a single Workbook.
        /// <typeparam name="T">A generic type which is contained within the IEnumerable.</typeparam>
        /// <param name="enumerable">An enumerable of type <typeparam name="T"/> which contains data to be exported to Excel.</param>
        /// <param name="postCreationAction">Determines whether the Excel file will be opened with the data, opened and saved or just saved.</param>
        /// <param name="path">If the path parameter determines where the Excel file will be saved to if a save action is selected from the <see cref="PostCreationActions"/>.</param>
        /// </summary>
        public static ExcelOutputItem CreateInstance<T>(IEnumerable<T> enumerable, PostCreationActions postCreationAction, string path)
        {
            return CreateInstance(enumerable, "Data", true, postCreationAction, path);
        }

        /// <summary>
        /// Instantiate an instance of the ExcelOutputItem with a particular <see cref="IEnumerable{T}"/> collection.
        /// This object allows for the creation of a single sheet, or when used in conjunction with <see cref="ExcelOutputCollection"/>,
        /// the creation of multiple sheets within a single Workbook.
        /// <typeparam name="T">A generic type which is contained within the IEnumerable.</typeparam>
        /// <param name="enumerable">An enumerable of type <typeparam name="T"/> which contains data to be exported to Excel.</param>
        /// <param name="worksheetName">The name given to the newly generated worksheet.</param>
        /// </summary>
        public static ExcelOutputItem CreateInstance<T>(IEnumerable<T> enumerable, string worksheetName)
        {
            return CreateInstance(enumerable, worksheetName, true, PostCreationActions.Open, null);
        }

        /// <summary>
        /// Instantiate an instance of the ExcelOutputItem with a particular <see cref="IEnumerable{T}"/> collection.
        /// This object allows for the creation of a single sheet, or when used in conjunction with <see cref="ExcelOutputCollection"/>,
        /// the creation of multiple sheets within a single Workbook.
        /// <typeparam name="T">A generic type which is contained within the IEnumerable.</typeparam>
        /// <param name="enumerable">An enumerable of type <typeparam name="T"/> which contains data to be exported to Excel.</param>
        /// <param name="worksheetName">The name given to the newly generated worksheet.</param>
        /// <param name="postCreationAction">Determines whether the Excel file will be opened with the data, opened and saved or just saved.</param>
        /// <param name="path">If the path parameter determines where the Excel file will be saved to if a save action is selected from the <see cref="PostCreationActions"/>.</param>
        /// </summary>
        public static ExcelOutputItem CreateInstance<T>(IEnumerable<T> enumerable, string worksheetName, PostCreationActions postCreationAction, string path)
        {
            return CreateInstance(enumerable, worksheetName, true, postCreationAction, path);
        }

        /// <summary>
        /// Instantiate an instance of the ExcelOutputItem with a particular <see cref="IEnumerable{T}"/> collection.
        /// This object allows for the creation of a single sheet, or when used in conjunction with <see cref="ExcelOutputCollection"/>,
        /// the creation of multiple sheets within a single Workbook.
        /// <typeparam name="T">A generic type which is contained within the IEnumerable.</typeparam>
        /// <param name="enumerable">An enumerable of type <typeparam name="T"/> which contains data to be exported to Excel.</param>
        /// <param name="trim">A parameter which determines whether strings should be trimmed before being written to Excel.</param>
        /// </summary>
        public static ExcelOutputItem CreateInstance<T>(IEnumerable<T> enumerable, bool trim)
        {
            return CreateInstance(enumerable, "Data", trim, PostCreationActions.Open, null);
        }

        /// <summary>
        /// Instantiate an instance of the ExcelOutputItem with a particular <see cref="IEnumerable{T}"/> collection.
        /// This object allows for the creation of a single sheet, or when used in conjunction with <see cref="ExcelOutputCollection"/>,
        /// the creation of multiple sheets within a single Workbook.
        /// <typeparam name="T">A generic type which is contained within the IEnumerable.</typeparam>
        /// <param name="enumerable">An enumerable of type <typeparam name="T"/> which contains data to be exported to Excel.</param>
        /// <param name="trim">A parameter which determines whether strings should be trimmed before being written to Excel.</param>
        /// <param name="postCreationAction">Determines whether the Excel file will be opened with the data, opened and saved or just saved.</param>
        /// <param name="path">If the path parameter determines where the Excel file will be saved to if a save action is selected from the <see cref="PostCreationActions"/>.</param>
        /// </summary>
        public static ExcelOutputItem CreateInstance<T>(IEnumerable<T> enumerable, bool trim, PostCreationActions postCreationAction, string path)
        {
            return CreateInstance(enumerable, "Data", trim, postCreationAction, path);
        }

        /// <summary>
        /// Instantiate an instance of the ExcelOutputItem with a particular <see cref="IEnumerable{T}"/> collection.
        /// This object allows for the creation of a single sheet, or when used in conjunction with <see cref="ExcelOutputCollection"/>,
        /// the creation of multiple sheets within a single Workbook.
        /// <typeparam name="T">A generic type which is contained within the IEnumerable.</typeparam>
        /// <param name="enumerable">An enumerable of type <typeparam name="T"/> which contains data to be exported to Excel.</param>
        /// <param name="worksheetName">The name given to the newly generated worksheet.</param>
        /// <param name="trim">A parameter which determines whether strings should be trimmed before being written to Excel.</param>
        /// <param name="postCreationAction">Determines whether the Excel file will be opened with the data, opened and saved or just saved.</param>
        /// <param name="path">If the path parameter determines where the Excel file will be saved to if a save action is selected from the <see cref="PostCreationActions"/>.</param>
        /// </summary>
        public static ExcelOutputItem CreateInstance<T>(IEnumerable<T> enumerable, string worksheetName, bool trim, PostCreationActions postCreationAction, string path)
        {
            var excelOutputItem = new ExcelOutputItem();
            var array = enumerable as T[] ?? enumerable.ToArray();
            excelOutputItem.Type = array.First().GetType();
            excelOutputItem.Enumerable = array;
            excelOutputItem.WorksheetName = worksheetName;
            excelOutputItem.Trim = trim;
            excelOutputItem.PostCreationActions = postCreationAction;
            excelOutputItem.Path = path;
            return excelOutputItem;
        }

        internal IEnumerable Enumerable { get; private set; }

        internal string WorksheetName { get; private set; }

        internal bool Trim { get; private set; }

        internal PostCreationActions PostCreationActions { get; private set; }

        internal string Path { get; private set; }

        internal Type Type { get; private set; }
    }
}