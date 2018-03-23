using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;

namespace UsefulDataTools
{
   public class ExcelOutputCollection : ICollection<ExcelOutputItem>
   {
      private readonly Collection<ExcelOutputItem> _collection;

      private int _dataTableCount;

      /// <summary>
      ///    Instantiate an instance of the ExcelOutputCollection with default values.
      ///    This collection allows for the creation of multiple sheets within a single Workbook.
      /// </summary>
      public ExcelOutputCollection()
      {
         _collection = new Collection<ExcelOutputItem>();
         PostCreationAction = PostCreationActions.Open;
         Path = string.Empty;
      }

      /// <summary>
      ///    Instantiate an instance of the ExcelOutputCollection. This collection allows for the creation of multiple sheets
      ///    within a single Workbook.
      ///    <param name="postCreationAction">
      ///       Determines whether the Excel file will be opened with the data, opened and saved or just saved.
      ///    </param>
      ///    <param name="path">
      ///       If the path parameter determines where the Excel file will be saved to if a save action is selected from the
      ///       <see cref="PostCreationActions" />.
      ///    </param>
      /// </summary>
      public ExcelOutputCollection(PostCreationActions postCreationAction, string path)
      {
         _collection = new Collection<ExcelOutputItem>();
         PostCreationAction = postCreationAction;
         Path = path;
      }

      public PostCreationActions PostCreationAction { get; }

      public string Path { get; }

      public int Count => _collection.Count;

      public bool IsReadOnly => false;

      public IEnumerator<ExcelOutputItem> GetEnumerator()
      {
         return _collection.GetEnumerator();
      }

      IEnumerator IEnumerable.GetEnumerator()
      {
         return GetEnumerator();
      }

      public void Clear()
      {
         _collection.Clear();
      }

      public bool Contains(ExcelOutputItem item)
      {
         if (item == null) throw new ArgumentNullException(nameof(item));
         return _collection.Contains(item);
      }

      public bool Contains(string worksheetName)
      {
         return _collection.Any(x => x.WorksheetName == worksheetName);
      }

      public void CopyTo(ExcelOutputItem[] array, int arrayIndex)
      {
         _collection.CopyTo(array, arrayIndex);
      }

      public bool Remove(ExcelOutputItem item)
      {
         if (item == null) throw new ArgumentNullException(nameof(item));
         return _collection.Remove(item);
      }

      public void Add(ExcelOutputItem item)
      {
         if (item == null) throw new ArgumentNullException(nameof(item));
         if (_collection.Any(x => x.WorksheetName == item.WorksheetName))
            throw new WorksheetNameExistsException();
         _collection.Add(item);
      }

      public void Add(params ExcelOutputItem[] items)
      {
         foreach (var item in items)
            Add(item);
      }

      public void Add(IEnumerable<ExcelOutputItem> items)
      {
         foreach (var item in items)
            Add(item);
      }

      public void Add(DataTable dataTable)
      {
         string tableName;
         if (string.IsNullOrWhiteSpace(dataTable.TableName))
         {
            do
            {
               _dataTableCount++;
               tableName = string.Concat("Table", _dataTableCount);
            } while (_collection.Any(x=>x.WorksheetName == tableName));

         }
         else
            tableName = dataTable.TableName;

         _collection.Add(ExcelOutputItem.CreateInstance(dataTable, tableName, true, PostCreationActions.Open, string.Empty));
      }

      public void Add(DataSet dataSet)
      {
         foreach (DataTable dataTable in dataSet.Tables)
            Add(dataTable);
      }
   }
}