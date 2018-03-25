using System;
using System.Data;
using UsefulDataTools;

namespace Test
{
   class Program
   {
      static void Main(string[] args)
      {
         var dataTable = new DataTable("TestTable");
         dataTable.Columns.Add(new DataColumn("id") {DataType = typeof(int), ReadOnly = true, AutoIncrement = true, AutoIncrementSeed = 0, AutoIncrementStep = 1});
         dataTable.Columns.Add(new DataColumn("Text") { DataType = typeof(string)});
         dataTable.Columns.Add(new DataColumn("Bool") { DataType = typeof(bool)});
         dataTable.Columns.Add(new DataColumn("Date") { DataType = typeof(DateTime)});
         dataTable.PrimaryKey = new[]{dataTable.Columns[0]};

         var dataSet = new DataSet();
         dataSet.Tables.Add(dataTable);

         var dataRow1 = dataTable.NewRow();
         dataRow1[1] = "Test1";
         dataRow1[2] = false;
         dataRow1[3] = DateTime.Now;
         dataTable.Rows.Add(dataRow1);

         var dataRow2 = dataTable.NewRow();
         dataRow2[1] = "Test2";
         dataRow2[2] = true;
         dataRow2[3] = DateTime.Now.AddDays(1);
         dataTable.Rows.Add(dataRow2);

         dataTable.ToExcel();

         //dataSet.ToExcel(true, PostCreationActions.Open, string.Empty);
      }
   }
}
