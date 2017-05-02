using System;
using System.Collections.ObjectModel;
using UsefulDataTools;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            var collection = new Collection<Test>();
            collection.Add(new Test
            {
                Accounting = 1000,
                Currency = 1000,
                DateProperty = DateTime.Today,
                DateTimeProperty = DateTime.Now,
                Numeric = 1000,
                StockCode = "A100",
                TimeProperty = DateTime.Now
            });
            collection.Add(new Test
            {
                Accounting = -1000,
                Currency = -1000,
                DateProperty = DateTime.Today,
                DateTimeProperty = DateTime.Now,
                Numeric = -1000,
                StockCode = "A200",
                TimeProperty = DateTime.Now
            });
            collection.Add(new Test
            {
                Accounting = 0,
                Currency = 0,
                DateProperty = DateTime.Today,
                DateTimeProperty = DateTime.Now,
                Numeric = 0,
                StockCode = "A300",
                TimeProperty = DateTime.Now
            });

            collection.ToExcel();
        }
    }

    public class Test
    {
        [ColumnNumberFormat(ExcelNumberFormats.Text)]
        public string StockCode { get; set; }
        [ColumnNumberFormat(ExcelNumberFormats.DateTime)]
        public DateTime DateTimeProperty { get; set; }
        [ColumnNumberFormat(ExcelNumberFormats.Date)]
        public DateTime DateProperty { get; set; }
        [ColumnNumberFormat(ExcelNumberFormats.Time)]
        public DateTime TimeProperty { get; set; }
        [ColumnNumberFormat(ExcelNumberFormats.Accounting)]
        public decimal Accounting { get; set; }
        [ColumnNumberFormat(ExcelNumberFormats.Currency)]
        public decimal Currency { get; set; }
        [ColumnNumberFormat(ExcelNumberFormats.Number)]
        public decimal Numeric { get; set; }
    }
}
