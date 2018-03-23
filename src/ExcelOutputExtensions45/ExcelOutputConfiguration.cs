using System.Globalization;
using System.Threading;

namespace UsefulDataTools
{
   public static class ExcelOutputConfiguration
   {
      public static CultureInfo DefaultCulture { get; set; }
      public static string DefaultDateTimeFormat { get; set; }
      public static string DefaultDateFormat { get; set; }
      public static string DefaultTimeFormat { get; set; }
      public static string DefaultWorksheetName { get; set; }

      static ExcelOutputConfiguration()
      {
         DefaultCulture = Thread.CurrentThread.CurrentCulture;
         DefaultDateTimeFormat = DefaultCulture.DateTimeFormat.ShortDatePattern + " " + Thread.CurrentThread.CurrentCulture.DateTimeFormat.LongTimePattern;
         DefaultDateFormat = DefaultCulture.DateTimeFormat.ShortDatePattern;
         DefaultTimeFormat = DefaultCulture.DateTimeFormat.LongTimePattern;
         DefaultWorksheetName = "Data";
      }
   }
}