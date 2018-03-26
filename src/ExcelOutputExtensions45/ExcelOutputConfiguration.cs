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
         DefaultCulture = Thread.CurrentThread.CurrentUICulture;
         DefaultDateFormat = DefaultCulture.DateTimeFormat.ShortDatePattern;
         DefaultTimeFormat = Thread.CurrentThread.CurrentCulture.DateTimeFormat.LongTimePattern.Replace(" tt", string.Empty);
         DefaultDateTimeFormat = string.Concat(DefaultDateFormat, " ", DefaultTimeFormat);
         DefaultWorksheetName = "Data";
      }
   }
}