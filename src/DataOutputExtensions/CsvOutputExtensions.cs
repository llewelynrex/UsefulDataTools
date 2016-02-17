using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace UsefulDataTools
{
    public static class CsvOutputExtensions
    {

        public static string ToCsv<T>(this IEnumerable<T> input, char separator = ',', string path = null)
        {
            var type = typeof(T);
            var properties = type.GetProperties();
            string output;

            if (type.IsSimpleType())
                output = string.Join(separator.ToString(), input);

            else
            {
                var stringBuilder = new StringBuilder();
                stringBuilder.AppendLine(string.Join(separator.ToString(), properties.Where(p => p.PropertyType.IsSimpleType()).Select(p => p.Name).ToArray()));
                foreach (var element in input)
                {
                    stringBuilder.AppendLine(string.Join(separator.ToString(), properties.Where(p => p.PropertyType.IsSimpleType()).Select(p => p.GetValue(element)).ToArray()));
                }

                output = stringBuilder.ToString();
            }

            if (!string.IsNullOrEmpty(path))
                WriteToPath(output, path);

            return output;
        }

        private static void WriteToPath(string outputString, string path)
        {
            using (var fileStream = new FileStream(path, FileMode.Create))
            {
                using (var streamWriter = new StreamWriter(fileStream))
                {
                    streamWriter.Write(outputString);
                    streamWriter.Flush();
                }
            }
        }
    }
}