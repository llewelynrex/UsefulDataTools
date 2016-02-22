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
            var type = typeof (T);
            string output;

            if (type.IsSimpleType())
                output = string.Join(separator.ToString(), input);

            else
            {
                var properties = type.GetProperties();
                var fields = type.GetFields();
                var stringBuilder = new StringBuilder();

                stringBuilder.AppendLine(
                                         string.Concat(
                                                       string.Join(separator.ToString(), properties.Where(p => p.PropertyType.IsSimpleType()).Select(p => p.Name).ToArray()),
                                                       fields.Any(f => f.FieldType.IsSimpleType()) ? separator.ToString() : string.Empty,
                                                       string.Join(separator.ToString(), fields.Where(f => f.FieldType.IsSimpleType()).Select(f => f.Name).ToArray())
                                             )
                    );

                foreach (var element in input)
                {
                    stringBuilder.AppendLine(
                                             string.Concat(
                                                           string.Join(separator.ToString(), properties.Where(p => p.PropertyType.IsSimpleType()).Select(p => p.GetValue(element)).ToArray()),
                                                           fields.Any(f => f.FieldType.IsSimpleType()) ? separator.ToString() : string.Empty,
                                                           string.Join(separator.ToString(), fields.Where(f => f.FieldType.IsSimpleType()).Select(f => f.GetValue(element)).ToArray())
                                                 )
                        );
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