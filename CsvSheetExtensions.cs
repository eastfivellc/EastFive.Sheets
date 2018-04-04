using System;
using System.Collections.Generic;
using System.IO;

namespace EastFive.Sheets
{
    public static class CsvSheetExtensions
    {
        public static IEnumerable<string[]> ReadRows(this Stream stream, string delimiter = ",")
        {
            stream.Seek(0, SeekOrigin.Begin);
            using (var parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(stream)
            {
                TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited,
                Delimiters = new[] { delimiter }
            })
            {
                while (!parser.EndOfData)
                {
                    string[] fields;
                    try
                    {
                        fields = parser.ReadFields();
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                    yield return fields;
                }
            }
        }

        public static IEnumerable<string[]> ReadRows(this string content, string delimiter = ",")
        {
            using (var reader = new StringReader(content))
            using (var parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(reader)
            {
                TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited,
                Delimiters = new[] { delimiter }
            })
            {
                while (!parser.EndOfData)
                {
                    string[] fields;
                    try
                    {
                        fields = parser.ReadFields();
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                    yield return fields;
                }
            }
        }

        public static TResult GetLegend<TInput,TEnum,TResult>(this TInput[] fields, Dictionary<string, TEnum> map, 
            Func<TInput, string> getFieldText,
            Func<Dictionary<TEnum, int>,TResult> onParsed)
        {
            var legend = new Dictionary<TEnum, int>();
            for (var i = 0; i < fields.Length; i++)
            {
                if (map.TryGetValue(getFieldText(fields[i]), out TEnum field))
                    legend.Add(field, i);
            }
            return onParsed(legend);
        }

        public static TResult GetColumn<TInput,TEnum,TResult>(this TInput[] fields, Dictionary<TEnum, int> legend, TEnum column,
            Func<TInput,TResult> setFunc, TResult defaultValue = default(TResult))
        {
            return legend.TryGetValue(column, out int index) ? setFunc(fields[index]) : defaultValue;
        }
    }
}
