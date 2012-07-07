using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Data;


namespace Vinco.ExcelReader
{
    public static class ObjectExtensions
    {
        public static IEnumerable<string> GetPublicPropertyNames(Type type)
        {
            if(type == null)
            {
                throw new ArgumentNullException("type");
            }
            return (from x in type.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly).Where(p => p.GetSetMethod() != null) select x.Name).ToList();
        }

        public static void SetPropertyInfo(object entity, string propertyName, object value)
        {
            string stringValue = null;
            if (value != null && !(value is DBNull))
            {
                stringValue = value.ToString().Trim();
            }

            var propertyDescriptorCollection = TypeDescriptor.GetProperties(entity);
            var propertyDescriptor = propertyDescriptorCollection[propertyName];

            // Parse boolean types.
            if (propertyDescriptor.PropertyType == typeof(bool) || propertyDescriptor.PropertyType == typeof(bool?))
            {
                // Nullable boolean
                if (propertyDescriptor.PropertyType == typeof(bool?) && string.IsNullOrWhiteSpace(stringValue))
                {
                    propertyDescriptor.SetValue(entity, null);
                    return;
                }

                // If set to false if null or empty.
                if (string.IsNullOrWhiteSpace(stringValue))
                {
                    propertyDescriptor.SetValue(entity, false);
                    return;
                }

                // Integar values
                if (stringValue == "1" || stringValue == "0")
                {
                    propertyDescriptor.SetValue(entity, stringValue == "1");
                    return;
                }
            }
            if (propertyDescriptor.PropertyType == typeof(DateTime) || propertyDescriptor.PropertyType == typeof(DateTime?))
            {
                DateTime? date = null;
                if (string.IsNullOrWhiteSpace(stringValue) == false)
                {
                    try
                    {
                        date = DateTime.Parse(stringValue);
                    }
                    catch (FormatException) // Heh this if for excel 97-2007
                    {
                        date = DateTime.FromOADate(Double.Parse(stringValue));
                    }
                }
                propertyDescriptor.SetValue(entity, date);
                return;
            }
            propertyDescriptor.SetValue(entity, propertyDescriptor.Converter.ConvertFromInvariantString(stringValue));
        }

        public static int FindColumnsRowStartIndex(DataTable table, IEnumerable<string> columnsToMatch, string endChar)
        {
            if (table.Rows.Count > 0 && table.Columns.Count > 0 && columnsToMatch.Any())
            {
                int count = table.Rows.Count;
                int absolute = columnsToMatch.Count();
                for (int i = 0; i < count; i++)
                {
                    int match = 0;
                    DataRow row = table.Rows[i];
                    int columnCount = row.ItemArray.Count();
                    for (int j = 0; j < columnCount; j++)
                    {
                        object expectedColumnName = row.ItemArray[j];

                        // If null or starts with end char then terminate, there nothing to read there.
                        if (expectedColumnName == null || expectedColumnName is DBNull || expectedColumnName.ToString().Trim().StartsWith(endChar))
                        {
                            break;
                        }

                        // Check wheter we match any columns.
                        bool result = columnsToMatch.Any(x => string.Equals(x, (string)expectedColumnName, StringComparison.OrdinalIgnoreCase));
                        if (result)
                        {
                            match++;
                        }

                        // If we find match then this can be our starting row index. 
                        if (IsCompatible(match, absolute, 70))
                        {
                            return i;
                        }
                    }
                }
            }
            return -1;
        }

        public static bool IsCompatible(decimal match, decimal absolute, decimal percent)
        {
            return (((match / absolute) * 100) > percent);
        }

        public static IEnumerable<string> GetDataTableRawColumns(DataRow row, string endChar)
        {
            var columns = new List<string>();
            int columnCount = row.ItemArray.Count();
            for (int i = 0; i < columnCount; i++)
            {
                object value = row.ItemArray[i];
                if (value == null || value is DBNull || string.IsNullOrWhiteSpace(((string)value).Trim()))
                {
                    throw new InvalidOperationException("Missing column name.");
                }
                // Terminate if end char found.
                string columnName = ((string)value).Trim();
                if (string.Equals(columnName, endChar, StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }
                columns.Add(columnName);
            }
            return columns;
        }
    }
}
