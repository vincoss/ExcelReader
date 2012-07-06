using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Reflection;


namespace Vinco.ExcelReader
{
    public class DataTableHarvester
    {
        private readonly DataTableServiceBase _dataTableReaderService;

        public DataTableHarvester(DataTableServiceBase dataTableReaderService)
        {
            if (dataTableReaderService == null)
            {
                throw new ArgumentNullException("dataTableReaderService");
            }
            this.EndChar = "###";
            this._dataTableReaderService = dataTableReaderService;
        }

        #region Public methods

        public void PerRowHarvest<T>(int tableIndex, Action<T> callback, IEnumerable<string> columnsToMatch = null) where T : class, new()
        {
            if (tableIndex < 0)
            {
                throw new ArgumentOutOfRangeException("tableIndex");
            }
            PerRowHarvest(_dataTableReaderService.FindTableName(tableIndex), callback, columnsToMatch);
        }

        public void PerRowHarvest<T>(string tableName, Action<T> callback, IEnumerable<string> columnsToMatch = null) where T : class, new()
        {
            if (string.IsNullOrWhiteSpace(tableName))
            {
                throw new ArgumentNullException("tableName");
            }
            if(callback == null)
            {
                throw new ArgumentNullException("callback");
            }
            ReadInternal(tableName, columnsToMatch, callback);
        }

        public IEnumerable<T> Harvest<T>(int tableIndex, IEnumerable<string> columnsToMatch = null) where T : class, new()
        {
            if (tableIndex < 0)
            {
                throw new ArgumentOutOfRangeException("tableIndex");
            }
            return Harvest<T>(_dataTableReaderService.FindTableName(tableIndex), columnsToMatch);
        }

        public IEnumerable<T> Harvest<T>(string tableName, IEnumerable<string> columnsToMatch = null) where T : class, new()
        {
            if (string.IsNullOrWhiteSpace(tableName))
            {
                throw new ArgumentNullException("tableName");
            }
            return ReadInternal<T>(tableName, columnsToMatch, null);
        }

        public void AddMessage(string message)
        {
            if (DiagnosticsCallback != null && string.IsNullOrWhiteSpace(message) == false)
            {
                DiagnosticsCallback(message);
            }
        }

        #endregion

        #region Private methods

        private IEnumerable<T> ReadInternal<T>(string tableName, IEnumerable<string> columnsToMatch, Action<T> callback) where T : class, new()
        {
            Type type = typeof (T);
            IList<T> items = new List<T>();
            IList<string> columns = EnsureColumnMap(type, columnsToMatch);

            // If read exel table name does not might find right dataset, so we find by columns comparison.
            Func<DataTable, bool> tableLocator = (dataTable) =>
                                                     {
                                                         if (dataTable == null)
                                                         {
                                                             throw new ArgumentNullException("dataTable");
                                                         }
                                                         int index = FindColumnsRowStartIndex(dataTable, columns);
                                                         if (index >= 0)
                                                         {
                                                             DataRow columnRow = dataTable.Rows[index];
                                                             List<string> rawColumns = GetDataTableRawColumns(columnRow).ToList();
                                                             bool result = true;//rawColumns.OrderBy(x => x).SequenceEqual(columns.OrderBy(x => x), StringComparer.Ordinal);
                                                             return result;
                                                         }
                                                         return false;
                                                     };

            DataTable table = _dataTableReaderService.GetTable(tableName, tableLocator);
            int columnsRowStartIndex = FindColumnsRowStartIndex(table, columns);

            // Start index should be equal or greater than zero.
            if (columnsRowStartIndex >= 0)
            {
                // This row should be columns row.
                var columnRow = table.Rows[columnsRowStartIndex];
                var dataTableRawColumns = GetDataTableRawColumns(columnRow);

                // Make sure that we match any columns.
                EnsureColumnMatch(dataTableRawColumns, columns);
                
                // This row should be data row.
                int dataRowStartIndex = columnsRowStartIndex + 1;

                if (table.Rows.Count >= dataRowStartIndex)
                {
                    int columnsCount = dataTableRawColumns.Count;

                    // Read rows
                    for (int i = dataRowStartIndex; i < table.Rows.Count; i++)
                    {
                        // Diagnostics information
                        AddMessage(string.Format("Reading [{0}], table [{1}], row [{2}]", _dataTableReaderService.Name, tableName, (i + 1)));

                        T item = new T();
                        bool end = false;

                        // Read columns
                        for (int j = 0; j < columnsCount; j++)
                        {
                            string key = dataTableRawColumns[j];
                            object columnValue = table.Rows[i][j];

                            // Terminate if end char found.
                            if (columnValue != null && columnValue.ToString().StartsWith(this.EndChar))
                            {
                                end = true;
                                break;
                            }
                            if (columnValue is DBNull)
                            {
                                columnValue = null;
                            }

                            if (columns.Any(x => string.Equals(x, key, StringComparison.OrdinalIgnoreCase)))
                            {
                                // Parse value
                                SetPropertyInfo(item, key, columnValue);
                            }
                        }
                        if (end)
                        {
                            break;
                        }
                        if (callback == null)
                        {
                            items.Add(item);
                        }
                        if (callback != null)
                        {
                            callback(item);
                        }
                    }
                }
            }
            return items;
        }

        private IList<string> EnsureColumnMap(Type type, IEnumerable<string> columnsToMatch)
        {
            if (columnsToMatch == null || columnsToMatch.Any() == false)
            {
                // If there are no columns specified get all public properties from type.
                columnsToMatch = GetPublicPropertyNames(type);
            }
            var columns = new List<string>();
            foreach (var columnName in columnsToMatch)
            {
                // Ensure columns are valid.
                if (string.IsNullOrWhiteSpace(columnName))
                {
                    throw new InvalidOperationException("Column name can't be null or empty.");
                }
                string column = columnName.Trim();
                if (columns.Any(x => string.Equals(x, column, StringComparison.OrdinalIgnoreCase)))
                {
                    throw new InvalidOperationException(string.Format("Column [{0}] already exists.", column));
                }
                columns.Add(column);
            }
            return columns;
        }

        private int FindColumnsRowStartIndex(DataTable table, IEnumerable<string> columnsToMatch)
        {
            if (table.Rows.Count > 0 && table.Columns.Count > 0)
            {
                int count = table.Rows.Count;
                for (int i = 0; i < count; i++)
                {
                    DataRow row = table.Rows[i];
                    int columnCount = row.ItemArray.Count();
                    if (columnCount > 0)
                    {
                        for (int j = 0; j < columnCount; j++)
                        {
                            object columnName = row.ItemArray[j];

                            // If null or starts with end char then terminate, there nothing to read there.
                            if (columnName == null || columnName is DBNull || columnName.ToString().Trim().StartsWith(this.EndChar))
                            {
                                break;
                            }
                            
                            // Check wheter we match any columns, if yes this can be our starting row index.
                            bool result = columnsToMatch.Any(x => string.Equals(x, (string)columnName, StringComparison.OrdinalIgnoreCase));
                            if (result)
                            {
                                return i;
                            }
                        }
                    }
                }
            }
            return -1;
        }

        private void EnsureColumnMatch(IList<string> dataTableRawColumns, IList<string> columnsToMatch)
        {
            if (dataTableRawColumns.Any() == false || columnsToMatch.Any() == false)
            {
                throw new InvalidOperationException("Could not find any columns.");
            }
            foreach (var column in columnsToMatch)
            {
                if (dataTableRawColumns.Any(x => string.Equals(x, column, StringComparison.OrdinalIgnoreCase)) == false)
                {
                   throw new InvalidOperationException(string.Format("Unmapped [{0}] column.", column));
                }
            }
        }

        private IList<string> GetDataTableRawColumns(DataRow row)
        {
            var columns = new List<string>();
            int columnCount = row.ItemArray.Count();
            if (columnCount > 0)
            {
                for (int i = 0; i < columnCount; i++)
                {
                    object value = row.ItemArray[i];
                    if (value == null || value is DBNull)
                    {
                        throw new InvalidOperationException("Missing column name.");
                    }
                    string columnName = ((string)value).Trim();

                    // Terminate if end char found.
                    if (string.Equals(columnName, this.EndChar, StringComparison.Ordinal))
                    {
                        break;
                    }
                    columns.Add(columnName);
                }
            }
            return columns;
        }

        private void SetPropertyInfo(object entity, string propertyName, object value)
        {
            AddMessage(string.Format("Parsing column : [{0}], Value : [{1}]", propertyName, value));

            string stringValue = null;
            if (value != null && !(value is DBNull))
            {
                stringValue = value.ToString();
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

        private static IEnumerable<string> GetPublicPropertyNames(Type type)
        {
            return (from x in type.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly).Where(p => p.GetSetMethod() != null) select x.Name).ToList();
        }

        #endregion

        public Action<string> DiagnosticsCallback;

        /// <summary>
        /// Get or set terminate character.
        /// <c>###</c>
        /// </summary>
        public string EndChar { get; set; }

    }
}
