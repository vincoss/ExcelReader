using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;


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
            IList<string> columns = EnsureColumnsToMap(type, columnsToMatch).ToList();

            // If read exel table name does not might find right dataset, so we find by columns comparison.
            //Func<DataTable, bool> tableLocator = (dataTable) =>
            //                                         {
            //                                             if (dataTable == null)
            //                                             {
            //                                                 throw new ArgumentNullException("dataTable");
            //                                             }
            //                                             int index = ObjectExtensions.FindColumnsRowStartIndex(dataTable, columns, this.EndChar);
            //                                             if (index >= 0)
            //                                             {
            //                                                 DataRow columnRow = dataTable.Rows[index];
            //                                                 List<string> rawColumns = ObjectExtensions.GetDataTableRawColumns(columnRow, this.EndChar).ToList();

            //                                                 // Try to match table
            //                                                 int absolute = columns.Count;
            //                                                 int match = (from r in rawColumns
            //                                                              from c in columns
            //                                                              where r.Equals(c, StringComparison.OrdinalIgnoreCase)
            //                                                              select r).Count();

            //                                                 bool result = ObjectExtensions.IsCompatible(match, absolute, 70);
            //                                                 return result;
            //                                             }
            //                                             return false;
            //                                         };

            DataTable table = _dataTableReaderService.GetTable(tableName);
            int columnsRowStartIndex = ObjectExtensions.FindColumnsRowStartIndex(table, columns, this.EndChar);

            // Start index should be equal or greater than zero.
            if (columnsRowStartIndex >= 0)
            {
                // This row should be columns row.
                var columnRow = table.Rows[columnsRowStartIndex];
                var dataTableRawColumns = ObjectExtensions.GetDataTableRawColumns(columnRow, this.EndChar).ToList();

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
                                AddMessage(string.Format("Parsing column : [{0}], Value : [{1}]", key, columnValue));

                                // Parse value
                                ObjectExtensions.SetPropertyInfo(item, key, columnValue);
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

        private IEnumerable<string> EnsureColumnsToMap(Type type, IEnumerable<string> columnsToMatch)
        {
            if (columnsToMatch == null || columnsToMatch.Any() == false)
            {
                // If there are no columns specified get all public properties from type.
                columnsToMatch = ObjectExtensions.GetPublicPropertyNames(type);
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

        private void EnsureColumnMatch(IEnumerable<string> dataTableRawColumns, IEnumerable<string> columnsToMatch)
        {
            if (dataTableRawColumns.Any() == false || columnsToMatch.Any() == false)
            {
                throw new InvalidOperationException("Could not find any columns.");
            }

            // Check wheter there are any columns to match, if no throw that there nothig to harvest.
            foreach (var column in columnsToMatch)
            {
                if (dataTableRawColumns.Any(x => string.Equals(x, column, StringComparison.OrdinalIgnoreCase)) == false)
                {
                   throw new InvalidOperationException(string.Format("Unmapped [{0}] column.", column));
                }
            }
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
