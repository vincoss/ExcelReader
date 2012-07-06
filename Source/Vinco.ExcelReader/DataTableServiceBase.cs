using System;
using System.Data;


namespace Vinco.ExcelReader
{
    public abstract class DataTableServiceBase
    {
        public abstract DataTable GetTable(int tableIndex);

        public abstract DataTable GetTable(string tableName, Func<DataTable, bool> tableLocator = null);

        public abstract string FindTableName(int tableIndex);

        public string Name { get; protected set; }
    }
}