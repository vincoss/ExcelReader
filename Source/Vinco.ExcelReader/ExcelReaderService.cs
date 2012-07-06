using System;
using System.Linq;
using System.IO;
using System.Data;
using Excel;


namespace Vinco.ExcelReader
{
    public class ExcelReaderService : DataTableServiceBase
    {
        public ExcelReaderService(string fileName)
        {
            if(string.IsNullOrWhiteSpace(fileName))
            {
                throw new ArgumentNullException("fileName");
            }
            this.Name = fileName;
        }

        public override string FindTableName(int tableIndex)
        {
            var table = GetTable(tableIndex);
            if (table != null)
            {
                return table.TableName;
            }
            return null;
        }

        public override DataTable GetTable(int tableIndex)
        {
            if (tableIndex < 0)
            {
                throw new ArgumentOutOfRangeException("tableIndex");
            }
            IExcelDataReader excelReader = GetReader(Name);
            DataTable table = excelReader.AsDataSet().Tables[tableIndex];
            return table;
        }

        public override DataTable GetTable(string tableName, Func<DataTable, bool> tableLocator = null)
        {
            if (string.IsNullOrWhiteSpace(tableName))
            {
                throw new ArgumentNullException("tableName");
            }
            IExcelDataReader excelReader = GetReader(Name);
            foreach (DataTable table in excelReader.AsDataSet().Tables)
            {
                if(tableLocator != null)
                {
                    // In case tables were reorderd and index was not upated data set wont match actual table name. 
                    // Use custom table locator to match table on columns.
                    if (tableLocator(table))
                    {
                        return table;
                    }
                }
                else
                {
                    if (string.Equals(table.TableName, tableName, StringComparison.OrdinalIgnoreCase))
                    {
                        return table;
                    }
                }
            }
            return null;
        }

        protected virtual IExcelDataReader GetReader(string fileName)
        {
            const string xls = "xls";
            const string xlsx = "xlsx";
            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new ArgumentNullException("fileName");
            }
            string extension = Path.GetExtension(fileName);
            if (extension != null)
            {
                extension = extension.Trim('.');
            }
            if (new[] { xls, xlsx }.Any(x => string.Equals(x, extension, StringComparison.OrdinalIgnoreCase)) == false)
            {
                throw new NotSupportedException(string.Format("Only [{0}] and [{1}] file types are supported.", xls, xlsx));
            }
            var stream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
            IExcelDataReader excelReader = string.Equals(extension, xlsx, StringComparison.OrdinalIgnoreCase) ? ExcelReaderFactory.CreateOpenXmlReader(stream) : ExcelReaderFactory.CreateBinaryReader(stream);
            return excelReader;
        }
    }
}
