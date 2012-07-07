using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Xunit;


namespace Vinco.ExcelReader.Test
{
    public class ObjectExtensionsTest
    {
        [Fact]
        public void GetDataTableRawColumns_Throws_If_Contains_Null_Or_Empty_Column_Test()
        {
            // Arrange
            DataRow row = CreateDataTable().Rows[1];
            row[2] = "";

            // Assert
            Assert.Throws<InvalidOperationException>(() =>
                                                         {
                                                             ObjectExtensions.GetDataTableRawColumns(row, null);
                                                         });
        }

        [Fact]
        public void GetDataTableRawColumns_Test()
        {
            // Arrange
            DataRow row = CreateDataTable().Rows[1];

            // Act
            IEnumerable<string> columns = ObjectExtensions.GetDataTableRawColumns(row, "###");

            // Assert
            Assert.True(new[] {"Bool", "Byte", "Sbyte"}.SequenceEqual(columns));
        }

        [Fact]
        public void IsCompatible_Test()
        {
            // Assert
            Assert.True(ObjectExtensions.IsCompatible(5, 7, 70));
        }

        [Fact]
        public void FindColumnsRowStartIndex_Test()
        {
            // Arrange
            var table = CreateDataTable();
            var columnsToMatch = new[] {"Bool", "Byte", "Sbyte", "Char"};

            // Act
            int index = ObjectExtensions.FindColumnsRowStartIndex(table, columnsToMatch, "###");

            // Asseert
            Assert.Equal(1, index);
        }

        [Fact]
        public void FindColumnsRowStartIndex_Cant_Find_If_No_Rows_Or_Columns_Test()
        { 
            // Arrange
            var columnsToMatch = new[] { "Bool", "Byte", "Sbyte" };

            // Act
            int index = ObjectExtensions.FindColumnsRowStartIndex(new DataTable(), columnsToMatch, "###");

            // Asseert
            Assert.Equal(-1, index);
        }

        [Fact]
        public void FindColumnsRowStartIndex_Cant_Find_If_Column_Is_Null_Or_End_Car_Test()
        {
            // Arrange
            var table = CreateDataTable();
            var columnsToMatch = new[] { "Bool", "Byte", "Sbyte" };
            table.Rows[1][2] = "###";

            // Act
            int index = ObjectExtensions.FindColumnsRowStartIndex(table, columnsToMatch, "###");

            // Asseert
            Assert.Equal(-1, index);
        }

        [Fact]
        public void FindColumnsRowStartIndex_Cant_Find_Did_Not_Match_Test()
        {
            // Arrange
            var table = CreateDataTable();
            var columnsToMatch = new[] { "String", "Byte", "Decimal" };

            // Act
            int index = ObjectExtensions.FindColumnsRowStartIndex(table, columnsToMatch, "###");

            // Asseert
            Assert.Equal(-1, index);
        }

        [Fact]
        public void GetPublicPropertyNames_Test()
        {
            // Act
            var properties = ObjectExtensions.GetPublicPropertyNames(typeof (TestContract));

            // Assert
            Assert.True(properties.SequenceEqual(new[] {"Bool", "Byte"}));
        }

        [Fact]
        public void SetPropertyInfo_Test()
        {
            Assert.True(false, "TODO: complete tests");
        }

        private static DataTable CreateDataTable()
        {
            var table = new DataTable();
            table.Columns.Add("0");
            table.Columns.Add("1");
            table.Columns.Add("2");
            table.Columns.Add("3");
            table.Columns.Add("4");

            // Make one row empty and then columns
            table.Rows.Add("");

            table.Rows.Add("Bool", "Byte", "Sbyte", "###", "Char");

            return table;
        }

        public class TestContract
        {
            public bool Bool { get; set; }
            public byte Byte { get; set; }
            public sbyte SbyteReadOnly { get; private set; }
        }
    }
}
