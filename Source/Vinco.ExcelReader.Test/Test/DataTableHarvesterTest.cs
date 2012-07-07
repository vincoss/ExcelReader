using System;
using System.Globalization;
using System.Linq;
using System.Threading;
using Xunit;
using Moq;
using System.Collections.Generic;


namespace Vinco.ExcelReader.Test
{
    public class DataTableHarvesterTest
    {
        public DataTableHarvesterTest()
        {
            var culture = new CultureInfo("en-AU");
            Thread.CurrentThread.CurrentCulture = culture;
        }

        [Fact]
        public void Constructor_Throws_If_Service_Is_Null_Test()
        {
            // Assert
            Assert.Throws<ArgumentNullException>(() =>
                                                     {
                                                         new DataTableHarvester(null);
                                                     });
        }

        [Fact]
        public void EndChar_Test()
        {
            // Arrange
            var service = new Mock<DataTableServiceBase>();

            // Act
            var harvester = new DataTableHarvester(service.Object);

            // Assert
            Assert.Equal("###", harvester.EndChar);
        }

        [Fact]
        public void DiagnosticsCallback_Null_Test()
        {
            // Arrange
            var service = new Mock<DataTableServiceBase>();

            // Act
            var harvester = new DataTableHarvester(service.Object);

            // Assert
            Assert.Null(harvester.DiagnosticsCallback);
        }

        #region PerRowHarvest

        [Fact]
        public void PerRowHarvest_Throws_If_Index_Is_Less_Than_Zero_Test()
        {
            // Act
            var harvester = CreateHarvester();

            // Assert
            Assert.Throws<ArgumentOutOfRangeException>(() => harvester.PerRowHarvest<TestContract>(-1, (x) => { }));
        }

        [Fact]
        public void PerRowHarvest_Throws_If_Table_Name_Is_Null_Or_Empty_Test()
        {
            // Act
            var harvester = CreateHarvester();

            // Assert
            Assert.Throws<ArgumentNullException>(() => harvester.PerRowHarvest<TestContract>(null, (x) => { }));
        }

        [Fact]
        public void PerRowHarvest_Throws_If_Callback_Is_Null_Test()
        {
            // Act
            var harvester = CreateHarvester();

            // Assert
            Assert.Throws<ArgumentNullException>(() => harvester.PerRowHarvest<TestContract>(null, null));
        }

        [Fact]
        public void PerRowHarvest_Test()
        {
            // Arrange
            string fileName = Environment.CurrentDirectory + @"\Contracts.xlsx";
            var sevice = new ExcelReaderService(fileName);
            var harvester = new DataTableHarvester(sevice);
            harvester.DiagnosticsCallback = Console.WriteLine;

            // Assert
            harvester.PerRowHarvest<TestContract>(0, Assert.NotNull);
        } 

        #endregion

        #region Harvest

        [Fact]
        public void Harvest_Throws_If_Index_Is_Less_Than_Zero_Test()
        {
            // Act
            var harvester = CreateHarvester();

            // Assert
            Assert.Throws<ArgumentOutOfRangeException>(() => harvester.Harvest<TestContract>(-1));
        }

        [Fact]
        public void Harvest_Throws_If_Table_Name_Is_Null_Or_Empty_Test()
        {
            // Act
            var harvester = CreateHarvester();

            // Assert
            Assert.Throws<ArgumentNullException>(() => harvester.Harvest<TestContract>(null));
        } 

        [Fact]
        public void Harvest_Test()
        {
            // Arrange
            string fileName = Environment.CurrentDirectory + @"\Contracts.xlsx";
            var sevice = new ExcelReaderService(fileName);
            var harvester = new DataTableHarvester(sevice);
            harvester.DiagnosticsCallback = Console.WriteLine;

            List<string> colums = new List<string>();
            colums.Add("String");
            colums.Add("Sbyte");

            // Act
            IEnumerable<TestContract> items = harvester.Harvest<TestContract>(0, colums);

            // Assert
            Assert.True(items.Count() == 4);
        }

        #endregion

        [Fact]
        public void AddMessage_Test()
        {
            // Act
            var harvester = CreateHarvester();

            // Assert
            harvester.DiagnosticsCallback = (x) => Assert.Equal("Test", x);
        }

        private DataTableHarvester CreateHarvester()
        {
            return new DataTableHarvester(CreateService());
        }

        private DataTableServiceBase CreateService()
        {
            var service = new Mock<DataTableServiceBase>();
            return service.Object;
        }

        class TestContract
        {
            public bool? Bool { get; set; }
            public byte? Byte { get; set; }
            public sbyte? Sbyte { get; set; }
            public char? Char { get; set; }
            public decimal? Decimal { get; set; }
            public double? Double { get; set; }
            public float? Float { get; set; }
            public int? Int { get; set; }
            public uint? Uint { get; set; }
            public long? Long { get; set; }
            public ulong? Ulong { get; set; }
            public short? Short { get; set; }
            public ushort? Ushort { get; set; }
            public string String { get; set; }
            public DateTime? Date { get; set; }
        }
    }
}
