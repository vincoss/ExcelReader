using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Xunit;


namespace Vinco.ExcelReader.Test
{
    public class ExcelReaderServiceTest
    {
        [Fact]
        public void Constructor_Throws_If_FileName_Is_Null_Test()
        {
            // Assert
            Assert.Throws<ArgumentNullException>(() =>
                                                     {
                                                         new ExcelReaderService(null);
                                                     });
        }

        [Fact]
        public void Name_Is_Equal_To_File_Name_Test()
        {
            // Act
            ExcelReaderService reader = new ExcelReaderService("Test");

            // Assert
            Assert.Equal("Test", reader.Name);
        }
    }
}
