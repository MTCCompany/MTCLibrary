using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MTC.MTCLibrary;
using NUnit.Framework;

namespace MTC.MTCLibraryTest
{
    /// <summary>
    /// ExcelWrapperのテスト
    /// </summary>
    [TestFixture]
    public class ExcelWrapperTest
    {
        ///<summary>バージョン取得テスト</summary>
        [Test]
        public void Version_Test()
        {
            ExcelWrapper xls = new ExcelWrapper();
            Console.WriteLine("Excelのバージョン:" + xls.Version.ToString());
            //インストールされているOfficeによってバージョンによって変更する
            Assert.True(xls.Version.ToString() == "11.0");
        }

    }
}
