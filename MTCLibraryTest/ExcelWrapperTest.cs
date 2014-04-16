using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MTC.MTCLibrary;
using NUnit.Framework;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace MTC.MTCLibraryTest
{
    /// <summary>
    /// ExcelWrapperのテスト
    /// </summary>
    [TestFixture]
    public class ExcelWrapperTest
    {
        string tempPath = Environment.ExpandEnvironmentVariables("%temp%");

        [TestFixtureSetUp]
        public void init()
        {
            CreateExcel("\\ExcelWrapperTest001.xls");
        }

        public void CreateExcel(string xlsName)
        {
            //テスト用に使用するExcelファイル作成
            using (ExcelWrapper xls = new ExcelWrapper())
            {
                string xlsPath = tempPath + xlsName;
                File.Delete(xlsPath);
                xls.AddBook();
                xls.SaveAs(xlsPath);
            }
        }

        ///<summary>
        ///バージョン取得テスト
        ///【テスト】Version
        ///</summary>
        [Test]
        public void Version()
        {
            using (ExcelWrapper xls = new ExcelWrapper())
            {
                //インストールされているOfficeによってバージョンによって変更する
                Assert.True(xls.Version.ToString() == "11.0");
            }
        }

        ///<summary>
        ///Excelオープン及びExcelアプリケーション終了テスト
        ///【テスト】Open,Dispose
        ///</summary>
        [Test]
        public void OpenDispose()
        {
            Process[] ps;
            using (ExcelWrapper xls = new ExcelWrapper())
            {
                xls.Open(tempPath + "\\ExcelWrapperTest001.xls");
                ps = Process.GetProcessesByName("Excel");
                Assert.True(ps.Length >= 1);
            }
            System.Threading.Thread.Sleep(100);
            Assert.True(ps.Length > Process.GetProcessesByName("Excel").Length);
        }

        ///<summary>
        ///Excel新規作成及び保存テスト
        ///【テスト】AddBook,SaveAd
        ///</summary>
        [Test]
        public void AddBookSaveAs()
        {
            using (ExcelWrapper xls = new ExcelWrapper())
            {
                string xlsPath = tempPath + "\\ExcelWrapperTest_addbook.xls";
                File.Delete(xlsPath);
                xls.AddBook();
                xls.SaveAs(xlsPath);
                Assert.True(File.Exists(xlsPath));
            }
        }

        ///<summary>
        ///Excelファイル閉じるテスト
        ///【テスト】
        ///</summary>
        [Test]
        public void Close()
        {
            using (ExcelWrapper xls = new ExcelWrapper())
            {
                xls.AddBook();
                xls.Close();
                Assert.Throws <COMException>(() => xls.Save());
            }

        }

        ///<summary>
        /// Rangeテストのシート名称用
        ///【テスト】GetRangeValue,SetRangeValue
        ///</summary>
        [Test]
        public void RangeValueName()
        {
            string[,] setValue = { { "A1", "B1" }, { "A2", "B2" }, { "A3", "B3" } };
            object[,] tmpValue;
            using (ExcelWrapper xls = new ExcelWrapper())
            {
                xls.AddBook();
                xls.SetRangeValue("Sheet1", "A1:B3", setValue);
                tmpValue = (object[,])xls.GetRangeValue("Sheet1", "A1:B3");
                Assert.True(tmpValue.Length == setValue.Length);
                for (int i = 0; i < 3; i++)
                {
                    for (int j = 0; j < 2; j++)
                    {
                        Assert.True(tmpValue[i+1, j+1].ToString() == setValue[i,j]);
                    }
                }
                xls.Close();
            }
        }

        ///<summary>
        /// Rangeテストのシートインデックス用
        ///【テスト】GetRangeValue,SetRangeValue
        ///</summary>
        [Test]
        public void RangeValueIndex()
        {
            string[,] setValue = { { "A1", "B1" }, { "A2", "B2" }, { "A3", "B3" } };
            object[,] tmpValue;
            using (ExcelWrapper xls = new ExcelWrapper())
            {
                xls.AddBook();
                xls.SetRangeValue(1, "A1:B3", setValue);
                tmpValue = (object[,])xls.GetRangeValue(1, "A1:B3");
                Assert.True(tmpValue.Length == setValue.Length);
                for (int i = 0; i < 3; i++)
                {
                    for (int j = 0; j < 2; j++)
                    {
                        Assert.True(tmpValue[i + 1, j + 1].ToString() == setValue[i, j]);
                    }
                }
                xls.Close();
            }
        }

        ///<summary>
        /// 最終行と最終列の取得テスト(シート名称とシートインデックス）
        ///【テスト】GetLastRowIndex,GetLastColIndex
        ///</summary>
        [Test]
        public void GetLastRowColIndex()
        {
            string[,] setValue = { { "A1", "B1" }, { "A2", "B2" }, { "A3", "B3" } };
            using (ExcelWrapper xls = new ExcelWrapper())
            {
                xls.AddBook();
                xls.SetRangeValue(1, "A1:B3", setValue);
                Assert.True(xls.GetLastRowIndex("Sheet1") == 3);
                Assert.True(xls.GetLastColIndex("Sheet1") == 2);
                Assert.True(xls.GetLastRowIndex(1) == 3);
                Assert.True(xls.GetLastColIndex(1) == 2);
                xls.Close();
            }
        }
    }
}
