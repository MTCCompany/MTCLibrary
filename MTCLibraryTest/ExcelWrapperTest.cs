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
        Process[] psExcel = null;
        
        [TestFixtureSetUp]
        public void init()
        {
            CreateExcel("\\ExcelWrapperTest001.xls");
            CreateExcel("\\ExcelWrapperTest002.xls");
            CreateExcel("\\ExcelWrapperTest003.xls");
            CreateExcel("\\ExcelWrapperTest004.xls");
            psExcel = Process.GetProcessesByName("Excel");
        }

        [TestFixtureTearDown]
        public void Cleanup()
        {
            //作成したExcelプロセスが残っていないことを確認
            Assert.True(psExcel.Length >= Process.GetProcessesByName("Excel").Length);
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
        ///【テスト】Close
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
        /// 最終行と最終列の取得テスト(シート名称とシート番号）
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

        ///<summary>
        /// 指定したセルの設定/取得のテスト(シート名称）
        ///【テスト】SetCellValue,GetCellValue
        ///</summary>
        [Test]
        public void GetCellValue_Name()
        {
          
            using (ExcelWrapper xls = new ExcelWrapper())
            {
                xls.AddBook();
                xls.SetCellValue("Sheet1", 1, 1, "A1");
                Assert.True(xls.GetCellValue("Sheet1",1,1) == "A1");
                xls.Close();
            }
        }

        ///<summary>
        /// 指定したセルの設定/取得のテスト(シート番号）
        ///【テスト】SetCellValue,GetCellValue
        ///</summary>
        [Test]
        public void GetCellValue_Index()
        {

            using (ExcelWrapper xls = new ExcelWrapper())
            {
                xls.AddBook();
                xls.SetCellValue(1, 1, 1, "A1");
                Assert.True(xls.GetCellValue(1, 1, 1) == "A1");
                xls.Close();
            }
        }

        ///<summary>
        /// シートの追加とシート数の確認
        ///【テスト】AddBook,GetSheetCount
        ///</summary>
        [Test]
        public void AddBookGetSheetCount()
        {

            using (ExcelWrapper xls = new ExcelWrapper())
            {
                xls.AddBook();
                int count = xls.GetSheetCount();
                xls.AddSheet(2);
                Assert.True(xls.GetSheetCount()-count == 2);
                xls.Close();
            }
        }

        ///<summary>
        /// 指定したセルを削除と名前変更（シート名）
        ///【テスト】DeleteSheet,SetNameSheet
        ///</summary>
        [Test]
        public void DeleteReNameSheet_Name()
        {

            using (ExcelWrapper xls = new ExcelWrapper())
            {
                xls.AddBook();
                xls.AddSheet(2);
                xls.DeleteSheet("Sheet1");
                xls.SetSheetName("Sheet2", "Sheet1");
                xls.Close();
            }
        }

        ///<summary>
        /// 指定したセルを削除と名前変更（シート番号）
        ///【テスト】DeleteSheet,SetNameSheet
        ///</summary>
        [Test]
        public void DeleteReNameSheet_Index()
        {

            using (ExcelWrapper xls = new ExcelWrapper())
            {
                xls.AddBook();
                xls.AddSheet(2);
                xls.DeleteSheet(1);
                xls.SetSheetName(2, "Sheet1");
                xls.DeleteSheet("Sheet1");
                xls.Close();
            }
        }

        ///<summary>
        /// 指定したセルをコピーする。
        ///【テスト】CopySheet
        ///</summary>
        [Test]
        public void CopySheet_Name()
        {
            　
            using (ExcelWrapper xls = new ExcelWrapper())
            {
                xls.AddBook();
                xls.SetCellValue("Sheet1", 1, 1, "A1");
                xls.CopySheet("Sheet1", "test");
                Assert.True(xls.GetCellValue("test", 1, 1) == "A1");
                xls.Close();
            }
        }

        ///<summary>
        /// 罫線の設定（シート名&シート番号）
        ///【テスト】SetRangeLine
        ///</summary>
        [Test]
        public void SetRangeLine()
        {

            using (ExcelWrapper xls = new ExcelWrapper())
            {
                xls.Open(tempPath + "\\ExcelWrapperTest002.xls");
                xls.AddSheet(1);

                xls.SetRangeLine("Sheet1", "B2:C4", ExcelWrapper.XlBordersIndex.xlEdgeTop);
                xls.SetRangeLine("Sheet1", "B2:C4", ExcelWrapper.XlBordersIndex.xlEdgeLeft);
                xls.SetRangeLine("Sheet1", "B2:C4", ExcelWrapper.XlBordersIndex.xlEdgeRight);
                xls.SetRangeLine("Sheet1", "B2:C4", ExcelWrapper.XlBordersIndex.xlEdgeBottom);
                xls.SetRangeLine("Sheet1", "B2:C4", ExcelWrapper.XlBordersIndex.xlInsideHorizontal);
                xls.SetRangeLine("Sheet1", "B2:C4", ExcelWrapper.XlBordersIndex.xlInsideVertical);
                xls.SetRangeLine("Sheet1", "B2:C4", ExcelWrapper.XlBordersIndex.xlDiagonalDown);
                xls.SetRangeLine("Sheet1", "B2:C4", ExcelWrapper.XlBordersIndex.xlDiagonalUp);

                xls.SetRangeLine("Sheet1", "E2", ExcelWrapper.XlBordersIndex.xlEdgeTop, ExcelWrapper.XlLineStyle.xlContinuous);
                xls.SetRangeLine("Sheet1", "E3", ExcelWrapper.XlBordersIndex.xlEdgeTop, ExcelWrapper.XlLineStyle.xlDash);
                xls.SetRangeLine("Sheet1", "E4", ExcelWrapper.XlBordersIndex.xlEdgeTop, ExcelWrapper.XlLineStyle.xlDashDot);
                xls.SetRangeLine("Sheet1", "E5", ExcelWrapper.XlBordersIndex.xlEdgeTop, ExcelWrapper.XlLineStyle.xlDashDotDot);
                xls.SetRangeLine("Sheet1", "E6", ExcelWrapper.XlBordersIndex.xlEdgeTop, ExcelWrapper.XlLineStyle.xlDot);
                // 下３つはうまく動作しない
                xls.SetRangeLine("Sheet1", "E7", ExcelWrapper.XlBordersIndex.xlEdgeTop, ExcelWrapper.XlLineStyle.xlDouble);
                xls.SetRangeLine("Sheet1", "E8", ExcelWrapper.XlBordersIndex.xlEdgeTop, ExcelWrapper.XlLineStyle.xlNone);
                xls.SetRangeLine("Sheet1", "E9", ExcelWrapper.XlBordersIndex.xlEdgeTop, ExcelWrapper.XlLineStyle.xlSlantDashDot);

                xls.SetRangeLine("Sheet1", "G2", ExcelWrapper.XlBordersIndex.xlEdgeTop,
                    ExcelWrapper.XlLineStyle.xlContinuous, ExcelWrapper.XlBorderWeight.xlHairline);
                xls.SetRangeLine("Sheet1", "G3", ExcelWrapper.XlBordersIndex.xlEdgeTop,
                    ExcelWrapper.XlLineStyle.xlContinuous, ExcelWrapper.XlBorderWeight.xlMedium);
                xls.SetRangeLine("Sheet1", "G4", ExcelWrapper.XlBordersIndex.xlEdgeTop,
                    ExcelWrapper.XlLineStyle.xlContinuous, ExcelWrapper.XlBorderWeight.xlThick);
                xls.SetRangeLine("Sheet1", "G5", ExcelWrapper.XlBordersIndex.xlEdgeTop,
                    ExcelWrapper.XlLineStyle.xlContinuous, ExcelWrapper.XlBorderWeight.xlThin);

                xls.SetRangeLine("Sheet1", "H2", ExcelWrapper.XlBordersIndex.xlEdgeTop,
                    ExcelWrapper.XlLineStyle.xlContinuous, ExcelWrapper.XlBorderWeight.xlThin,2);
                xls.SetRangeLine("Sheet1", "H3", ExcelWrapper.XlBordersIndex.xlEdgeTop,
                    ExcelWrapper.XlLineStyle.xlContinuous, ExcelWrapper.XlBorderWeight.xlThin, 3);

                xls.SetRangeLine(2, "B2:C4", ExcelWrapper.XlBordersIndex.xlEdgeTop);
                xls.SetRangeLine(2, "B2:C4", ExcelWrapper.XlBordersIndex.xlEdgeLeft);
                xls.SetRangeLine(2, "B2:C4", ExcelWrapper.XlBordersIndex.xlEdgeRight);
                xls.SetRangeLine(2, "B2:C4", ExcelWrapper.XlBordersIndex.xlEdgeBottom);
                xls.SetRangeLine(2, "B2:C4", ExcelWrapper.XlBordersIndex.xlInsideHorizontal);
                xls.SetRangeLine(2, "B2:C4", ExcelWrapper.XlBordersIndex.xlInsideVertical);
                xls.SetRangeLine(2, "B2:C4", ExcelWrapper.XlBordersIndex.xlDiagonalDown);
                xls.SetRangeLine(2, "B2:C4", ExcelWrapper.XlBordersIndex.xlDiagonalUp);

                xls.SetRangeLine(2, "E2", ExcelWrapper.XlBordersIndex.xlEdgeTop, ExcelWrapper.XlLineStyle.xlContinuous);
                xls.SetRangeLine(2, "E3", ExcelWrapper.XlBordersIndex.xlEdgeTop, ExcelWrapper.XlLineStyle.xlDash);
                xls.SetRangeLine(2, "E4", ExcelWrapper.XlBordersIndex.xlEdgeTop, ExcelWrapper.XlLineStyle.xlDashDot);
                xls.SetRangeLine(2, "E5", ExcelWrapper.XlBordersIndex.xlEdgeTop, ExcelWrapper.XlLineStyle.xlDashDotDot);
                xls.SetRangeLine(2, "E6", ExcelWrapper.XlBordersIndex.xlEdgeTop, ExcelWrapper.XlLineStyle.xlDot);
                // 下３つはうまく動作しない
                xls.SetRangeLine(2, "E7", ExcelWrapper.XlBordersIndex.xlEdgeTop, ExcelWrapper.XlLineStyle.xlDouble);
                xls.SetRangeLine(2, "E8", ExcelWrapper.XlBordersIndex.xlEdgeTop, ExcelWrapper.XlLineStyle.xlNone);
                xls.SetRangeLine(2, "E9", ExcelWrapper.XlBordersIndex.xlEdgeTop, ExcelWrapper.XlLineStyle.xlSlantDashDot);

                xls.SetRangeLine(2, "G2", ExcelWrapper.XlBordersIndex.xlEdgeTop,
                    ExcelWrapper.XlLineStyle.xlContinuous, ExcelWrapper.XlBorderWeight.xlHairline);
                xls.SetRangeLine(2, "G3", ExcelWrapper.XlBordersIndex.xlEdgeTop,
                    ExcelWrapper.XlLineStyle.xlContinuous, ExcelWrapper.XlBorderWeight.xlMedium);
                xls.SetRangeLine(2, "G4", ExcelWrapper.XlBordersIndex.xlEdgeTop,
                    ExcelWrapper.XlLineStyle.xlContinuous, ExcelWrapper.XlBorderWeight.xlThick);
                xls.SetRangeLine(2, "G5", ExcelWrapper.XlBordersIndex.xlEdgeTop,
                    ExcelWrapper.XlLineStyle.xlContinuous, ExcelWrapper.XlBorderWeight.xlThin);

                xls.SetRangeLine(2, "H2", ExcelWrapper.XlBordersIndex.xlEdgeTop,
                    ExcelWrapper.XlLineStyle.xlContinuous, ExcelWrapper.XlBorderWeight.xlThin, 2);
                xls.SetRangeLine(2, "H3", ExcelWrapper.XlBordersIndex.xlEdgeTop,
                    ExcelWrapper.XlLineStyle.xlContinuous, ExcelWrapper.XlBorderWeight.xlThin, 3);

                xls.Save();
                xls.Close();
            }
        }

        ///<summary>
        /// フォント色、背景色、アクティブシートの設定テスト（シート名）
        ///【テスト】SetRangeFontColor,SetRangePatternColor,SetActivateSheet
        ///</summary>
        [Test]
        public void SetRangeFontColor_Name()
        {
            using (ExcelWrapper xls = new ExcelWrapper())
            {
                xls.Open(tempPath + "\\ExcelWrapperTest003.xls");
                xls.AddSheet(1);
                xls.SetRangeValue("Sheet1", "A1", "あ");
                xls.SetRangeFontColor("Sheet1", "A1", 3);
                xls.SetRangePatternColor("Sheet1", "B1:B3", 3);
                xls.Save();
                xls.Close();
            }
        }

        ///<summary>
        /// フォント色、背景色、アクティブシートの設定テスト（シート番号）
        ///【テスト】SetRangeFontColor,SetRangePatternColor,SetActivateSheet
        ///</summary>
        [Test]
        public void SetRangeFontColor_Index()
        {
            using (ExcelWrapper xls = new ExcelWrapper())
            {
                xls.Open(tempPath + "\\ExcelWrapperTest004.xls");
                xls.AddSheet(1);
                xls.SetRangeValue(2, "A1", "あ");
                xls.SetRangeFontColor(2, "A1", 3);
                xls.SetRangePatternColor(2, "B1:B3", 3);
                xls.SetActivateSheet(2);
                xls.Save();
                xls.Close();
            }
        }
    }
}
