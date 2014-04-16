using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;

namespace MTC.MTCLibrary
{
    ///<summary>
    ///ExcelCOMオブジェクトを使用してExcelを操作するクラス
    ///参照設定は必要がない為、バージョンに依存しない
    ///</summary>
    public class ExcelWrapper : IDisposable
    {

        private const int PARAM_NUM_1 = 1;   // パラメータ数(1)
        private const int PARAM_NUM_2 = 2;   // パラメータ数(2)
        private const int PARAM_NUM_3 = 3;   // パラメータ数(3)
        private const int PARAM_NUM_4 = 4;   // パラメータ数(4)
        private const int PARAM_NUM_5 = 5;   // パラメータ数(5)
        private const int PARAM_NUM_6 = 6;   // パラメータ数(6)
        private const int PARAM_NUM_7 = 7;   // パラメータ数(7)
        private const int PARAM_NUM_8 = 8;   // パラメータ数(8)
        private const int PARAM_NUM_9 = 9;   // パラメータ数(9)
        private const int PARAM_NUM_10 = 10; // パラメータ数(10)
        private const int PARAM_NUM_11 = 11; // パラメータ数(11)
        private const int PARAM_NUM_12 = 12; // パラメータ数(12)
        private const int PARAM_NUM_13 = 13; // パラメータ数(13)
        private const int PARAM_NUM_14 = 14; // パラメータ数(14)
        private const int PARAM_NUM_15 = 15; // パラメータ数(15)
        private const int PARAM_NUM_16 = 16; // パラメータ数(16)
        private const int PARAM_NUM_17 = 17; // パラメータ数(17)
        private const int PARAM_NUM_18 = 18; // パラメータ数(18)
        private const int PARAM_NUM_19 = 19; // パラメータ数(19)
        private const int PARAM_NUM_20 = 20; // パラメータ数(20)
        private const int PARAM_NUM_21 = 21; // パラメータ数(21)
        private const int PARAM_NUM_22 = 22; // パラメータ数(22)
        private const int PARAM_NUM_23 = 23; // パラメータ数(23)
        private const int PARAM_NUM_24 = 24; // パラメータ数(24)
        private const int PARAM_NUM_25 = 25; // パラメータ数(25)
        private const int PARAM_NUM_26 = 26; // パラメータ数(26)
        private const int PARAM_NUM_27 = 27; // パラメータ数(27)
        private const int PARAM_NUM_28 = 28; // パラメータ数(28)
        private const int PARAM_NUM_29 = 29; // パラメータ数(29)
        private const int PARAM_NUM_30 = 30; // パラメータ数(30)
        private const int PARAM_NUM_31 = 31; // パラメータ数(31)
        private const int PARAM_NUM_32 = 32; // パラメータ数(32)
        private const int PARAM_NUM_33 = 33; // パラメータ数(33)
        private const int PARAM_NUM_34 = 34; // パラメータ数(34)
        private const int PARAM_NUM_35 = 35; // パラメータ数(35)
        private const int PARAM_NUM_36 = 36; // パラメータ数(36)

        ///<summary>Excelアプリケーションオブジェクト</summary>
        private object xlsApplication = null;
        ///<summary>Workbooksオブジェクト</summary>
        private object xlsBooks = null;
        ///<summary>Workbookオブジェクト</summary>
        private object xlsBook = null;
        ///<summary>Workbookオブジェクト</summary>
        private object xlsSheets = null;

        ///<summary>ExcelのCOMオブジェクトを参照できます。</summary>
        ///<value>getのみ使用可能で、Object型を返す</value>
        protected object XlsApplication
        {
            get
            {
                // 存在しない場合は作成する
                if (xlsApplication == null)
                {
                    Type classType = Type.GetTypeFromProgID("Excel.Application");
                    xlsApplication = Activator.CreateInstance(classType);
                }
                return xlsApplication;
            }
        }

        ///<summary>WorkBooksオブジェクトを参照できます。</summary>
        ///<value>getのみ使用可能で、Objetct型を返す</value>
        protected object XlsBooks
        {
            get
            {
                if (xlsBooks == null)
                {
                    xlsBooks = XlsApplication.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, XlsApplication, null);
                }
                return xlsBooks;
            }
        }

        ///<summary>
        ///インストールされているExcelのバージョンを参照できます。
        ///バージョンは2003(11.0),2007(12.0)，2010(14.0)
        ///</summary>
        ///<value>getのみ使用可能で、String型を返す</value>
        public string Version
        {
            get
            {
                object versionObj = XlsApplication.GetType().InvokeMember("Version", BindingFlags.GetProperty, null, XlsApplication, null);
                return versionObj.ToString();
            }
        }

        ///<summary>
        ///Excelの新規作成時に作成するシートを設定出来ます。
        ///</summary>
        ///<value>setのみ使用可能で、int型を渡す</value>
        public int SheetsInNewWorkBook
        {
            set
            {
                object[] parameters = new object[PARAM_NUM_1];
                parameters[0] = value;
                XlsApplication.GetType().InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, XlsApplication, parameters);
            }
        }
        ///<summary>
        ///Excel保存時の確認メッセージを有効及び無効に設定できます。
        ///</summary>
        ///<value>setのみ使用可能で、Bool型を渡す</value>
        public bool DisplayAlerts
        {
            set
            {
                object[] parameters = new object[PARAM_NUM_1];
                parameters[0] = value;
                XlsApplication.GetType().InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, XlsApplication, parameters);
            }
        }

        ///<summary>
        ///Excelアプリケーションの表示及び非表示を設定出来ます。
        ///</summary>
        ///<value>setのみ使用可能で、Bool型を渡す</value>
        public bool Visible
        {
            set
            {
                object[] parameters = new object[PARAM_NUM_1];
                parameters[0] = value;
                XlsApplication.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, XlsApplication, parameters);
            }
        }

        
        
        ///<summary>
        ///ExcelCOMオブジェクトのリリース
        ///使用したExcelCOMオブジェクトを解放する際に使用する
        ///</summary>
        ///<param name="target">解放するオブジェクト</param>
        private static void ReleaseComObject(object target)
        {
            try
            {
                if ((target != null))
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(target);
                }
            }
            finally
            {
                target = null;
            }
        }

        ///<summary>
        ///ExcelBookオブジェクトを取得する。
        ///</summary>
        ///<param name="index">インデックス</param>
        ///<returns>ExcelBookオブジェクトを渡す。</returns>
        private object GetBook(int index)
        {
            object[] parameters = new object[PARAM_NUM_1];
            parameters[0] = index;
            return XlsBooks.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, XlsBooks, parameters);
        }

        ///<summary>
        ///ExcelSheetsオブジェクトを取得する。
        ///</summary>
        ///<param name="book">ExcelBookオブジェクト</param>
        ///<returns>ExcelSheetsオブジェクトを渡す。</returns>
        private object GetSheets(object book)
        {
            return book.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, book, null);
        }

        ///<summary>
        ///ExcelSheetオブジェクトを取得する。
        ///</summary>
        ///<param name="sheets">ExcelSheetsオブジェクト</param>
        ///<param name="sheetName">シート名称</param>
        ///<returns>ExcelSheetオブジェクトを渡す。</returns>
        private object GetSheet(object sheets, string sheetName)
        {
            object[] parameters = new object[PARAM_NUM_1];
            parameters[0] = sheetName;
            return sheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, sheets, parameters);
        }

        ///<summary>
        ///ExcelSheetオブジェクトを取得する。
        ///</summary>
        ///<param name="sheets">ExcelSheetsオブジェクト</param>
        ///<param name="index">インデックス</param>
        ///<returns>ExcelSheetオブジェクトを渡す。</returns>
        private object GetSheet(object sheets, int index)
        {
            object[] parameters = new object[PARAM_NUM_1];
            parameters[0] = index;
            return sheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, sheets, parameters); ;
        }

        ///<summary>
        ///ExcelRangeオブジェクトを取得する。
        ///</summary>
        ///<param name="sheet">Excelオブジェクト</param>
        ///<param name="range">レンジオブジェクト</param>
        ///<returns>ExcelRangeオブジェクトを渡す。</returns>
        private object GetRange(object sheet, string range)
        {
            object[] parameters = new Object[PARAM_NUM_2];
            parameters[0] = range;
            parameters[1] = Type.Missing;
            return sheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, sheet, parameters);
        }

        ///<summary>
        ///指定されたレンジ範囲に値を設定する。
        ///</summary>
        ///<param name="range">Rangeオブジェクト</param>
        ///<param name="value">設定する値</param>
        private void SetRangeValue(object range, object value)
        {
            object[] parameters = new Object[PARAM_NUM_1];
            parameters[0] = value;
            range.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, range, parameters);
        }

        ///<summary>
        ///Excelのセルに設定されている値を2次元配列で取得する。
        ///</summary>
        ///<param name="range">レンジオブジェクト</param>
        ///<returns>指定範囲のセルの値を2次元配列で渡す。</returns>
        private object GetRangeValue(object range)
        {
            return range.GetType().InvokeMember("Value", BindingFlags.GetProperty, null, range, null);
        }


        ///<summary>
        ///ExcelCellsオブジェクトを取得する。
        ///</summary>
        ///<param name="sheet">Sheetオブジェクト</param>
        ///<returns>ExcelCellsオブジェクトで渡す。</returns>
        private object GetCells(object sheet)
        {
            return sheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, sheet, null);
        }

        ///<summary>
        ///ExcelCellオブジェクトを取得する。
        ///</summary>
        ///<param name="cells">Cellsオブジェクト</param>
        ///<param name="row">行</param>
        ///<param name="column">列</param>
        ///<returns>ExcelCellオブジェクトで渡す。</returns>
        private object GetCell(object cells, int row, int column)
        {
            object[] parameters = new Object[PARAM_NUM_2];
            parameters[0] = row;
            parameters[1] = column;
            return cells.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, cells, parameters);
        }

        ///<summary>シートの最終行又は最終列を取得する。</summary>
        ///<param name="cells">Cellsオブジェクト</param>
        ///<param name="property">Column又はRowを指定</param>
        ///<returns>最終行又は最終列を返す。</returns>
        private int GetCellsLast(object cells, string property)
        {
            object specialCells = null;
            object[] parameters = new object[PARAM_NUM_2];
            try
            {
                parameters[0] = 11; //xlCellTypeLastCell
                parameters[1] = Type.Missing;
                specialCells = cells.GetType().InvokeMember("SpecialCells", BindingFlags.InvokeMethod, null, cells, parameters);
                return (int)specialCells.GetType().InvokeMember(property, BindingFlags.GetProperty, null, specialCells, null);
            }finally
            {
                ReleaseComObject(specialCells);
            }
        }

        ///<summary>ExcelCOMオブジェクトの破棄を実施する</summary>
        public void Dispose()
        {
            XlsApplication.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, XlsApplication, null);
            ReleaseComObject(xlsSheets);
            xlsSheets = null;
            ReleaseComObject(xlsBook);
            xlsBook = null;
            ReleaseComObject(xlsBooks);
            xlsBooks = null;
            ReleaseComObject(xlsApplication);
            xlsApplication = null;
        }

        ///<summary>
        ///ExcelファイルのBookを開きます。
        ///</summary>
        ///<param name="xlsFilePath">Excelファイルパス</param>
        public void Open(string xlsFilePath)
        {
            object[] parameters = new object[PARAM_NUM_15];
            parameters[0] = xlsFilePath;
            parameters[1] = Type.Missing;
            parameters[2] = Type.Missing;
            parameters[3] = Type.Missing;
            parameters[4] = Type.Missing;
            parameters[5] = Type.Missing;
            parameters[6] = Type.Missing;
            parameters[7] = Type.Missing;
            parameters[8] = Type.Missing;
            parameters[9] = Type.Missing;
            parameters[10] = Type.Missing;
            parameters[11] = Type.Missing;
            parameters[12] = Type.Missing;
            parameters[13] = Type.Missing;
            parameters[14] = Type.Missing;
            XlsBooks.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, XlsBooks, parameters);
            xlsBook = GetBook(1);
            xlsSheets = GetSheets(xlsBook);
        }

        ///<summary>
        ///Excelファイル新規作成してBookを開きます。
        ///</summary>
        public void AddBook()
        {
            object[] parameters = new object[PARAM_NUM_1];
            parameters[0] = Type.Missing;
            XlsBooks.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, XlsBooks, parameters);
            xlsBook = GetBook(1);
            xlsSheets = GetSheets(xlsBook);

        }

        ///<summary>
        ///ExcelファイルのBookを閉じます。
        ///</summary>
        public void Close(bool saveChanges, string filename)
        {
            object[] parameters = new object[PARAM_NUM_2];
            parameters[0] = saveChanges;
            parameters[1] = filename;
            xlsBook.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, xlsBook, parameters);
        }

        ///<summary>
        ///ExcelファイルのBookを保存します。
        ///</summary>
        public void Save()
        {
            xlsBook.GetType().InvokeMember("Save", BindingFlags.InvokeMethod, null, xlsBook, null);
        }

        ///<summary>
        ///ExcelファイルのBookを保存します。
        ///</summary>
        ///<param name="xlsFilePath">Excelファイルパス</param>
        public void SaveAs(string xlsFilePath)
        {
            object[] parameters = new object[PARAM_NUM_12];
            parameters[0] = xlsFilePath;
            parameters[1] = Type.Missing;
            parameters[2] = Type.Missing;
            parameters[3] = Type.Missing;
            parameters[4] = Type.Missing;
            parameters[5] = Type.Missing;
            parameters[6] = Type.Missing;
            parameters[7] = Type.Missing;
            parameters[8] = Type.Missing;
            parameters[9] = Type.Missing;
            parameters[10] = Type.Missing;
            parameters[11] = Type.Missing;
            xlsBook.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, xlsBook, parameters);           
        }

        ///<summary>
        ///指定したレンジ範囲のデータを一括で取得する。
        ///</summary>
        ///<param name="sheetName">シート名称</param>
        ///<param name="rangeMap">レンジ範囲</param>
        ///<returns>データを2次元配列で渡す</returns>
        public object GetRangeValue(string sheetName, string rangeMap)
        {
            object sheet = null;
            object range = null;

            try
            {
                sheet = GetSheet(xlsSheets, sheetName);
                range = GetRange(sheet, rangeMap);
                return GetRangeValue(range);;
            }finally
            {
                ReleaseComObject(range);
                ReleaseComObject(sheet);
            }

        }

        ///<summary>
        ///指定したレンジ範囲のデータを一括で取得する。
        ///</summary>
        ///<param name="sheetIndex">シートインデックス</param>
        ///<param name="rangeMap">レンジ範囲</param>
        ///<returns>データを2次元配列で渡す</returns>
        public object GetRangeValue(int sheetIndex, string rangeMap)
        {
            object sheet = null;
            object range = null;

            try
            {
                sheet = GetSheet(xlsSheets, sheetIndex);
                range = GetRange(sheet, rangeMap);
                return GetRangeValue(range);
            }
            finally
            {
                ReleaseComObject(range);
                ReleaseComObject(sheet);
            }

        }

        ///<summary>
        ///指定したレンジ範囲のデータを一括で設定する。
        ///※配列の開始位置は'1'からなので注意して下さい。
        ///</summary>
        ///<param name="sheetName">シート名称</param>
        ///<param name="rangeMap">レンジ範囲</param>
        ///<param name="value">設定する値</param>
        public void SetRangeValue(string sheetName, string rangeMap,object value)
        {
            object sheet = null;
            object range = null;

            try
            {
                sheet = GetSheet(xlsSheets, sheetName);
                range = GetRange(sheet, rangeMap);
                SetRangeValue(range, value);
            }
            finally
            {
                ReleaseComObject(range);
                ReleaseComObject(sheet);
            }

        }

        ///<summary>
        ///指定したレンジ範囲のデータを一括で設定する。
        ///※配列の開始位置は'1'からなので注意して下さい。
        ///</summary>
        ///<param name="sheetIndex">シート番号</param>
        ///<param name="rangeMap">レンジ範囲</param>
        ///<param name="value">設定する値</param>
        public void SetRangeValue(int sheetIndex, string rangeMap, object value)
        {
            object sheet = null;
            object range = null;

            try
            {
                sheet = GetSheet(xlsSheets, sheetIndex);
                range = GetRange(sheet, rangeMap);
                SetRangeValue(range, value);
            }
            finally
            {
                ReleaseComObject(range);
                ReleaseComObject(sheet);
            }

        }


        ///<summary>
        ///指定したシートの最終行を取得する。
        ///</summary>
        ///<param name="sheetName">シート名</param>
        public int GetLastRowIndex( string sheetName)
        {
            object sheet = null;
            object cells = null;

            try
            {
                sheet = GetSheet(xlsSheets, sheetName);
                cells = GetCells(sheet);
                return GetCellsLast(cells, "Row");

            }finally
            {
                ReleaseComObject(cells);
                ReleaseComObject(sheet);
            }
        }

        ///<summary>
        ///指定したシートの最終行を取得する。
        ///</summary>
        ///<param name="sheetIndex">シートインデックス</param>
        public int GetLastRowIndex(int sheetIndex)
        {
            object sheet = null;
            object cells = null;
         
            try
            {
                sheet = GetSheet(xlsSheets, sheetIndex);
                cells = GetCells(sheet);
                return GetCellsLast(cells, "Row");
            }
            finally
            {
                ReleaseComObject(cells);
                ReleaseComObject(sheet);
            }
        }

        ///<summary>
        ///指定したシートの最終列を取得する。
        ///</summary>
        ///<param name="sheetName">シート名</param>
        public int GetLastColIndex(string sheetName)
        {
            object sheet = null;
            object cells = null;

            try
            {
                sheet = GetSheet(xlsSheets, sheetName);
                cells = GetCells(sheet);
                return GetCellsLast(cells, "Column");
            }
            finally
            {
                ReleaseComObject(cells);
                ReleaseComObject(sheet);
            }
        }

        ///<summary>
        ///指定したシートの最終列を取得する。
        ///</summary>
        ///<param name="sheetIndex">シートインデックス</param>
        public int GetLastColIndex(int sheetIndex)
        {
            object sheet = null;
            object cells = null;

            try
            {
                sheet = GetSheet(xlsSheets, sheetIndex);
                cells = GetCells(sheet);
                return GetCellsLast(cells, "Column");
            }
            finally
            {
                ReleaseComObject(cells);
                ReleaseComObject(sheet);
            }
        }


    }
}
