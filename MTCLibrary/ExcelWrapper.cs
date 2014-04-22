using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;

namespace MTC.MTCLibrary
{
    ///<summary>
    ///ExcelCOMオブジェクトを使用してExcelを操作するクラスです。
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

        ///<summary>罫線の位置を指定する定数</summary>
        public enum XlBordersIndex : int
        {
            ///<summary>セルに右下がりの斜線 (対角線)</summary>
            xlDiagonalDown = 5,
            ///<summary>セルに右上がりの斜線 (対角線)</summary>
            xlDiagonalUp      = 6,
            ///<summary>セルの下の線</summary>
            xlEdgeBottom      = 9,
            ///<summary>セルの左の線</summary>
            xlEdgeLeft        = 7,
            ///<summary>セルの右の線</summary>
            xlEdgeRight       =10,
            ///<summary>セルの上の線</summary>
            xlEdgeTop         = 8,
            ///<summary>セセル範囲の内側の横線</summary>
            xlInsideHorizontal=12,
            ///<summary>セル範囲の内側の縦線</summary>
            xlInsideVertical  =11
        }
        ///<summary>罫線の種類を指定する定数</summary>
        public enum XlLineStyle : int
        {
            ///<summary>実線 (初期値です)</summary>
            xlContinuous   = 1,
            ///<summary>破線</summary>
            xlDash	       =-4115,
            ///<summary>一点鎖線</summary>
            xlDashDot      = 4,
            ///<summary>二点鎖線</summary>
            xlDashDotDot   = 5,
            ///<summary>点線</summary>
            xlDot          =-4118,
            ///<summary>二重線</summary>
            xlDouble       =-4119,
            ///<summary>斜め一点鎖線</summary>
            xlSlantDashDot =13,
            ///<summary>線なし</summary>
            xlNone = -4142
        }
        ///<summary>罫線の太さを指定する定数</summary>
        public enum XlBorderWeight : int
        {
            ///<summary>極細</summary>
            xlHairline = 1,
            ///<summary>細 (初期値です)</summary>
            xlThin     = 2,
            ///<summary>太</summary>
            xlMedium   =-4138,
            ///<summary>極太</summary>
            xlThick    = 4
        }

        //罫線の位置を指定する定数
        ///<summary>セルに右下がりの罫線</summary>
        public const int XLCONTINUOS = 1;

        ///<summary>ExcelのCOMオブジェクトを参照できます。</summary>
        ///<value>getのみ使用可能で、Object型を返す</value>
        private object XlsApplication
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
        private object XlsBooks
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
        ///ExcelBookオブジェクト及びExcelSheetsオブジェクトを設定する
        ///</summary>
        ///<param name="index">インデックス</param>
        private void SetBookSheets(int index)
        {
            object[] parameters = new object[PARAM_NUM_1];
            parameters[0] = index;
            xlsBook = XlsBooks.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, XlsBooks, parameters);
            xlsSheets = xlsBook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, xlsBook, null);
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
        ///ExcelBordersオブジェクトを取得する。
        ///</summary>
        ///<param name="type">ExcelRangeオブジェクト</param>
        ///<param name="range">レンジオブジェクト</param>
        ///<returns>ExcelRangeオブジェクトを渡す。</returns>
        private object GetBorders(object range, XlBordersIndex type)
        {
            object[] parameters = new Object[PARAM_NUM_1];
            parameters[0] = type;
            return range.GetType().InvokeMember("Borders", BindingFlags.GetProperty, null, range, parameters);
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
        ///<param name="row">行位置</param>
        ///<param name="column">列位置</param>
        ///<returns>ExcelCellオブジェクトで渡す。</returns>
        private object GetCell(object cells, int row, int column)
        {
            object[] parameters = new Object[PARAM_NUM_2];
            parameters[0] = row;
            parameters[1] = column;
            return cells.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, cells, parameters);
        }

        ///<summary>
        ///指定のセルから値を取得する。
        ///</summary>
        ///<param name="cell">Cellオブジェクト</param>
        ///<returns>セル値をオブジェクトで渡す。</returns>
        private object GetCellText(object cell)
        {
            return cell.GetType().InvokeMember("Text", BindingFlags.GetProperty, null, cell, null);
        }

        ///<summary>
        ///指定のセルから値を設定する。
        ///</summary>
        ///<param name="cell">Cellオブジェクト</param>
        ///<param name="value">設定する値の文字列</param>
        private void SetCellText(object cell, string value)
        {
            object[] parameters = new Object[PARAM_NUM_1];
            parameters[0] = value;
            cell.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, cell, parameters);
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
        ///既存のExcelファイルを開きます。
        ///</summary>
        ///<remarks>SetBookSheets(1)を呼び出している為、再度呼び出す必要はありません</remarks>
        ///<example>
        /// 次のコードでは、Excelファイル(@"C:\\Test.xls")を開きます。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        /// }
        /// </code>
        /// </example>
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
            SetBookSheets(1);
        }

        ///<summary>
        ///Excelファイル新規作成してBookを開きます。
        ///</summary>
        ///<remarks>SetBookSheets(1)を呼び出している為、再度呼び出す必要はありません</remarks>
        ///<example>
        /// 次のコードでは、新規Excelファイルを作成します。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.AddBook();
        ///     xls.Close();
        /// }
        /// </code>
        ///</example>
        public void AddBook()
        {
            object[] parameters = new object[PARAM_NUM_1];
            parameters[0] = Type.Missing;
            XlsBooks.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, XlsBooks, parameters);
            SetBookSheets(1);
        }

        ///<summary>
        ///ExcelファイルのBookを閉じます。
        ///</summary>
        /// 次のコードでは、保存せずに新規作成したExcelファイルを閉じます
        /// <example>
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.AddBook();
        ///     xls.Close();
        /// }
        /// </code>
        /// </example>
        public void Close()
        {
            object[] parameters = new object[PARAM_NUM_2];
            parameters[0] = false;
            parameters[1] = null;
            xlsBook.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, xlsBook, parameters);
        }

        ///<summary>
        ///ExcelファイルのBookを保存します。
        ///</summary>
        ///<example>
        /// 次のコードでは、Excelファイルを保存します。
        /// 
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     xls.Save();
        /// }
        ///</example>
        public void Save()
        {
            xlsBook.GetType().InvokeMember("Save", BindingFlags.InvokeMethod, null, xlsBook, null);
        }

        ///<summary>
        ///ExcelファイルのBookを保存します。
        ///</summary>
        ///<example>
        /// 次のコードでは、ファイル名"C:\Test.xls"で新規作成したExcelファイルを閉じます。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.AddBook();
        ///     xls.SaveAs("C:\Test.xls");
        /// }
        /// </code>
        ///</example>
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
        ///<remarks>配列の開始位置は'1'からなので注意して下さい。</remarks>
        ///<param name="sheetName">シート名称</param>
        ///<param name="rangeMap">レンジ範囲</param>
        ///<returns>データを2次元配列で渡す</returns>
        ///<example>
        /// 次のコードでは、 シート名[Sheet1]のレンジ[A1:A2]からセル値を取得します。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     object[,] temp = GetRangeValue("Sheet1", "A1:B2");
        ///     //セル[A1]のデータを出力します。
        ///     System.Console.WriteLine(temp[1,1].toString());
        /// }
        /// </code>
        ///</example>
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
        ///<remarks>配列の開始位置は'1'からなので注意して下さい。</remarks>
        ///<param name="sheetIndex">シート番号</param>
        ///<param name="rangeMap">レンジ範囲</param>
        ///<returns>データを2次元配列で渡す</returns>
        ///<example>
        /// 次のコードでは、シート番号[1]のレンジ[A1:A2]からセル値を取得します。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     object[,] temp = GetRangeValue(1, "A1:B2");
        ///     //セル[A1]のデータを出力します。
        ///     System.Console.WriteLine(temp[1,1].toString());
        /// }
        /// </code>
        ///</example>
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
        ///</summary>
        ///<param name="sheetName">シート名称</param>
        ///<param name="rangeMap">レンジ範囲</param>
        ///<param name="value">設定する値</param>
        ///<example>
        /// 次のコードでは、シート名[Sheet1]のレンジ[A1:A2]にセル値を設定します。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     string[,] setValue = { { "A1", "B1" }, { "A2", "B2" } };
        ///     SetRangeValue("Sheet1","A1:B2", setValue);
        ///     xls.Save();
        /// }
        /// </code>
        ///</example>
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
        ///</summary>
        ///<param name="sheetIndex">シート番号</param>
        ///<param name="rangeMap">レンジ範囲</param>
        ///<param name="value">設定する値</param>
        ///<example>
        /// 次のコードでは、 シート番号[1]のレンジ[A1:A2"]にセル値を設定します。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     string[,] setValue = { { "A1", "B1" }, { "A2", "B2" } };
        ///     SetRangeValue(1,"A1:B2", setValue);
        ///     xls.Save();
        /// }
        /// </code>
        ///</example>
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
        ///<returns>シートの最終行を渡す</returns>
        ///<example>
        /// 次のコードでは、シート名[Sheet1]の最終行を取得します。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     System.Console.WriteLine(xls.GetLastRowIndex("Sheet1").toString());
        /// }
        /// </code>
        ///</example>
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
        ///<param name="sheetIndex">シート番号</param>
        ///<returns>シートの最終行を渡す</returns>
        ///<example>
        /// 次のコードでは、シート番号[1]の最終行を取得します。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     System.Console.WriteLine(xls.GetLastRowIndex(1).toString());
        /// }
        /// </code>
        ///</example>
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
        ///<returns>シートの最終列を渡す</returns>
        ///<example>
        /// 次のコードでは、シート名[Sheet1]の最終列を取得します。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     System.Console.WriteLine(xls.GetLastColIndex("Sheet1").toString());
        /// </code>
        /// }
        ///</example>
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
        ///<param name="sheetIndex">シート番号</param>
        ///<returns>シートの最終列を渡す</returns>
        ///<example>
        /// 次のコードでは、シート番号[1]の最終列を取得します。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     System.Console.WriteLine(xls.GetLastColIndex(1).toString());
        /// }
        /// </code>
        ///</example>
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

        ///<summary>
        ///指定したセルの値を文字列で取得する。
        ///</summary>
        ///<param name="sheetName">シート名</param>
        ///<param name="row">行位置</param>
        ///<param name="column">列位置</param>
        ///<returns>セルの値を文字列で渡す</returns>
        ///<example>
        /// 次のコードでは、シート名[Sheet1]のセル[行1,列1](A1)からセル値を取得します。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     System.Console.WriteLine(xls.GetCellValue("Sheet1",1,1));
        /// }
        /// </code>
        ///</example>
        public string GetCellValue(string sheetName, int row, int column)
        {
            object sheet = null;
            object cells = null;
            object cell = null;

            try
            {
                sheet = GetSheet(xlsSheets, sheetName);
                cells = GetCells(sheet);
                cell = GetCell(cells, row, column);
                return GetCellText(cell).ToString();
            }
            finally
            {
                ReleaseComObject(cell);
                ReleaseComObject(cells);
                ReleaseComObject(sheet);
            }
        }

        ///<summary>
        ///指定したセルの値を文字列で取得する。
        ///</summary>
        ///<param name="sheetIndex">シート番号</param>
        ///<param name="row">行位置</param>
        ///<param name="column">列位置</param>
        ///<returns>セルの値を文字列で渡す</returns>
        ///<example>
        /// 次のコードでは、シート番号[1]のセル[行1,列1](A1)からセル値を取得します。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     System.Console.WriteLine(xls.GetCellValue(1,1,1));
        /// }
        /// </code>
        ///</example>
        public string GetCellValue(int sheetIndex, int row, int column)
        {
            object sheet = null;
            object cells = null;
            object cell = null;

            try
            {
                sheet = GetSheet(xlsSheets, sheetIndex);
                cells = GetCells(sheet);
                cell = GetCell(cells, row, column);
                return GetCellText(cell).ToString();
            }
            finally
            {
                ReleaseComObject(cell);
                ReleaseComObject(cells);
                ReleaseComObject(sheet);
            }
        }

        ///<summary>
        ///指定したセルに文字列を設定する。
        ///</summary>
        ///<param name="sheetName">シート名</param>
        ///<param name="row">行位置</param>
        ///<param name="column">列位置</param>
        ///<param name="value">文字列</param>        
        ///<example>
        /// 次のコードでは、シート名[Sheet1]のセル[行1,列1](A1)に文字列[A1]を設定する。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     xls.SetCellValue("Sheet1",1,1,"A1");
        /// }
        /// </code>
        ///</example>
        public void SetCellValue(string sheetName, int row, int column,string value)
        {
            object sheet = null;
            object cells = null;
            object cell = null;

            try
            {
                sheet = GetSheet(xlsSheets, sheetName);
                cells = GetCells(sheet);
                cell = GetCell(cells, row, column);
                SetCellText(cell, value);
            }
            finally
            {
                ReleaseComObject(cell);
                ReleaseComObject(cells);
                ReleaseComObject(sheet);
            }
        }

        ///<summary>
        ///指定したセルに文字列を設定する。
        ///</summary>
        ///<param name="sheetIndex">シート番号</param>
        ///<param name="row">行位置</param>
        ///<param name="column">列位置</param>
        ///<param name="value">文字列</param>        
        ///<example>
        /// 次のコードでは、シート番号[1]のセル[行1,列1](A1)に文字列[A1]を設定する。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     xls.SetCellValue(1,1,1,"A1");
        /// }
        /// </code>
        ///</example>
        public void SetCellValue(int sheetIndex, int row, int column, string value)
        {
            object sheet = null;
            object cells = null;
            object cell = null;

            try
            {
                sheet = GetSheet(xlsSheets, sheetIndex);
                cells = GetCells(sheet);
                cell = GetCell(cells, row, column);
                SetCellText(cell, value);
            }
            finally
            {
                ReleaseComObject(cell);
                ReleaseComObject(cells);
                ReleaseComObject(sheet);
            }
        }

        ///<summary>
        ///Excelファイルにシートを一番後ろに追加する。
        ///</summary>
        ///<param name="count">追加するシート数</param>
        ///<example>
        /// 次のコードでは、新規Excelファイルに２シート追加します。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.AddBook();
        ///     xls.AddSheet(2);
        ///     xls.Close();
        /// }
        /// </code>
        ///</example>
        public void AddSheet(int count)
        {
            object rnt = null;
            object sheet = null;
            
            try
            {
                sheet = GetSheet(xlsSheets, GetSheetCount());
                object[] parameters = new object[PARAM_NUM_4];
                parameters[0] = Type.Missing;
                parameters[1] = sheet;
                parameters[2] = count;
                parameters[3] = Type.Missing;
                rnt = xlsSheets.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, xlsSheets, parameters);
            }
            finally
            {
                ReleaseComObject(sheet);
                ReleaseComObject(rnt);
            }
        }

        ///<summary>
        ///指定したシートを削除する
        ///</summary>
        ///<param name="sheetName">シート名</param>
        ///<example>
        /// 次のコードでは、シート名「Sheet1］のシートを削除する。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     xls.DisplayAlerts = true;
        ///     xls.DeleteSheet("Sheet1");
        ///     xls.Save();
        /// }
        /// </code>
        ///</example>
        public void DeleteSheet(string sheetName)
        {
            object sheet = null;

            try
            {
                sheet = GetSheet(xlsSheets, sheetName);
                sheet.GetType().InvokeMember("Delete", BindingFlags.InvokeMethod, null, sheet, null);
            }
            finally
            {
                ReleaseComObject(sheet);
            }
        }

        ///<summary>
        ///指定したシートを削除する
        ///</summary>
        ///<param name="sheetIndex">シート番号</param>
        ///<example>
        /// 次のコードでは、シート番号「1」のシートを削除する。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     xls.DisplayAlerts = true;
        ///     xls.DeleteSheet(1);
        ///     xls.Save();
        /// }
        /// </code>
        ///</example>
        public void DeleteSheet(int sheetIndex)
        {
            object sheet = null;

            try
            {
                sheet = GetSheet(xlsSheets, sheetIndex);
                sheet.GetType().InvokeMember("Delete", BindingFlags.InvokeMethod, null, sheet, null);
            }
            finally
            {
                ReleaseComObject(sheet);
            }
        }

        ///<summary>
        ///指定したシートの名前を変更する。
        ///</summary>
        ///<param name="sheetName">変更対象のシート名</param>
        ///<param name="reName">変更後のシート名</param>
        ///<example>
        /// 次のコードでは、シート名を「Sheet1］から「テスト」に変更する。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     xls.SetSheetName("Sheet1","テスト");
        /// }
        /// </code>
        ///</example>
        public void SetSheetName(string sheetName,string reName)
        {
            object sheet = null;

            try
            {
                sheet = GetSheet(xlsSheets, sheetName);
                object[] parameters = new object[PARAM_NUM_1];
                parameters[0] = reName;
                sheet.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, sheet, parameters);
            }
            finally
            {
                ReleaseComObject(sheet);
            }
        }

        ///<summary>
        ///指定したシートの名前を変更する。
        ///</summary>
        ///<param name="sheetIndex">シート番号</param>
        ///<param name="reName">変更後のシート名</param>
        ///<example>
        /// 次のコードでは、シート番号「1」のシート名を「テスト」に変更する。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     xls.SetSheetName("Sheet1","テスト");
        /// }
        /// </code>
        ///</example>
        public void SetSheetName(int sheetIndex, string reName)
        {
            object sheet = null;

            try
            {
                sheet = GetSheet(xlsSheets, sheetIndex);
                object[] parameters = new object[PARAM_NUM_1];
                parameters[0] = reName;
                sheet.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, sheet, parameters);
            }
            finally
            {
                ReleaseComObject(sheet);
            }
        }

        ///<summary>
        ///シート数を取得する。
        ///</summary>
        ///<example>
        /// 次のコードでは、シート数を取得する。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     xls.SetSheetName("Sheet1","テスト");
        /// }
        /// </code>
        ///</example>
        public int GetSheetCount()
        {
            return (int)xlsSheets.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, xlsSheets, null);
        }


        ///<summary>
        ///指定したシートを一番後ろにコピーし、新しい名称に設定する。
        ///</summary>
        ///<param name="sheetName">シート名</param>
        ///<param name="newName">新シート名</param>
        ///<example>
        /// 次のコードでは、シート名を「Sheet1］から「テスト」に変更する。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     xls.SetSheetName("Sheet1","テスト");
        /// }
        /// </code>
        ///</example>
        public void CopySheet(string sheetName, string newName)
        {
            object sheet = null;
            object sheetBack = null;
            try
            {
                sheet = GetSheet(xlsSheets, sheetName);
                sheetBack = GetSheet(xlsSheets, GetSheetCount());
                object[] parameters = new object[PARAM_NUM_2];
                parameters[0] = Type.Missing;
                parameters[1] = sheetBack;
                sheet.GetType().InvokeMember("Copy", BindingFlags.InvokeMethod, null, sheet, parameters);
                SetSheetName(GetSheetCount(), newName);
            }
            finally
            {
                ReleaseComObject(sheet);
                ReleaseComObject(sheetBack);
            }
        }

        ///<summary>
        ///指定したシートを一番後ろにコピーし、新しい名称に設定する。
        ///</summary>
        ///<param name="sheetIndex">シート番号</param>
        ///<param name="newName">新シート名</param>
        ///<example>
        /// 次のコードでは、シート番号「1」のシート名を「テスト」に変更する。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     xls.SetSheetName(1,"テスト");
        /// }
        /// </code>
        ///</example>
        public void CopySheet(int sheetIndex, string newName)
        {
            object sheet = null;
            try
            {
                sheet = GetSheet(xlsSheets, sheetIndex);
                object[] parameters = new object[PARAM_NUM_2];
                parameters[0] = Type.Missing;
                parameters[1] = GetSheet(xlsSheets, GetSheetCount()); ;
                sheet.GetType().InvokeMember("Copy", BindingFlags.InvokeMethod, null, sheet, parameters);
                SetSheetName(GetSheetCount(), newName);
            }
            finally
            {
                ReleaseComObject(sheet);
            }
        }

        ///<summary>
        ///指定したレンジ範囲に罫線を設定する。
        ///</summary>
        ///<param name="sheetName">シート名称</param>
        ///<param name="rangeMap">レンジ範囲</param>
        ///<param name="xlBordersIndex">罫線の位置</param>
        ///<param name="xlLineStyle">罫線の種類※省略時は実線</param>
        ///<param name="xlBorderWeight">罫線の太さ※省略時は細</param>
        ///<param name="color">罫線の色※省略時は黒</param>
        ///<example>
        /// 次のコードでは、シート名「Sheet1」のレンジ「B2:C4」の上に罫線を設定する。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     xls.SetRangeLine("Sheet1", "B2:C4", ExcelWrapper.XlBordersIndex.xlEdgeTop);
        /// }
        /// </code>
        ///</example>
        public void SetRangeLine(string sheetName, string rangeMap, 
                                XlBordersIndex xlBordersIndex,
                                XlLineStyle xlLineStyle = XlLineStyle.xlContinuous,
                                XlBorderWeight xlBorderWeight = XlBorderWeight.xlThin,
                                int color = 1)
        {
            object sheet = null;
            object range = null;
            object boders = null;
            
            try
            {
                sheet = GetSheet(xlsSheets, sheetName);
                range = GetRange(sheet, rangeMap);
                boders = GetBorders(range, xlBordersIndex);

                object[] parameters = new object[PARAM_NUM_1];
 
                parameters[0] = xlLineStyle;
                boders.GetType().InvokeMember("LineStyle", BindingFlags.SetProperty, null, boders, parameters);

                parameters[0] = xlBorderWeight;
                boders.GetType().InvokeMember("Weight", BindingFlags.SetProperty, null, boders, parameters);

                parameters[0] = color;
                boders.GetType().InvokeMember("ColorIndex", BindingFlags.SetProperty, null, boders, parameters);

            }
            finally
            {
                ReleaseComObject(boders);
                ReleaseComObject(range);
                ReleaseComObject(sheet);
            }
        }

        ///<summary>
        ///指定したレンジ範囲に罫線を設定する。
        ///</summary>
        ///<param name="sheetIndex">シート番号</param>
        ///<param name="rangeMap">レンジ範囲</param>
        ///<param name="xlBordersIndex">罫線の位置</param>
        ///<param name="xlLineStyle">罫線の種類※省略時は実線</param>
        ///<param name="xlBorderWeight">罫線の太さ※省略時は細</param>
        ///<param name="color">罫線の色※省略時は黒</param>
        ///<example>
        /// 次のコードでは、シート番号「1」のレンジ「B2:C4」の上に罫線を設定する。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     xls.SetRangeLine(1, "B2:C4", ExcelWrapper.XlBordersIndex.xlEdgeTop);
        /// }
        /// </code>
        ///</example>
        public void SetRangeLine(int sheetIndex, string rangeMap,
                                XlBordersIndex xlBordersIndex,
                                XlLineStyle xlLineStyle = XlLineStyle.xlContinuous,
                                XlBorderWeight xlBorderWeight = XlBorderWeight.xlThin,
                                int color = 1)
        {
            object sheet = null;
            object range = null;
            object boders = null;

            try
            {
                sheet = GetSheet(xlsSheets, sheetIndex);
                range = GetRange(sheet, rangeMap);
                boders = GetBorders(range, xlBordersIndex);

                object[] parameters = new object[PARAM_NUM_1];

                parameters[0] = xlLineStyle;
                boders.GetType().InvokeMember("LineStyle", BindingFlags.SetProperty, null, boders, parameters);

                parameters[0] = xlBorderWeight;
                boders.GetType().InvokeMember("Weight", BindingFlags.SetProperty, null, boders, parameters);

                parameters[0] = color;
                boders.GetType().InvokeMember("ColorIndex", BindingFlags.SetProperty, null, boders, parameters);

            }
            finally
            {
                ReleaseComObject(boders);
                ReleaseComObject(range);
                ReleaseComObject(sheet);
            }
        }

        ///<summary>
        ///指定したレンジ範囲のフォントの色を変更する。
        ///</summary>
        ///<param name="sheetName">シート名</param>
        ///<param name="rangeMap">レンジ範囲</param>
        ///<param name="color">フォントの色</param>
        ///<example>
        /// 次のコードでは、シート名「Sheet1］のレンジ「A1」のフォント色を赤色に変更する。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     xls.SetRangeFontColor("Sheet1", "A1", 3);
        /// }
        /// </code>
        ///</example>
        public void SetRangeFontColor(string sheetName, string rangeMap, int color)
        {
            object sheet = null;
            object range = null;
            object font = null;
            try
            {
                sheet = GetSheet(xlsSheets, sheetName);
                range = GetRange(sheet, rangeMap);
                font =  range.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, range, null);

                object[] parameters = new object[PARAM_NUM_1];
                parameters[0] = color;
                font.GetType().InvokeMember("ColorIndex", BindingFlags.SetProperty, null, font, parameters);
            }
            finally
            {
                ReleaseComObject(font);
                ReleaseComObject(range);
                ReleaseComObject(sheet);
            }
        }

        ///<summary>
        ///指定したレンジ範囲のフォントの色を変更する。
        ///</summary>
        ///<param name="sheetIndex">シート番号</param>
        ///<param name="rangeMap">レンジ範囲</param>
        ///<param name="color">フォントの色</param>
        ///<example>
        /// 次のコードでは、シート番号「1］のレンジ「A1」のフォント色を赤色に変更する。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///     xls.SetRangeFontColor(1, "A1", 3);
        /// }
        /// </code>
        ///</example>
        public void SetRangeFontColor(int sheetIndex, string rangeMap, int color)
        {
            object sheet = null;
            object range = null;
            object font = null;
            try
            {
                sheet = GetSheet(xlsSheets, sheetIndex);
                range = GetRange(sheet, rangeMap);
                font = range.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, range, null);

                object[] parameters = new object[PARAM_NUM_1];
                parameters[0] = color;
                font.GetType().InvokeMember("ColorIndex", BindingFlags.SetProperty, null, font, parameters);
            }
            finally
            {
                ReleaseComObject(font);
                ReleaseComObject(range);
                ReleaseComObject(sheet);
            }
        }

        ///<summary>
        ///指定したレンジ範囲の背景色を変更する。
        ///</summary>
        ///<param name="sheetName">シート名</param>
        ///<param name="rangeMap">レンジ範囲</param>
        ///<param name="color">フォントの色</param>
        ///<example>
        /// 次のコードでは、シート名「Sheet1］のレンジ「A1」の背景色を赤色に変更する。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///    xls.SetRangePatternColor("Sheet1", "A1", 3);
        /// }
        /// </code>
        ///</example>
        public void SetRangePatternColor(string sheetName, string rangeMap, int color)
        {
            object sheet = null;
            object range = null;
            object interior = null;
            try
            {
                sheet = GetSheet(xlsSheets, sheetName);
                range = GetRange(sheet, rangeMap);
                interior = range.GetType().InvokeMember("Interior", BindingFlags.GetProperty, null, range, null);

                object[] parameters = new object[PARAM_NUM_1];
                parameters[0] = color;
                interior.GetType().InvokeMember("ColorIndex", BindingFlags.SetProperty, null, interior, parameters);
            }
            finally
            {
                ReleaseComObject(interior);
                ReleaseComObject(range);
                ReleaseComObject(sheet);
            }
        }

        ///<summary>
        ///指定したレンジ範囲の背景色を変更する。
        ///</summary>
        ///<param name="sheetIndex">シート番号</param>
        ///<param name="rangeMap">レンジ範囲</param>
        ///<param name="color">フォントの色</param>
        ///<example>
        /// 次のコードでは、シート番号「1」のレンジ「A1」の背景色を赤色に変更する。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///     xls.Open("C:\Test.xls");
        ///    xls.SetRangePatternColor(1, "A1", 3);
        /// }
        /// </code>
        ///</example>
        public void SetRangePatternColor(int sheetIndex, string rangeMap, int color)
        {
            object sheet = null;
            object range = null;
            object interior = null;
            try
            {
                sheet = GetSheet(xlsSheets, sheetIndex);
                range = GetRange(sheet, rangeMap);
                interior = range.GetType().InvokeMember("Interior", BindingFlags.GetProperty, null, range, null);

                object[] parameters = new object[PARAM_NUM_1];
                parameters[0] = color;
                interior.GetType().InvokeMember("ColorIndex", BindingFlags.SetProperty, null, interior, parameters);
            }
            finally
            {
                ReleaseComObject(interior);
                ReleaseComObject(range);
                ReleaseComObject(sheet);
            }
        }

        ///<summary>
        ///指定したシートをアクティブに設定する。
        ///</summary>
        ///<param name="sheetName">シート名</param>
        ///<example>
        /// 次のコードでは、シート名「Sheet1」をアクティブに設定する。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///    xls.Open("C:\Test.xls");
        ///    xls.SetActivateSheet("Sheet1");
        /// }
        /// </code>
        ///</example>
        public void SetActivateSheet(string sheetName)
        {
            object sheet = null;
            try
            {
                sheet = GetSheet(xlsSheets, sheetName);
                sheet.GetType().InvokeMember("Activate", BindingFlags.InvokeMethod, null, sheet, null);
            }
            finally
            {
                ReleaseComObject(sheet);
            }
        }

        ///<summary>
        ///指定したシートをアクティブに設定する。
        ///</summary>
        ///<param name="sheetIndex">シート番号</param>
        ///<example>
        /// 次のコードでは、シート番号「1」をアクティブに設定する。
        /// <code>
        /// using (ExcelWrapper xls = new ExcelWrapper()){
        ///    xls.Open("C:\Test.xls");
        ///    xls.SetActivateSheet(1);
        /// }
        /// </code>
        ///</example>
        public void SetActivateSheet(int sheetIndex)
        {
            object sheet = null;
            try
            {
                sheet = GetSheet(xlsSheets, sheetIndex);
                sheet.GetType().InvokeMember("Activate", BindingFlags.InvokeMethod, null, sheet, null);
            }
            finally
            {
                ReleaseComObject(sheet);
            }
        }
    }
}
