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
    public class ExcelWrapper
    {
        ///<summary>Excelアプリケーションオブジェクト</summary>
        private object xlsApplication = null;
        ///<summary>Workbooksオブジェクト</summary>
        private object xlsBooks = null;

        public ExcelWrapper() { }

        ///<summary>Excelアプリケーションオブジェクト</summary>
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

        ///<summary>Excelのバージョン</summary>
        public string Version
        {
            get
            {
                object versionObj = XlsApplication.GetType().InvokeMember("Version", BindingFlags.GetProperty, null, XlsApplication, null);
                return versionObj.ToString();
            }
        }
    }
}
