using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.ComInterop;
using ExcelDna.Integration;
using NLog;

namespace NlogCL
{
    //VBA関数
    [System.Runtime.InteropServices.ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class ComLibrary
    {
        static NLog.Logger logger = LogManager.GetCurrentClassLogger();
        static NLog.Logger logger2 = LogManager.GetLogger("databaseLogger");

        public string ComLibraryHello()
        {
            return "Hello from NlogCL.ComLibrary";
        }

        public double Add(double x, double y)
        {
            return x + y;
        }
        public int InfoLog(string slog)
        {
            logger.Info("Hello World " + slog + ";");

            Exception ex = new Exception();
            logger.Error(ex, "Whoops!");

            logger.Info(slog);
            logger2.Info(slog);

            return 0;
        }
        public int ErrLog(string slog)
        {
            try
            {
                int zero = 0;
                int result = 5 / zero;  //ゼロで割り算して異常終了を発生させる
            }
            catch (DivideByZeroException ex)
            {
                // add custom message and pass in the exception
                logger.Error(ex, "Whoops! " + slog + ";");
                logger2.Error(ex, "Whoops! " + slog + ";");
            }
            return 0;
        }

    }

    [System.Runtime.InteropServices.ComVisible(false)]
    public class ExcelAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            ComServer.DllRegisterServer();
        }
        public void AutoClose()
        {
            ComServer.DllUnregisterServer();
        }
    }

    public static class Functions
    {
        static NLog.Logger logger = LogManager.GetCurrentClassLogger();
        static NLog.Logger logger2 = LogManager.GetLogger("databaseLogger");

        [ExcelFunction(Description = "TestHello .NET function")]
        public static object TestHello()
        {
            return "Hello from NlogCL!";
        }

        [ExcelFunction(Description = "SayHello .NET function")]
        public static string SayHello(string name)
        {
            return "Hello " + name;
        }

        [ExcelFunction(Description = "PrintInfoLog .NET function")]
        public static void PrintInfoLog(string slog)
        {
            logger.Info("Hello World " + slog + ";");

            Exception ex = new Exception();
            logger.Error(ex, "Whoops!");

            logger.Info(slog);
            logger2.Info(slog);

            return;
        }

        [ExcelFunction(Description = "PrintErrLog .NET function")]
        public static void PrintErrLog(string slog)
        {
            try
            {
                int zero = 0;
                int result = 5 / zero;  //ゼロで割り算して異常終了を発生させる
            }
            catch (DivideByZeroException ex)
            {
                // add custom message and pass in the exception
                logger.Error(ex, "Whoops! " + slog + ";");
                logger2.Error(ex, "Whoops! " + slog + ";");
            }
            return;
        }

        [ExcelFunction(Description = "MyMethod1 .NET function")]
        public static void MyMethod1()
        {
            //各ログレベルの出力サンプル
            logger.Trace("Sample trace message");
            logger.Debug("Sample debug message");
            logger.Info("Sample informational message");
            logger.Warn("Sample warning message");
            logger.Error("Sample error message");
            logger.Fatal("Sample fatal error message");

            // Exseption情報を出力する例
            try
            {
                //またはLog()メソッドにログレベルとメッセージを渡すことで出力することが可能
                logger.Log(LogLevel.Info, "Sample informational message");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "ow noos!"); // render the exception with ${exception}
                throw;
            }
            return;
        }

//        [ExcelFunction(Description = "DisplayMenu .NET function")]
        [ExcelCommand(MenuName = "AddIn", MenuText = "Worksheet")]
        public static void DisplayMenu()
        {

            //  新規作成したワークシートがアクティブになる
            MessageBox.Show("ワークシートを2枚(Sheet2, Sheet3)、新規作成します");
            XlCall.Excel(XlCall.xlcWorkbookInsert);
            XlCall.Excel(XlCall.xlcWorkbookInsert);


            //  存在しないワークシート名を指定するとエラーになる
            MessageBox.Show("ワークシート名を Sheet1 から hoge へ変更します");
            XlCall.Excel(XlCall.xlcWorkbookName, "Sheet1", "hoge");


            MessageBox.Show("アクティブになっているワークシート(Sheet3)を削除します");
            XlCall.Excel(XlCall.xlcWorkbookDelete);


            //  COM経由でワークシートの数を計算する(ワークシートの移動で使う)
            //  COM経由以外では、Excelのワークシート数を取得する方法がわからなかった
            dynamic app = ExcelDnaUtil.Application;
            var count = app.Worksheets.Count;


            MessageBox.Show("Sheet2を一番右へ移動します");
            XlCall.Excel(XlCall.xlcWorkbookMove, "Sheet2", "test.xlsx", count);


            //  一番左のワークシートの番号は 1
            MessageBox.Show("Sheet2を一番左へ移動します");
            XlCall.Excel(XlCall.xlcWorkbookMove, "Sheet2", "test.xlsx", 1);


            //  一番右に新規作成するので、現在の枚数に +1 しておく
            MessageBox.Show("Sheet2を、ブック内の一番右にコピーして移動します");
            XlCall.Excel(XlCall.xlcWorkbookCopy, "Sheet2", "test.xlsx", count + 1);


            MessageBox.Show("Sheet2を選択してアクティブにします");
            XlCall.Excel(XlCall.xlcWorkbookActivate, "Sheet2");


            MessageBox.Show("Sheet2を非表示にします");
            XlCall.Excel(XlCall.xlcWorkbookHide, "Sheet2");


            MessageBox.Show("Sheet2を再表示します");
            XlCall.Excel(XlCall.xlcWorkbookUnhide, "Sheet2");


            //  左上が「0, 0」で始まる
            MessageBox.Show("アクティブなワークシートのセルに値を設定します");
            var hoge = new ExcelReference(0, 0).SetValue("fuga");
            var fuga = new ExcelReference(5, 5).SetValue("piyo");


            //  セルの値を取得する
            var piyo = new ExcelReference(0, 0).GetValue().ToString();
            MessageBox.Show("セルA1の値は " + piyo + " です");

            //  セルの選択：R1C1形式で記述する必要あり、選択しているセルの相対参照でもある
            //  第一引数は範囲、第二引数は範囲内でアクティブになっているセル
            //  See: http://www.moug.net/tech/exvba/0050098.html
            MessageBox.Show("セルを範囲選択します");
            XlCall.Excel(XlCall.xlcSelect, "R[0]C[0]:R[4]C[4]", "R[0]C[0]");


            MessageBox.Show("選択範囲内のセルの値をクリアします");
            XlCall.Excel(XlCall.xlcClear);


            MessageBox.Show("全シートを選択します");
            XlCall.Excel(XlCall.xlcSelectAll);

            return;
        }
    }
}
