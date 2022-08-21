using System;
using System.Threading;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System.Runtime.InteropServices;
using ExcelDna.ComInterop;

namespace FastExcelDNA
{
    [ComVisible(false)]
    public class AddInManager : IExcelAddIn
    {
        public static Microsoft.Office.Interop.Excel.Application xlApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
        //public static dynamic xlApp = ExcelDnaUtil.Application;

        public void AutoOpen()
        {
            ComServer.DllRegisterServer();
            IntelliSenseServer.Install();
            ExcelIntegration.RegisterUnhandledExceptionHandler(delegate(object ex) { return "!!! CRIT ERROR: " + ex.ToString(); });
            ExcelAsyncUtil.CalculationCanceled += CalculationCanceled;
            ExcelAsyncUtil.CalculationEnded += CalculationEnded;
        }
        public void AutoClose()
        {
            ComServer.DllUnregisterServer();
            IntelliSenseServer.Uninstall();
            ExcelAsyncUtil.CalculationCanceled -= CalculationCanceled;
            ExcelAsyncUtil.CalculationEnded -= CalculationEnded;
        }

        #region Cancellation support
        // We keep a CancellationTokenSource around, and set to a new one whenever a calculation has finished.
        public static CancellationTokenSource Cancellation = new CancellationTokenSource();

        // Обработчик событий когда все расчеты завершены
        void xlApp_AfterCalculate()
        {
            if (Cancellation.IsCancellationRequested)
                Cancellation = new CancellationTokenSource();
        }

        public static void CalculationCanceled()
        {
            Cancellation.Cancel();
            ExcelAsyncUtil.QueueAsMacro(delegate
            {
                XlCall.Excel(XlCall.xlcMessage, true, "Расчет отменен!");
            });
        }

        public static void CalculationEnded()
        {
            // Maybe we only need to set a new one when it was actually used...(when IsCanceled is true)?
            //Cancellation = new CancellationTokenSource();
            // Отключил, тк при нажатии отмена (ESC) происходит сброс токена, и при повторном обновлении листа все равно идет пересчет
            // Вместо этого добавил xlApp_AfterCalculate

            if (Cancellation.IsCancellationRequested)
                Cancellation = new CancellationTokenSource();
        }
        #endregion

        #region Расчет формул Excel
        [ExcelFunction(IsHidden = true, Description = "Делает расчет формулы через C API - XlCall.xlfEvaluate. Рассчитывает значение в памяти и выводит результат.")]
        public static object EvalWithCAPI(string expression)
        {
            object value = XlCall.Excel(XlCall.xlfEvaluate, expression);
            if (value is ExcelReference)
                return ((ExcelReference)value).GetValue();
            else
                return value;
        }

        //[ExcelFunction(IsHidden = true, IsMacroType = true, Description = "Делает расчет формулы через COM - Microsoft.Office.Interop.Excel. Выводит результат в ячейку А1 текущего листа, потом очищает.")]
        //public static object AdvancedEvaluate(string formula)
        //{
        //    try
        //    {
        //        xlApp.Range["A1"].FormulaLocal = formula;
        //        var res = xlApp.Range["A1"].Value2;
        //        xlApp.Range["A1"].ClearContents();
        //        return res;

        //    }
        //    catch (Exception ex)
        //    {
        //        return ex.Message;
        //    }
        //}

        // Делает расчет формулы через COM - Microsoft.Office.Interop.Excel. Предварительно сохраняем COM - Application, чтобы не нагружать процесс и делать расчеты асинхронно!!!

        [ExcelFunction(IsHidden = true, IsMacroType = true, Description = "Делает расчет формулы через COM - Microsoft.Office.Interop.Excel. Рассчитывает значение в памяти и выводит результат.")]
        public static object EvalWithCOM(string formula)
        {
            try
            {
                return xlApp.Evaluate(formula);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }


        #endregion
    }
}
