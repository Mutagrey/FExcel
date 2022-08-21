using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Threading;
using Microsoft.VisualBasic;
using System.Diagnostics;
using ExcelDna.Integration;
using System.IO;

namespace FastExcelDNA.ExcelDNA.ADOUtil
{
    public static class ADOManager
    {
        private const string categoryFunction = "ADO READ DATA";

        #region Локальные переменные. Открытые подключения и статистика по ним.
        // Пул подключений - все существующие подключения в текущей сессии
        private static Dictionary<string, OleDbConnection> opennedConnections = new Dictionary<string, OleDbConnection>();
        // Инфо подключений для вывода статистики: подключение, время_подключения
        public static Dictionary<string, string> connectionsInfo = new Dictionary<string, string>();
        // Количество запросов в книге - в текущей сессии
        private static Dictionary<string, int> adoRequestCount = new Dictionary<string, int>();
        // Суммарное Количество запросов - в текущей сессии
        private static int adoSumRequestCount {
            get {
                int sumRequestCount = 0;
                foreach (var kv in adoRequestCount)
                    sumRequestCount += kv.Value;
                return sumRequestCount;
            }
        }
        #endregion

        #region Создание и обработка подключений 
        // Получаем строку подключения для файла Excel
        public static string GetConnectionString(string filePath)
        {
            string conSTR = "";
            string provider = "Microsoft.ACE.OLEDB.12.0";
            filePath = filePath.ToLower();
            if (filePath.Contains("xlsx"))
            {
                conSTR = "Provider=" + provider + ";Data Source=" + filePath + "; " + "Extended Properties=\"Excel 12.0 Xml;HDR=No;IMEX=1\";";
            }
            else if (filePath.Contains("xlsm"))
            {
                conSTR = "Provider=" + provider + ";Data Source=" + filePath + "; " + "Extended Properties=\"Excel 12.0 Macro;HDR=No;IMEX=1\";";
            }
            else if (filePath.Contains("xlsb"))
            {
                conSTR = "Provider=" + provider + ";Data Source=" + filePath + "; " + "Extended Properties=\"Excel 12.0;HDR=No;IMEX=1\";";
            }
            else if (filePath.Contains("xls"))
            {
                conSTR = "Provider=" + provider + ";Data Source=" + filePath + "; " + "Extended Properties=\"Excel 8.0;HDR=No;IMEX=1\";";
            }
            return conSTR;
        }

        // Прозвон подключений, проверка работы пула подключений методов (OpenConnectionAsync, OpenConnectionSync), есть ли соревнование в подключени. Как ведут себя подключения при многопоточных запросах?
        // При тестировании вроде открывает все подключения корректно, несмотря на потоконезависимость, конфликтов вроде не возникало, но нужно очень осторожно быть с этим!
        [ExcelFunction(IsThreadSafe = true)]
        public static async void OpenConnectionEchoTest(string filePath, bool isAsyncCall, ExcelAsyncHandle asyncHandle)
        {
            try
            {
                if (isAsyncCall)
                {
                    var con = await OpenConnectionAsync(filePath);
                    asyncHandle.SetResult(con.State.ToString());
                }
                else
                {
                    var con = OpenConnectionSync(filePath);
                    asyncHandle.SetResult(con.State.ToString());
                }
            }
            catch (Exception ex)
            {
                asyncHandle.SetException(ex);
            }
        }

        // Открываем подключение OleDbConnection синхронно
        public static OleDbConnection OpenConnectionSync(string filePath)
        {
            // Строка подключения
            string connectionString = GetConnectionString(filePath);
            // Подключение
            var con = new OleDbConnection(connectionString);

            try
            {
                if (opennedConnections.ContainsKey(filePath))
                {
                    return opennedConnections[filePath];
                }
                else
                {
                    // Таймер подключения
                    var stopwatch = new Stopwatch();
                    stopwatch.Start();
                    
                    // Подключаемся синхронно и добавляем в словарь
                    con.Open();
                    opennedConnections.Add(filePath, con);

                    // Окончание таймера подключения
                    stopwatch.Stop();
                    var elapsed = (double)stopwatch.ElapsedMilliseconds / 1000d;
                    // Добавить Инфо подключения в словарь
                    if (!connectionsInfo.ContainsKey(filePath))
                        connectionsInfo.Add(filePath, elapsed.ToString());
                    return con;
                }
            }
            catch (Exception ex)
            {
                // Добавить Инфо подключения в словарь
                if (!connectionsInfo.ContainsKey(filePath))
                    connectionsInfo.Add(filePath, ex.Message);
                return con;
            }
        }

        // Открываем подключение OleDbConnection асинхронно (async - >= .Net 4.5)
        public static async Task<OleDbConnection> OpenConnectionAsync(string filePath, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (opennedConnections.ContainsKey(filePath))
            {
                return opennedConnections[filePath];
            } 
            else 
            {
                // Строка подключения
                string connectionString = GetConnectionString(filePath);
                // Подключение
                var con = new OleDbConnection(connectionString);

                try
                {
                    // Таймер подключения
                    var stopwatch = new Stopwatch();
                    stopwatch.Start();

                    // Как правильно подключаться? Сначала добавить в словарь, а потом ждать подключение или наоборот, сначала подключаемся, а потом добавляем в словарь?

                    // Способ 1: Асинхронно ждем подключения, если задача завершилась успешно, то проверяем добавлено ли подключение в словарь (из другого потока) и добавляем, если не добавлено.
                    await con.OpenAsync(cancellationToken).ContinueWith(t =>
                    {
                        if (t.Status == TaskStatus.RanToCompletion)
                        {
                            if (!opennedConnections.ContainsKey(filePath))
                                opennedConnections.Add(filePath, con);
                        }
                    }, cancellationToken);

                    //// Способ 2: Ждем окончания подключения, после чего добавляем в словарь
                    //await con.OpenAsync(cancellationToken).ConfigureAwait(false);
                    //opennedConnections.Add(filePath, con);

                    // Окончание таймера подключения
                    stopwatch.Stop();
                    var elapsed = (double)stopwatch.ElapsedMilliseconds / 1000d;
                    // Добавить Инфо подключения в словарь
                    if (!connectionsInfo.ContainsKey(filePath))
                        connectionsInfo.Add(filePath, elapsed.ToString());

                    return con;
                }
                catch (Exception ex)
                {
                    // Добавить Инфо подключения в словарь
                    if (!connectionsInfo.ContainsKey(filePath))
                        connectionsInfo.Add(filePath, ex.Message);
                    return con;
                }

            }
        }
        #endregion

        #region Чтение данных через ADO Асинхронно
        // Считываем данные Excel через открытое подключение.
        private static async Task<object> ADO_ReadDataAsync(OleDbConnection connection, string filePath, string sheetName, string AdoCellRef, bool isNumericOnly = false, CancellationToken cancellationToken = default(CancellationToken))
        {
            try
            {
                // Проверяем, что подключение открыто
                if (connection.State != ConnectionState.Open)
                    return "Can't open, connection: " + connection.State.ToString();

                // Количество запросов в текущей книге
                if (!adoRequestCount.ContainsKey(filePath))
                    adoRequestCount.Add(filePath, 1);
                else
                    adoRequestCount[filePath] += 1;

                // Запрос
                string queryString = "SELECT * FROM [" + sheetName + "$" + AdoCellRef + "]";
                // Таблица в которой хранятся результаты загрузки
                System.Data.DataTable resTable = new System.Data.DataTable();
                // Создаем адаптер 
                OleDbDataAdapter adapter = new OleDbDataAdapter();
                // Указываем запрос для адаптера 
                adapter.SelectCommand = new OleDbCommand(queryString, connection);
                // Получить данные
                await Task.Run(() => adapter.Fill(resTable), cancellationToken);
                // Преобразовать в массив
                var res = await Task.Run(() =>
                {
                    var newRes = new object[resTable.Rows.Count, resTable.Columns.Count];
                    for (int i = 0; i < resTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < resTable.Columns.Count; j++)
                        {
                            var curData = resTable.Rows[i].ItemArray[j];
                            if (!Convert.IsDBNull(curData))
                                if (Information.IsNumeric(curData) || !isNumericOnly)
                                    newRes[i, j] = curData;
                        }
                    }
                    return newRes;
                }, cancellationToken);

                // Вычитаем количество на 1, т.к. уже завершили загрузку с этим подключением.
                if (adoRequestCount.ContainsKey(filePath))
                    adoRequestCount[filePath] -= 1;

                return res;
            }
            catch (Exception ex)
            {
                // Нужно ли вычитать в случае исключения??? Вроде нужно...
                if (adoRequestCount.ContainsKey(filePath))
                    adoRequestCount[filePath] -= 1;
                return ex.Message.ToString();
            }
        }
        
        // Асинхронно. Считываем данные Excel через открытое подключение. Получаем данные с учетом формул и изменений диапазона. 
        public static async Task<object> ADO_ReadDataFormulaAsync(OleDbConnection connection, string filePath, string sheetName, string sRngFormula, int OffsetRW, int OffsetCL, int ResizeRW, int ResizeCL, bool isNumericOnly = false, CancellationToken cancellationToken = default(CancellationToken))
        {
            try
            {
                if (connection.State != ConnectionState.Open)
                    return "Подключение закрыто!";

                // Распознаем в формуле все адреса типов XlA1, XlR1C1
                string[] formulas = await Task.Run(() => ExcelRefConverter.GetCellRefFromFormula(sRngFormula, ";").Split(';'), cancellationToken);

                if (formulas.Length > 1)
                {
                    // ---- Формулы - сначала получаем весь диапазон, после считаем формулу по каждой ячейке ---------- Ускорил кратно, возможно есть еще способ успорить
                    Dictionary<string, object> adoFormulaData = new Dictionary<string, object>();
                    foreach (string curRef in formulas)
                    {
                        // Преобразует адрес диапазона для ADO в нужный формат A1:A1
                        string ADOAddres = await Task.Run(() => ExcelRefConverter.RefToADO(curRef, OffsetRW, OffsetCL, ResizeRW, ResizeCL), cancellationToken);//, cancellationToken);

                        // READ DATA FROM ADO
                        object adoData = await ADO_ReadDataAsync(connection, filePath, sheetName, ADOAddres, isNumericOnly, cancellationToken);

                        // Add data to dictionary formulas
                        if (!adoFormulaData.ContainsKey(curRef))
                            adoFormulaData.Add(curRef, adoData);
                    }
                    // Результаты. Обработка и замена формул значениями. Пересчет через COM
                    var res = await Task.Run(() =>
                    {
                        object[,] resData = new object[ResizeRW, ResizeCL];
                        for (int i = 0; i < ResizeRW; i++)
                        {
                            for (int j = 0; j < ResizeCL; j++)
                            {
                                string resStr = sRngFormula;
                                string strToReplace = string.Empty;
                                int index = 0;
                                foreach (string curRef in formulas)
                                {
                                    var curData = adoFormulaData[curRef];
                                    //Заменяем значения текущей ячейки
                                    strToReplace = Convert.ToString(((object[,])curData)[i, j]);
                                    if (strToReplace.Length == 0 || !Information.IsNumeric(strToReplace))
                                        strToReplace = "0";
                                    index = resStr.IndexOf(curRef);
                                    if (index >= 0)
                                        resStr = resStr.Remove(index) + strToReplace + resStr.Substring(index + curRef.Length);
                                    index = index + strToReplace.Length;
                                }

                                resStr = resStr.Replace(",", ".");
                                var eval = AddInManager.EvalWithCOM("=" + resStr); // await???
                                resData[i, j] = eval;
                            }
                        }
                        return resData;
                    }, cancellationToken);


                    var fileName = Path.GetFileName(filePath);// filePath.Substring(filePath.LastIndexOf("\\") + 1);
                    var statusMessage = "Запросов осталось: (" + adoSumRequestCount + ")  Формула: " + sRngFormula + "  |  Файл: " + fileName + "  |  Лист: " + sheetName;
                    //ExcelAsyncUtil.QueueAsMacro(delegate
                    //{
                    //    XlCall.Excel(XlCall.xlcMessage, true, statusMessage);
                    //});
                    return res;
                }
                else
                {
                    // ---- Один Дианазон - считаем сразу через один запрос -----
                    // Преобразует адрес диапазона для ADO в нужный формат A1:A1
                    string ADOAddres = await Task.Run(() => ExcelRefConverter.RefToADO(sRngFormula, OffsetRW, OffsetCL, ResizeRW, ResizeCL), cancellationToken);//, cancellationToken);
                    // READ DATA FROM ADO
                    object adoData =  await ADO_ReadDataAsync(connection, filePath, sheetName, ADOAddres, isNumericOnly, cancellationToken);
                    var fileName = filePath.Substring(filePath.LastIndexOf("\\") + 1);
                    var statusMessage = "Запросов осталось: (" + adoSumRequestCount + ")  Диапазон: " + ADOAddres + "  |  Файл: " + fileName + "  |  Лист: " + sheetName;
                    //ExcelAsyncUtil.QueueAsMacro(delegate
                    //{
                    //    XlCall.Excel(XlCall.xlcMessage, true, statusMessage);
                    //});
                    return adoData;
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        
        //// Подключение и загрузка в одной операции. Сделана для удобства вызова. Не обязательна!
        //public static async Task<object> ADO_ConnectAndLoadAsync(string filePath, string sheetName, string sRngFormula, int OffsetRW, int OffsetCL, int ResizeRW, int ResizeCL, bool isNumericOnly = false, CancellationToken cancellationToken = default(CancellationToken))
        //{
        //    var con = await OpenConnectionAsync(filePath, AddInManager.Cancellation.Token);
        //    var res = await ADO_ReadDataFormulaAsync(con, filePath, sheetName, sRngFormula, OffsetRW, OffsetCL, ResizeRW, ResizeCL, isNumericOnly, AddInManager.Cancellation.Token);
        //    return res;
        //}
        #endregion

        #region Чтение данных через ADO Синхронно
        // Считываем данные Excel через открытое подключение.
        private static object ADO_ReadData(OleDbConnection connection, string filePath, string sheetName, string AdoCellRef, bool isNumericOnly = false)
        {
            try
            {
                // Проверяем, что подключение открыто
                if (connection.State != ConnectionState.Open)
                    return "Can't open, connection: " + connection.State.ToString();

                // Количество запросов в текущей книге
                if (!adoRequestCount.ContainsKey(filePath))
                    adoRequestCount.Add(filePath, 1);
                else
                    adoRequestCount[filePath] += 1;

                // УБРАЛ т.к преобразование делаем вне данной функции!!!! МОГУТ ВОЗНИКАТЬ ИСКЛЮЧЕНИЯ, если ссылка не соответствует формату А1:А1
                //// Преобразовать адрес в нужный формат А1:А1. 
                //var newAdoRef = ExcelRefConverter.RefToADO(AdoCellRef);

                // Запрос
                string queryString = "SELECT * FROM [" + sheetName + "$" + AdoCellRef + "]";
                // Таблица в которой хранятся результаты загрузки
                System.Data.DataTable resTable = new System.Data.DataTable();
                // Создаем адаптер 
                OleDbDataAdapter adapter = new OleDbDataAdapter();
                // Указываем запрос для адаптера 
                adapter.SelectCommand = new OleDbCommand(queryString, connection);
                // Получить данные
                adapter.Fill(resTable);
                // Преобразовать в массив
                var newRes = new object[resTable.Rows.Count, resTable.Columns.Count];
                for (int i = 0; i < resTable.Rows.Count; i++)
                {
                    for (int j = 0; j < resTable.Columns.Count; j++)
                    {
                        var curData = resTable.Rows[i].ItemArray[j];
                        if (!Convert.IsDBNull(curData))
                            if (Information.IsNumeric(curData) || !isNumericOnly)
                                newRes[i, j] = curData;
                    }
                }
                // Вычитаем количество на 1, т.к. уже завершили загрузку с этим подключением.
                if (adoRequestCount.ContainsKey(filePath))
                    adoRequestCount[filePath] -= 1;

                return newRes;
            }
            catch (Exception ex)
            {
                // Нужно ли вычитать в случае исключения??? Вроде нужно...
                if (adoRequestCount.ContainsKey(filePath))
                    adoRequestCount[filePath] -= 1;
                return ex.Message.ToString();
            }
        }

        // Считываем данные Excel через открытое подключение. Получаем данные с учетом формул и изменений диапазона. 
        public static object ADO_ReadDataFormula(OleDbConnection connection, string filePath, string sheetName, string sRngFormula, int OffsetRW, int OffsetCL, int ResizeRW, int ResizeCL, bool isNumericOnly = false)
        {
            try
            {
                if (connection.State != ConnectionState.Open)
                    return "Подключение закрыто!";

                // Распознаем в формуле все адреса типов XlA1, XlR1C1
                string[] formulas = ExcelRefConverter.GetCellRefFromFormula(sRngFormula, ";").Split(';');

                string strToReplace = string.Empty;
                int index = 0;
                //string resStr = sRngFormula;
                if (formulas.Length > 1)
                {
                    // ---- Формулы - считаем по одной ячейке ---------- Ускорил кратно, возможно есть еще способ успорить
                    Dictionary<string, object> adoFormulaData = new Dictionary<string, object>();
                    foreach (string curRef in formulas)
                    {
                        // Преобразует адрес диапазона для ADO в нужный формат A1:A1
                        string ADOAddres = ExcelRefConverter.RefToADO(curRef, OffsetRW, OffsetCL, ResizeRW, ResizeCL);

                        // READ DATA FROM ADO
                        object adoData = ADO_ReadData(connection, filePath, sheetName, ADOAddres, isNumericOnly);

                        // Add data to dictionary formulas
                        if (!adoFormulaData.ContainsKey(curRef))
                            adoFormulaData.Add(curRef, adoData);
                    }
                    // Результаты. Обработка и замена формул значениями. Пересчет через COM
                    object[,] resData = new object[ResizeRW, ResizeCL];
                    for (int i = 0; i < ResizeRW; i++)
                    {
                        for (int j = 0; j < ResizeCL; j++)
                        {
                            string resStr = sRngFormula;
                            foreach (string curRef in formulas)
                            {
                                var curData = adoFormulaData[curRef];
                                //Заменяем значения текущей ячейки
                                strToReplace = Convert.ToString(((object[,])curData)[i, j]);
                                if (strToReplace.Length == 0)
                                    strToReplace = "0";
                                index = resStr.IndexOf(curRef);
                                if (index >= 0)
                                    resStr = resStr.Remove(index) + strToReplace + resStr.Substring(index + curRef.Length);
                                index = index + strToReplace.Length;
                            }

                            resStr = resStr.Replace(",", ".");
                            var eval = AddInManager.EvalWithCOM("=" + resStr);
                            resData[i, j] = eval;
                        }
                    }

                    var fileName = filePath.Substring(filePath.LastIndexOf("\\") + 1);
                    var statusMessage = "Запросов осталось: (" + adoSumRequestCount + ")  Формула: " + sRngFormula + "  |  Файл: " + fileName + "  |  Лист: " + sheetName;
                    //ExcelAsyncUtil.QueueAsMacro(delegate
                    //{
                    //    XlCall.Excel(XlCall.xlcMessage, true, statusMessage);
                    //});
                    return resData;
                }
                else
                {
                    // ---- Один Дианазон - считаем сразу через один запрос -----
                    // Преобразует адрес диапазона для ADO в нужный формат A1:A1
                    string ADOAddres = ExcelRefConverter.RefToADO(sRngFormula, OffsetRW, OffsetCL, ResizeRW, ResizeCL);
                    // READ DATA FROM ADO
                    var adoData = ADO_ReadData(connection, filePath, sheetName, ADOAddres, isNumericOnly);
                    var fileName = filePath.Substring(filePath.LastIndexOf("\\") + 1);
                    var statusMessage = "Запросов осталось: (" + adoSumRequestCount + ")  Диапазон: " + ADOAddres + "  |  Файл: " + fileName + "  |  Лист: " + sheetName;
                    //ExcelAsyncUtil.QueueAsMacro(delegate
                    //{
                    //    XlCall.Excel(XlCall.xlcMessage, true, statusMessage);
                    //});
                    return adoData;
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        #endregion

        #region Чтение списка листов книги через ADO. Read Schema.
        // Считываем список листов - схему Excel через подключение
        private static object ADO_ReadSheetNamesExcelDNA(string filePath)
        {
            try
            {
                // Создаем и открываем подключение асинхронно и записываем его а КЭШ, чтобы не открывать повторно
                var connection = OpenConnectionSync(filePath);
                // Проверяем, что подключение открыто
                if (connection.State != ConnectionState.Open)
                    return "Подключение: " + connection.State.ToString();

                // Таблица со схемой книги - листы и прочие таблицы
                System.Data.DataTable infoTable = (System.Data.DataTable)connection.GetSchema("Tables");  //System.Data.DataTable infoTable = (System.Data.DataTable)connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                // Преобразовать в массив
                var res = new object[infoTable.Rows.Count];
                for (int i = 0; i < infoTable.Rows.Count; i++)
                    res[i] = ((DataRow)infoTable.Rows[i]).ItemArray[2];
                // Вернуть в качестве строки
                return String.Join(";", res.Where(c => c.ToString().Contains("$")));
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }
        
        [ExcelFunction(Category = categoryFunction, Description = "Возвращает список листов книги - асинхронно", Name = "ADO.ReadSheetNamesAsync")]
        public static object ADO_ReadSheetNamesAsync([ExcelArgument(Name = "filePath", Description = "путь к фалу excel")] string filePath)
        {
            object result = ExcelAsyncUtil.Run("ADO_ReadSheetNamesAsync", new object[] { filePath },
                delegate
                {
                    return ADOManager.ADO_ReadSheetNamesExcelDNA(filePath);
                });
            if (result.Equals(ExcelError.ExcelErrorNA))
                return "GettingSchema...";
            return result;
        }
        #endregion

        #region Статистика открытых подключений
        [ExcelFunction(Category = categoryFunction, Description = "Статистика существующих подключений в текущей книге", IsVolatile = true, Name = "ADO.OpennedConnectionsInfo")]
        public static object ADO_OpennedConnectionsInfo([ExcelArgument(Name = "Тип", Description = "0 - статистика, 1 - список со временем подключения, 2 - список с количеством подключений")] int statType = 0)
        {
            var info = "Открытых подключений: " + ADOManager.connectionsInfo.Count;
            // Количество запросов
            info += " Количество запросов: " + adoSumRequestCount;

            // Суммарное время подключения
            double sumOpenTime = 0;
            foreach (var kv in ADOManager.connectionsInfo)
            {
                double curTime = 0;
                double.TryParse(kv.Value, out curTime);
                sumOpenTime += curTime;
            }
            // Среднее время подключения
            double avgOpenTime = 0;
            if (ADOManager.connectionsInfo.Count > 0)
                avgOpenTime = sumOpenTime / ADOManager.connectionsInfo.Count;

            info += " (Суммарное время открытия: " + Math.Round(sumOpenTime, 3) + " сек)";
            info += " (Среднее время открытия: " + Math.Round(avgOpenTime, 3) + " сек)";

            if (statType == 0)
                return info;
            else if (statType == 1)
            {
                return String.Join(";", ADOManager.connectionsInfo.ToArray());
            }
            else if (statType == 2)
            {
                return String.Join(";", ADOManager.adoRequestCount.ToArray());
            }
            return info;
        }
        #endregion
    }
}
