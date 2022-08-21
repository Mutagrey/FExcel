using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Threading;
using FastExcelDNA.ExcelDNA.ADOUtil;
using System.IO;

namespace FastExcelDNA
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [ProgId("ExcelDNA.ADOCOM")]
    public class ADOCOM
    {
        #region Private fields
        private Dictionary<string, object> _adoResults = new Dictionary<string, object>(); //Резульаты загрузки.
        private object _sortedData = new object(); // Отсортированный массив данных по строкам.
        private object _sortedRows = new object(); // Отсортированный массив строк.
        private const string adoSpliter = ";";
        private int taskCounter = 0;
        private int finishedTaskCounter = 0;
        private int openfilesCounter = 0;
        private int filesCounter = 0;
        private readonly object sync = new object();
        #endregion

        #region Переменные видимые для VBA
        // Массив результатов. Для VBA.
        public object ResValues { get { return _adoResults.Values.ToArray() as object; } }
        // Массив ключей. Для VBA.
        public object ResKeys { get { return _adoResults.Keys.ToArray() as object; } }
        // Отсортированный массив данных по строкам. Для VBA.
        public object SortedData { get { return _sortedData; } }
        // Отсортированный массив строк. Для VBA.
        public object SortedRows { get { return _sortedRows; } }
        // Status результатов. Для VBA.
        public string Status = "ReadyToRun";
        // Все возможные виды статусов
        public readonly string[] StatusTypes =  { "ReadyToRun", "Preparing", "Calculating", "Failed", "Completed" };
        #endregion

        #region Асинхронное копирование файлов в указанную дирректорию
        // Копировать файлы в указанную дирректорию
        public object CopyFilesAsync(object filesToCopy, string destinationPath)
        {
            try
            {
                FileAttributes attr = File.GetAttributes(destinationPath);
                if (!attr.HasFlag(FileAttributes.Directory))
                    return "Ошибка в имени дирректории. Укажите путь к каталогу!";

                var res = new List<string>();
                string[] sFiles = filesToCopy as string[];
                var tasks = new Dictionary<string, Task>();
                foreach (var file in sFiles)
                {
                    var fileName = Path.GetFileName(file);
                    res.Add(destinationPath + fileName);
                    if (!tasks.ContainsKey(file) && fileName.ToLower().Contains(".xls"))
                        tasks.Add(file, CopyFileAsync(file, destinationPath));
                }
                Task.WaitAll(tasks.Values.ToArray());
                return res.ToArray();
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        // Асинхронная задача копирования файла в указанную дирректорию
        private async Task CopyFileAsync(string sourcePath, string destinationPath)
        {
            var fileName = Path.GetFileName(sourcePath);
            using (Stream source = File.Open(sourcePath, FileMode.Open))
            using (Stream destination = File.Create(destinationPath + fileName))
                await source.CopyToAsync(destination);
        }
        #endregion

        #region Асинхронная загрузка
        // Основной расчет. Запускаем загрузку асинхронно.
        public async void CalcADOAsync(object RowsID, object files, object sheets, object ranges, object offsetRW, object offsetCL, object resizeRW, object resizeCL, object tagInfo, bool isNumericOnly)
        {
            // 1. Словарь файлов и параметров для загрузки с учетом объединения в единый запрос по листу
            Status = "Preparing";
            var filesToLoad = await Task.Run(() =>
            {
                //var loadListFirst = GetLoadBookSheetsList(IDs, files, sheets, ranges, offsetRW, offsetCL, resizeRW, resizeCL, isNumericOnly, tagInfo);
                //return CombineSheetRequests(loadListFirst);
                return GetLoadList(RowsID, files, sheets, ranges, offsetRW, offsetCL, resizeRW, resizeCL, isNumericOnly, tagInfo);
            });
            // 2. Запуск задач асинхронно по каждому файлу и списку загрузки
            Status = "Calculating";
            var tasks = new List<Task>();
            filesCounter = filesToLoad.Count;
            foreach (var item in filesToLoad)
            {
                openfilesCounter++;
                tasks.Add(ReadAsync(item.Key, item.Value));
            }
            // 3. Ждем результатов задач загрузки
            try
            {
                await Task.WhenAll(tasks);                                             // Получаем основные результаты
                //_adoResults = await Task.Run(() => SplitSheetRequests(filesToLoad));   // Разделяем результаты по диапазонам
                await Task.Run(() => SortDataForPrint());                              // Получаем отсортированные по строка результаты
                Status = String.Format("Completed!...({0}/{1})...Поток: {2}", finishedTaskCounter, taskCounter, Thread.CurrentThread.ManagedThreadId);
            }
            catch (Exception ex)
            {
                Status = "Failed!    " + ex.Message;
            }
        }

        // Асинхронное чтение данных текущего файла
        private async Task ReadAsync(string file, IList<string> itemsToLoad)
        {
            using (OleDbConnection connection = new OleDbConnection(ADOManager.GetConnectionString(file)))
            {
                try
                {
                    var fileName = Path.GetFileName(file);
                    //Status = "Openning... " + fileName;
                    Status = String.Format("Openning...({0}/{1})...Поток: {2}...{3}", openfilesCounter, filesCounter, Thread.CurrentThread.ManagedThreadId, fileName);
                    // Открываем подключение
                    await connection.OpenAsync(AddInManager.Cancellation.Token); // AddInManager.Cancellation.Token
                    // Создаем задачи
                    Status = String.Format("Create tasks for {0}...", fileName);
                    var tasks = new Dictionary<Task<object>, string>();
                    foreach (var curItem in itemsToLoad)
                    {
                        var split = curItem.Split(char.Parse(adoSpliter));
                        var id = Convert.ToInt32(split[0]);
                        var sheet = split[1];
                        var range = split[2];
                        var oRW = int.Parse(split[3]);
                        var oCL = int.Parse(split[4]);
                        var rRW = int.Parse(split[5]);
                        var rCL = int.Parse(split[6]);
                        var isNumOnly = bool.Parse(split[7]);
                        var tag = split[8];

                        var key = string.Join(adoSpliter, new object[] { id, fileName, sheet, range, tag });
                        var task = ADOManager.ADO_ReadDataFormulaAsync(connection, file, sheet, range, oRW, oCL, rRW, rCL, isNumOnly, AddInManager.Cancellation.Token);

                        if (!tasks.ContainsKey(task))
                        {
                            lock (sync)
                            {
                                taskCounter++;
                            }
                            tasks.Add(task, key); //  AddInManager.Cancellation.Token
                        }
                    }
                    // Фильтруем задачи по текущему файлу из которого загружаем данные
                    while (tasks.Count > 0)
                    {
                        var finishedTask = await Task.WhenAny(tasks.Keys);
                        var key = tasks[finishedTask];
                        if (!_adoResults.ContainsKey(key))
                            _adoResults.Add(key, finishedTask.Result);
                        lock (sync)
                        {
                            finishedTaskCounter++;
                        }
                        Status = String.Format("Calculating...({0}/{1})...Поток: {3}...[ {2} ]", finishedTaskCounter, taskCounter, key, Thread.CurrentThread.ManagedThreadId);
                        tasks.Remove(finishedTask);
                    }
                }
                catch (Exception ex)
                {
                    Status = Status + "  Failed: " + ex.Message.Remove(100);
                }
            }

        }
        #endregion

        #region Cинхронная загрузка
        // Основной расчет. Запускаем загрузку синхронно.
        public void CalcADOSync(object IDs, object files, object sheets, object ranges, object offsetRW, object offsetCL, object resizeRW, object resizeCL, bool isNumericOnly, object tagInfo)
        {
            try
            {
                //var t = Task.Run(() =>
                //{

                //});

                Status = "Prepare";
                // Словарь файлов и параметров для загрузки
                var filesToLoad = GetLoadList(IDs, files, sheets, ranges, offsetRW, offsetCL, resizeRW, resizeCL, isNumericOnly, tagInfo);
                // Запуск задач синхронно по каждому файлу и списку загрузки  
                Status = "Calculating";
                filesCounter = filesToLoad.Count;
                foreach (var item in filesToLoad)
                {
                    openfilesCounter++;
                    ReadSync(item.Key, item.Value);
                }
                SortDataForPrint(); // Получаем отсортированные по строка результаты
                Status = String.Format("Completed!...({0}/{1})...Поток: {2}", finishedTaskCounter, taskCounter, Thread.CurrentThread.ManagedThreadId);
            }
            catch (Exception ex)
            {
                Status = "Failed!    " + ex.Message;
            }
        }

        private void ReadSync(string file, IList<string> itemsToLoad)
        {
            using (OleDbConnection connection = new OleDbConnection(ADOManager.GetConnectionString(file)))
            {
                try
                {
                    // Open connection sync
                    connection.Open();
                    foreach (var curItem in itemsToLoad)
                    {
                        var split = curItem.Split(char.Parse(adoSpliter));
                        var id = Convert.ToInt32(split[0]);
                        var sheet = split[1];
                        var range = split[2];
                        var oRW = int.Parse(split[3]);
                        var oCL = int.Parse(split[4]);
                        var rRW = int.Parse(split[5]);
                        var rCL = int.Parse(split[6]);
                        var isNumOnly = bool.Parse(split[7]);
                        var tag = split[8];

                        var key = string.Join(adoSpliter, new object[] { id, file, sheet, range, tag });
                        Status = "Calculating..." + key;
                        // Загрузка синхронная
                        var res = ADOManager.ADO_ReadDataFormula(connection, file, sheet, range, oRW, oCL, rRW, rCL, isNumOnly);
                        // Выводим результаты
                        if (!_adoResults.ContainsKey(key))
                            _adoResults.Add(key, res);
                    }
                }
                catch (Exception ex)
                {
                    if (!_adoResults.ContainsKey(ex.Message))
                        _adoResults.Add(ex.Message, ex.Message);
                }
            }

        }
        #endregion

        #region Вспомогательные методы
        // Получаем список файлов для загрузки
        private Dictionary<string, IList<string>> GetLoadList(object IDs, object files, object sheets, object ranges, object offsetRW, object offsetCL, object resizeRW, object resizeCL, bool isNumericOnly, object tagInfo)
        {
            short[] sIDs = IDs as short[];
            string[] sFiles = files as string[];
            string[] sSheets = sheets as string[];
            string[] sRanges = ranges as string[];
            short[] sOffsetRW = offsetRW as short[];
            short[] sOffsetCL = offsetCL as short[];
            short[] sResizeRW = resizeRW as short[];
            short[] sResizeCL = resizeCL as short[];
            string[] sTagInfo = tagInfo as string[];
            // Словарь файлов и параметров для загрузки
            Dictionary<string, IList<string>> filesToLoad = new Dictionary<string, IList<string>>();
            for (int i = 0; i < sFiles.Count(); i++)
            {
                var curID = sIDs[i];
                var file = sFiles[i];
                var sheet = sSheets[i];
                var range = sRanges[i];
                var oRW = sOffsetRW[i];
                var oCL = sOffsetCL[i];
                var rRW = sResizeRW[i];
                var rCL = sResizeCL[i];
                var tag = sTagInfo[i];

                var curKey = new object[] { curID, sheet, range, oRW, oCL, rRW, rCL, isNumericOnly, tag };
                Status = String.Format("Preparing...({0}/{1})...Поток: {3}...[ {2} ]", i, sFiles.Count(), string.Join(" | ", curKey), Thread.CurrentThread.ManagedThreadId);
                if (!filesToLoad.ContainsKey(file))
                    filesToLoad.Add(file, new List<string>());
                filesToLoad[file].Add(string.Join(adoSpliter, curKey));
            }
            return filesToLoad;
        }
        
        // Сортировка по строкам полученных результатов
        private void SortDataForPrint()
        {
            List<int> IDs = new List<int>(_adoResults.Keys.Select(k => Convert.ToInt32(k.Split(char.Parse(adoSpliter))[0])));
            IDs.Sort();
            var maxNum = IDs.Max() + 1;
            var temp = new object[maxNum, 4];

            Status = String.Format(Status + "   Сортировка по строкам...Поток: {0}", Thread.CurrentThread.ManagedThreadId);

            //Сортированный массив вместе с пустотами
            foreach (var kv in _adoResults)
            {
                var split = kv.Key.Split(char.Parse(adoSpliter));
                var id = Convert.ToInt32(split[0]);
                temp[id, 0] = id;
                temp[id, 1] = kv.Key;
                temp[id, 2] = kv.Value;
            }

            //Сортированный словарь по областям
            var areaID = "0"; var flag = false;
            var dicSorted = new Dictionary<string, List<object>>();
            for (var i = 0; i < maxNum; i++)
            {
                if (temp[i, 0] != null)
                {
                    if (!flag)
                    {
                        flag = !flag;
                        areaID = temp[i, 0].ToString();
                        if (!dicSorted.ContainsKey(temp[i, 0].ToString()))
                            dicSorted.Add(temp[i, 0].ToString(), new List<object>());
                    }
                    dicSorted[areaID].Add(temp[i, 2]);
                }
                else
                {
                    if (flag)
                        flag = !flag;
                }
            }
            // Сортированный массив по областям
            var resSortData = new object[dicSorted.Count];
            var resSortRows = new object[dicSorted.Count];
            var ind = 0;
            foreach (var kv in dicSorted)
            {
                resSortData[ind] = kv.Value.ToArray();
                resSortRows[ind] = kv.Key;
                ind++;
            }
            // Результаты
            _sortedData = resSortData;
            _sortedRows = resSortRows;
        }

        // Получаем список книг-листов для загрузки
        private Dictionary<string, Dictionary<string, IList<string>>> GetLoadBookSheetsList(object IDs, object files, object sheets, object ranges, object offsetRW, object offsetCL, object resizeRW, object resizeCL, bool isNumericOnly, object tagInfo)
        {
            short[] sIDs = IDs as short[];
            string[] sFiles = files as string[];
            string[] sSheets = sheets as string[];
            string[] sRanges = ranges as string[];
            short[] sOffsetRW = offsetRW as short[];
            short[] sOffsetCL = offsetCL as short[];
            short[] sResizeRW = resizeRW as short[];
            short[] sResizeCL = resizeCL as short[];
            string[] sTagInfo = tagInfo as string[];
            // Словарь файлов и параметров для загрузки
            var bookSheetsToLoad = new Dictionary<string, Dictionary<string, IList<string>>>();
            for (int i = 0; i < sFiles.Count(); i++)
            {
                var curID = sIDs[i];
                var file = sFiles[i];
                var sheet = sSheets[i];
                var range = sRanges[i];
                var oRW = sOffsetRW[i];
                var oCL = sOffsetCL[i];
                var rRW = sResizeRW[i];
                var rCL = sResizeCL[i];
                var tag = sTagInfo[i];

                var curKey = new object[] { curID, sheet, range, oRW, oCL, rRW, rCL, isNumericOnly, tag };
                Status = String.Format("Preparing...({0}/{1})...Поток: {3}...[ {2} ]", i, sFiles.Count(), string.Join(" | ", curKey), Thread.CurrentThread.ManagedThreadId);
                if (!bookSheetsToLoad.ContainsKey(file))
                    bookSheetsToLoad.Add(file, new  Dictionary<string, IList<string>>());
                if (!bookSheetsToLoad[file].ContainsKey(sheet))
                    bookSheetsToLoad[file].Add(sheet, new List<string>());// string.Join(adoSpliter, curKey));
                bookSheetsToLoad[file][sheet].Add(string.Join(adoSpliter, curKey));
            }
            return bookSheetsToLoad;
        }

        // Объединение диапазонов в единый запрос
        private Dictionary<string, IList<string>> CombineSheetRequests(Dictionary<string, Dictionary<string, IList<string>>> bookSheetDic)
        {
            var sheetRequests = new Dictionary<string, IList<string>>();
            
            foreach (var kv in bookSheetDic) // По книгам
            {
                var file = kv.Key;
                var index = 0;
                foreach (var sKV in kv.Value) // По листам
                {
                    var sheet = sKV.Key;
                    var combineReq = string.Empty;
                    var minRW = int.MaxValue; var maxRW = 0; var minCL = int.MaxValue; var maxCL = 0;
                    var isNumOnly = true;
                    foreach (var item in sKV.Value) // По диапазонам
                    {
                        var split = item.Split(char.Parse(adoSpliter));
                        var id = Convert.ToInt32(split[0]);
                        var range = split[2];
                        var oRW = int.Parse(split[3]);
                        var oCL = int.Parse(split[4]);
                        var rRW = int.Parse(split[5]);
                        var rCL = int.Parse(split[6]);
                        isNumOnly = bool.Parse(split[7]);
                        var tag = split[8];

                        var curAddress = ExcelRefConverter.RefToADO(range, oRW, oCL, rRW, rCL);
                        object[] curAddressVal = (object[])ExcelRefConverter.ToNumericCoordinates(curAddress);
                        if (Convert.ToInt32(curAddressVal[0]) < minRW)
                            minRW = Convert.ToInt32(curAddressVal[0]);
                        if (Convert.ToInt32(curAddressVal[1]) < minCL)
                            minCL = Convert.ToInt32(curAddressVal[1]);
                        if (Convert.ToInt32(curAddressVal[2]) > maxRW)
                            maxRW = Convert.ToInt32(curAddressVal[2]);
                        if (Convert.ToInt32(curAddressVal[3]) > maxCL)
                            maxCL = Convert.ToInt32(curAddressVal[3]);

                    }
                    var newRange = ExcelRefConverter.ToExcelCoordinates(minRW, minCL, maxRW, maxCL);
                    var curKey = new object[] { index, sheet, newRange, 0, 0, 0, 0, isNumOnly, "CombinedRange", minRW, minCL, maxRW, maxCL, "@", string.Join(adoSpliter, sKV.Value.ToArray()) };

                    if (!sheetRequests.ContainsKey(file))
                        sheetRequests.Add(file, new List<string>());
                    sheetRequests[file].Add(string.Join(adoSpliter, curKey));

                    index++;
                }
            }
            return sheetRequests;
        }

        // Обратное разделение результата загрузки по диапазонам
        private Dictionary<string, object> SplitSheetRequests(Dictionary<string, IList<string>> sheetRequests)
        {
            var splitedResults = new Dictionary<string, object>();
            foreach (var kv in sheetRequests) // По книгам
            {
                var file = kv.Key;
                //var loadedData = _adoResults[file];
                foreach (var item in kv.Value) // По листам
                {
                    var sheetsStr = item.Split(char.Parse(adoSpliter));

                    var id = Convert.ToInt32(sheetsStr[0]);
                    var sheet = sheetsStr[1];
                    var range = sheetsStr[2];
                    var oRW = int.Parse(sheetsStr[3]);
                    var oCL = int.Parse(sheetsStr[4]);
                    var rRW = int.Parse(sheetsStr[5]);
                    var rCL = int.Parse(sheetsStr[6]);
                    var isNumOnly = bool.Parse(sheetsStr[7]);
                    var tag = sheetsStr[8];

                    var minRW = sheetsStr[9];
                    var minCL = sheetsStr[10];
                    var maxRW = sheetsStr[11];
                    var maxCL = sheetsStr[12];
                    var list = sheetsStr[13].Split('@');

                    var loadedKey = string.Join(adoSpliter, new object[] { id, file, sheet, range, tag });
                    var loadedData = _adoResults[loadedKey];

                    foreach (var itm in list)   // По диапазонам
                    {
                        var split = itm.Split(char.Parse(adoSpliter));
                        var id0 = Convert.ToInt32(split[0]);
                        var sheet0 = split[1];
                        var range0 = split[2];
                        var oRW0 = int.Parse(split[3]);
                        var oCL0 = int.Parse(split[4]);
                        var rRW0 = int.Parse(split[5]);
                        var rCL0 = int.Parse(split[6]);
                        var isNumOnly0 = bool.Parse(split[7]);
                        var tag0 = split[8];

                        // Преобразовать в массив
                        object[] curAddressVal = (object[])ExcelRefConverter.ToNumericCoordinates(range);
                        var rowsCount = Convert.ToInt32(curAddressVal[2]) - Convert.ToInt32(curAddressVal[0]) + 1;
                        var colsCount = Convert.ToInt32(curAddressVal[3]) - Convert.ToInt32(curAddressVal[1]) + 1;
                        var resData = new object[rowsCount, colsCount];



                        for (var i = 0; i < rowsCount; i++)
                        {
                            for (var j = 0; j < rowsCount; j++)
                            {
                                resData[i, j] = 1;// loadedData[0, 0];
                            }
                        }


                        //var curDATA = loadedData[0,0];
                        var key = string.Join(adoSpliter, new object[] { id, file, sheet, range, tag });
                        if (!splitedResults.ContainsKey(key))
                            splitedResults.Add(key, resData);


                    }

                }
                
            }
            return splitedResults;
        }



        #endregion
    }
}
