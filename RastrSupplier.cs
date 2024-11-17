using ASTRALib;
using CsvHelper;
using System;
using System.Globalization;
using System.IO;

namespace RustabTest
{
    /// <summary>
    /// Класс, осуществляющий взаимодействие с Rastrwin3
    /// </summary>
    public abstract class RastrSupplier
    {
        /// <summary>
        /// Экземпляр класса Rastr
        /// </summary>
        public static Rastr _rastr = new Rastr();

        /// <summary>
        /// Загрузка файла в рабочую область
        /// </summary>
        /// <param name="filePath">Путь до файла</param>
        /// <param name="shablon">Шаблон файла</param>
        public static void LoadFile(string filePath, string shablon)
        {
            _rastr.Load(RG_KOD.RG_REPL, filePath, shablon);
        }

        /// <summary>
        /// Сохранение файла
        /// </summary>
        /// <param name="fileName">Имя файла</param>
        /// <param name="shablon">Шаблон файла</param>
        public static void SaveFile(string fileName, string shablon)
        {
            // Проверка лицензии
            try
            {
                _rastr.Save(fileName, shablon);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// Создание файла
        /// </summary>
        /// <param name="shablon">Шаблон файла</param>
        public static void CreateFile(string shablon)
        {
            _rastr.NewFile(shablon);
        }

        /// <summary>
        /// Поиск индекса из таблицы по номеру
        /// </summary>
        /// <param name="tableName">Название таблицы</param>
        /// <param name="parameterName">Название параметра</param>
        /// <param name="number">Номер узла</param>
        /// <returns>Индекс</returns>
        public static int GetIndexByNumber(string tableName,
            string parameterName, int number)
        {
            ITable table = _rastr.Tables.Item(tableName);
            ICol columnItem = table.Cols.Item(parameterName);

            for (int index = 0; index < table.Count; index++)
            {
                if (columnItem.get_ZN(index) == number)
                {
                    return index;
                }
            }
            throw new Exception($"Элемент с номером {number} не найден");
        }

        /// <summary>
        /// Поиск индекса из таблицы по номеру
        /// </summary>
        /// <param name="tableName">Название таблицы</param>
        /// <param name="parameterName">Название параметра</param>
        /// <param name="value">Параметр узла</param>
        /// <returns>Индекс</returns>
        public static int GetIndexByValue(string tableName,
            string parameterName, string value)
        {
            ITable table = _rastr.Tables.Item(tableName);
            ICol columnItem = table.Cols.Item(parameterName);

            for (int index = 0; index < table.Count; index++)
            {
                if (columnItem.get_ZN(index) == value)
                {
                    return index;
                }
            }
            throw new Exception($"Элемент {value} не найден");
        }

        /// <summary>
        /// Задание логического значения
        /// </summary>
        /// <param name="tableName">Название таблицы</param>
        /// <param name="parameterName">Название параметра</param>
        /// <param name="number">Номер узла</param>
        /// <param name="chosenParameter">Целевой параметр</param>
        /// <param name="value">Записываемое значение</param>
        public static void SetBoolValue(string tableName, 
            string parameterName, int number,
            string chosenParameter, bool value)
        {
            ITable table = _rastr.Tables.Item(tableName);
            ICol columnItem = table.Cols.Item(chosenParameter);

            int index = GetIndexByNumber(tableName, parameterName, number);
            columnItem.set_ZN(index, value);
        }

        /// <summary>
        /// Задание строкового значения
        /// </summary>
        /// <param name="tableName">Название таблицы</param>
        /// <param name="parameterName">Название параметра</param>
        /// <param name="number">Номер узла</param>
        /// <param name="chosenParameter">Целевой параметр</param>
        /// <param name="value">Записываемое значение</param>
        public static void SetStringValue(string tableName,
            string parameterName, int number,
            string chosenParameter, string value)
        {
            ITable table = _rastr.Tables.Item(tableName);
            ICol columnItem = table.Cols.Item(chosenParameter);

            int index = GetIndexByNumber(tableName, parameterName, number);
            columnItem.set_ZN(index, value);
        }

        /// <summary>
        /// Коммутация ветви
        /// </summary>
        /// <param name="ipNumber">Номер начала ветви</param>
        /// <param name="iqNumber">Номер конца ветви</param>
        /// <param name="npNumber">Номер параллельности ветви</param>
        /// <param name="state">Состояние ветви</param>
        public static void ChangeBranchState(int ipNumber, int iqNumber,
            int npNumber, bool state)
        {
            ITable table = _rastr.Tables.Item("vetv");

            ICol ipColumnItem = table.Cols.Item("ip");
            ICol iqColumnItem = table.Cols.Item("iq");
            ICol npColumnItem = table.Cols.Item("np");
            ICol staColumnItem = table.Cols.Item("sta");

            for (int index = 0; index < table.Count; index++)
            {
                if ((ipColumnItem.get_ZN(index) == ipNumber) &&
                    (iqColumnItem.get_ZN(index) == iqNumber))
                {
                    if (npNumber == 0)
                    {
                        staColumnItem.set_ZN(index, state);
                        break;
                    }
                    else if (npColumnItem.get_ZN(index) == npNumber)
                    {
                        staColumnItem.set_ZN(index, state);
                        break;
                    }
                }

            }
        }

        /// <summary>
        /// Получение списка генераторов станции
        /// </summary>
        /// <param name="isPlantResearched">true - исследуемая станция, false - влияющая</param>
        /// <returns>Список генераторов станции</returns>
        public static List<int> GetResGenList(bool isPlantResearched)
        {
            ITable table = _rastr.Tables.Item("ut_node");

            ICol nyColumnItem = table.Cols.Item("ny");
            ICol pgColumnItem = table.Cols.Item("pg");

            List<int> genList = new List<int>();

            for (int index = 0; index < table.Count; index++)
            {
                if ((pgColumnItem.get_ZN(index) > 0) && 
                    (isPlantResearched == true))
                {
                    genList.Add(nyColumnItem.get_ZN(index));
                }
                else if ((pgColumnItem.get_ZN(index) < 0) &&
                    (isPlantResearched == false))
                {
                    genList.Add(nyColumnItem.get_ZN(index));
                }
            }

            return genList;
        }

        /// <summary>
        /// Получение числового значения
        /// </summary>
        /// <param name="tableName">Название таблицы</param>
        /// <param name="parameterName">Название параметра</param>
        /// <param name="number">Номер узла</param>
        /// <param name="chosenParameter">Получаемый параметр</param>
        /// <returns>Числовое значение</returns>
        public static double GetValue(string tableName, string parameterName,
            int number, string chosenParameter)
        {
            ITable table = _rastr.Tables.Item(tableName);
            ICol columnItem = table.Cols.Item(chosenParameter);

            int index = GetIndexByNumber(tableName, parameterName, number);
            return columnItem.get_ZN(index);
        }

        /// <summary>
        /// Получение строкового значения
        /// </summary>
        /// <param name="tableName">Название таблицы</param>
        /// <param name="parameterName">Название параметра</param>
        /// <param name="number">Номер узла</param>
        /// <param name="chosenParameter">Получаемый параметр</param>
        /// <returns>Строковое значение</returns>
        public static string GetStringValue(string tableName,
            string parameterName, int number, string chosenParameter)
        {
            ITable table = _rastr.Tables.Item(tableName);
            ICol columnItem = table.Cols.Item(chosenParameter);

            int index = GetIndexByNumber(tableName, parameterName, number);
            return columnItem.get_ZN(index);
        }

        /// <summary>
        /// Проверка существования установившегося режима
        /// </summary>
        /// <returns>true - существует, false - нет</returns>
        public static bool IsRegimeOK()
        {
            var statusRgm = _rastr.rgm("");

            if (statusRgm == 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Расчет установившегося режима
        /// </summary>
        public static void Regime()
        {
            _rastr.rgm("");
        }

        /// <summary>
        /// Запись значения 
        /// </summary>
        /// <param name="tableName">Наименование таблицы</param>
        /// <param name="parameterName">Наименование параметра</param>
        /// <param name="number">Номер узла</param>
        /// <param name="chosenParameter">Целевой параметр</param>
        /// <param name="value">Новое значение</param>
        public static void SetValue(string tableName, string parameterName,
            int number, string chosenParameter, double value)
        {
            ITable table = _rastr.Tables.Item(tableName);
            ICol columnItem = table.Cols.Item(chosenParameter);

            int index = GetIndexByNumber(tableName, parameterName, number);
            columnItem.set_ZN(index, value);
        }

        /// <summary>
        /// Заполнение списка номерами узлов из файла
        /// </summary>
        /// <param name="tableName">Наименование таблицы</param>
        /// <param name="parameterName">Наименование параметра</param>
        /// <returns>Список значений параметра</returns>
        public static List<int> FillNumbersListFromRastr(string tableName,
            string parameterName)
        {
            List<int> ListOfNumbersFromRastr = new List<int>();

            try
            {
                ITable table = _rastr.Tables.Item(tableName);
                ICol column = table.Cols.Item(parameterName);

                for (int index = 0; index < table.Count; index++)
                {
                    ListOfNumbersFromRastr.Add(column.get_ZN(index));
                }

                return ListOfNumbersFromRastr;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// Шаг назад по траектории
        /// </summary>
        public static void StepBack()
        {
            ITable table = _rastr.Tables.Item("ut_common");
            ICol columnItem = table.Cols.Item("kfc");

            int index = table.Count - 1;
            double step = columnItem.get_ZN(index);

            columnItem.set_Z(index, -step);

            RastrRetCode kd;

            // шаг утяжеления
            kd = _rastr.step_ut("z");
            if (((kd == 0) && (_rastr.ut_Param[ParamUt.UT_ADD_P] == 0))
                || _rastr.ut_Param[ParamUt.UT_TIP] == 1)
            {
                _rastr.AddControl(-1, "");
            }

            columnItem.set_Z(index, step);
        }

        /// <summary>
        /// Добавление новой строки
        /// </summary>
        /// <param name="tableName">Наименование таблицы</param>
        /// <param name="id">Идентификатор в таблице</param>
        /// <param name="number">Номер в таблице</param>
        public static void AddRow(string tableName, int id, int number)
        {
            ITable table = _rastr.Tables.Item(tableName);
            table.AddRow();
            ICol columnItem = table.Cols.Item("Id");
            columnItem.set_ZN(id, number);
        }

        public static void AddKpr(string tableName, int id, int number) 
        {
            ITable table = _rastr.Tables.Item(tableName);
            table.AddRow();
            ICol columnItem = table.Cols.Item("Num");
            columnItem.set_ZN(id, number);
        }

        /// <summary>
        /// Формирование сценариев возмущений I, II группы
        /// </summary>
        /// <param name="shunt">шунт КЗ</param>
        /// <param name="shuntRecloser">Шунт после АПВ</param>
        /// <param name="protectionDelay">Выдержка времени РЗА</param>
        /// <param name="recloserDelay">Выдержка времени АПВ</param>
        /// <param name="faultNode">Узел КЗ</param>
        /// <param name="newFaultNode">Новый узел КЗ</param>
        /// <param name="line1">Линия 1</param>
        /// <param name="line2">Линия 2</param>
        /// <param name="line3">Линия 3</param>
        /// <param name="line4">Линия 4</param>
        public static void MakeScnI(double shunt, double shuntRecloser,
            double protectionDelay, double recloserDelay, int faultNode,
            int newFaultNode, string line1, string line2, string line3,
            string line4)
        {
            #region Формирование действий сценария
            AddRow("DFWAutoActionScn", 0, 1);

            SetValue("DFWAutoActionScn", "Id", 1, "ParentId", 1);
            SetValue("DFWAutoActionScn", "Id", 1, "Type", 6);
            SetStringValue("DFWAutoActionScn", "Id", 1, "Formula",
                shunt.ToString());
            SetValue("DFWAutoActionScn", "Id", 1, "ObjectKey", faultNode);
            SetValue("DFWAutoActionScn", "Id", 1, "OutputMode", 0);
            SetValue("DFWAutoActionScn", "Id", 1, "RunsCount", 1);

            AddRow("DFWAutoActionScn", 1, 2);

            SetValue("DFWAutoActionScn", "Id", 2, "ParentId", 2);
            SetValue("DFWAutoActionScn", "Id", 2, "Type", 3);
            SetValue("DFWAutoActionScn", "Id", 2, "Formula", 0);
            SetStringValue("DFWAutoActionScn", "Id", 2, "ObjectKey", line1);
            SetValue("DFWAutoActionScn", "Id", 2, "OutputMode", 0);
            SetValue("DFWAutoActionScn", "Id", 2, "RunsCount", 1);

            AddRow("DFWAutoActionScn", 2, 3);

            SetValue("DFWAutoActionScn", "Id", 3, "ParentId", 2);
            SetValue("DFWAutoActionScn", "Id", 3, "Type", 3);
            SetValue("DFWAutoActionScn", "Id", 3, "Formula", 0);
            SetStringValue("DFWAutoActionScn", "Id", 3, "ObjectKey", line2);
            SetValue("DFWAutoActionScn", "Id", 3, "OutputMode", 0);
            SetValue("DFWAutoActionScn", "Id", 3, "RunsCount", 1);

            AddRow("DFWAutoActionScn", 3, 4);

            SetValue("DFWAutoActionScn", "Id", 4, "ParentId", 3);
            SetValue("DFWAutoActionScn", "Id", 4, "Type", 5);
            SetValue("DFWAutoActionScn", "Id", 4, "Formula", 0);
            SetValue("DFWAutoActionScn", "Id", 4, "ObjectKey", newFaultNode);
            SetValue("DFWAutoActionScn", "Id", 4, "OutputMode", 0);
            SetValue("DFWAutoActionScn", "Id", 4, "RunsCount", 1);

            AddRow("DFWAutoActionScn", 4, 5);

            SetValue("DFWAutoActionScn", "Id", 5, "ParentId", 4);
            SetValue("DFWAutoActionScn", "Id", 5, "Type", 6);
            SetStringValue("DFWAutoActionScn", "Id", 5, "Formula",
                shuntRecloser.ToString());
            SetValue("DFWAutoActionScn", "Id", 5, "ObjectKey", newFaultNode);
            SetValue("DFWAutoActionScn", "Id", 5, "OutputMode", 0);
            SetValue("DFWAutoActionScn", "Id", 5, "RunsCount", 1);

            AddRow("DFWAutoActionScn", 5, 6);

            SetValue("DFWAutoActionScn", "Id", 6, "ParentId", 5);
            SetValue("DFWAutoActionScn", "Id", 6, "Type", 3);
            SetValue("DFWAutoActionScn", "Id", 6, "Formula", 0);
            SetStringValue("DFWAutoActionScn", "Id", 6, "ObjectKey",
                line3);
            SetValue("DFWAutoActionScn", "Id", 6, "OutputMode", 0);
            SetValue("DFWAutoActionScn", "Id", 6, "RunsCount", 1);

            AddRow("DFWAutoActionScn", 6, 7);

            SetValue("DFWAutoActionScn", "Id", 7, "ParentId", 5);
            SetValue("DFWAutoActionScn", "Id", 7, "Type", 3);
            SetValue("DFWAutoActionScn", "Id", 7, "Formula", 0);
            SetStringValue("DFWAutoActionScn", "Id", 7, "ObjectKey",
                line4);
            SetValue("DFWAutoActionScn", "Id", 7, "OutputMode", 0);
            SetValue("DFWAutoActionScn", "Id", 7, "RunsCount", 1);

            AddRow("DFWAutoActionScn", 7, 8);

            SetValue("DFWAutoActionScn", "Id", 8, "ParentId", 6);
            SetValue("DFWAutoActionScn", "Id", 8, "Type", 3);
            SetValue("DFWAutoActionScn", "Id", 8, "Formula", 1);
            SetStringValue("DFWAutoActionScn", "Id", 8, "ObjectKey", line3);
            SetValue("DFWAutoActionScn", "Id", 8, "OutputMode", 0);
            SetValue("DFWAutoActionScn", "Id", 8, "RunsCount", 1);

            AddRow("DFWAutoActionScn", 8, 9);

            SetValue("DFWAutoActionScn", "Id", 9, "ParentId", 6);
            SetValue("DFWAutoActionScn", "Id", 9, "Type", 3);
            SetValue("DFWAutoActionScn", "Id", 9, "Formula", 1);
            SetStringValue("DFWAutoActionScn", "Id", 9, "ObjectKey", line4);
            SetValue("DFWAutoActionScn", "Id", 9, "OutputMode", 0);
            SetValue("DFWAutoActionScn", "Id", 9, "RunsCount", 1);

            AddRow("DFWAutoActionScn", 9, 10);

            SetValue("DFWAutoActionScn", "Id", 10, "ParentId", 7);
            SetValue("DFWAutoActionScn", "Id", 10, "Type", 3);
            SetValue("DFWAutoActionScn", "Id", 10, "Formula", 0);
            SetStringValue("DFWAutoActionScn", "Id", 10, "ObjectKey", line3);
            SetValue("DFWAutoActionScn", "Id", 10, "OutputMode", 0);
            SetValue("DFWAutoActionScn", "Id", 10, "RunsCount", 1);

            AddRow("DFWAutoActionScn", 10, 11);

            SetValue("DFWAutoActionScn", "Id", 11, "ParentId", 7);
            SetValue("DFWAutoActionScn", "Id", 11, "Type", 3);
            SetValue("DFWAutoActionScn", "Id", 11, "Formula", 0);
            SetStringValue("DFWAutoActionScn", "Id", 11, "ObjectKey", line4);
            SetValue("DFWAutoActionScn", "Id", 11, "OutputMode", 0);
            SetValue("DFWAutoActionScn", "Id", 11, "RunsCount", 1);
            #endregion

            #region Формирование логики сценария
            AddRow("DFWAutoLogicScn", 0, 1);

            SetValue("DFWAutoLogicScn", "Id", 1, "ParentId", 1);
            SetValue("DFWAutoLogicScn", "Id", 1, "Type", 1);
            SetStringValue("DFWAutoLogicScn", "Id", 1, "Actions", "A1");
            SetValue("DFWAutoLogicScn", "Id", 1, "Delay", 0);
            SetValue("DFWAutoLogicScn", "Id", 1, "OutputMode", 0);

            AddRow("DFWAutoLogicScn", 1, 2);

            SetValue("DFWAutoLogicScn", "Id", 2, "ParentId", 2);
            SetValue("DFWAutoLogicScn", "Id", 2, "Type", 1);
            SetStringValue("DFWAutoLogicScn", "Id", 2, "Actions", "A2");
            SetStringValue("DFWAutoLogicScn", "Id", 2, "Delay",
                protectionDelay.ToString());
            SetValue("DFWAutoLogicScn", "Id", 2, "OutputMode", 0);

            AddRow("DFWAutoLogicScn", 2, 3);

            SetValue("DFWAutoLogicScn", "Id", 3, "ParentId", 3);
            SetValue("DFWAutoLogicScn", "Id", 3, "Type", 1);
            SetStringValue("DFWAutoLogicScn", "Id", 3, "Actions", "A3");
            SetStringValue("DFWAutoLogicScn", "Id", 3, "Delay",
                protectionDelay.ToString());
            SetValue("DFWAutoActionScn", "Id", 3, "OutputMode", 0);

            AddRow("DFWAutoLogicScn", 3, 4);

            SetValue("DFWAutoLogicScn", "Id", 4, "ParentId", 4);
            SetValue("DFWAutoLogicScn", "Id", 4, "Type", 1);
            SetStringValue("DFWAutoLogicScn", "Id", 4, "Actions", "A4");
            SetStringValue("DFWAutoLogicScn", "Id", 4, "Delay",
                protectionDelay.ToString());
            SetValue("DFWAutoActionScn", "Id", 4, "OutputMode", 0);

            AddRow("DFWAutoLogicScn", 4, 5);

            SetValue("DFWAutoLogicScn", "Id", 5, "ParentId", 5);
            SetValue("DFWAutoLogicScn", "Id", 5, "Type", 1);
            SetStringValue("DFWAutoLogicScn", "Id", 5, "Actions", "A5");
            SetStringValue("DFWAutoLogicScn", "Id", 5, "Delay",
                (protectionDelay + 0.05).ToString());
            SetValue("DFWAutoLogicScn", "Id", 5, "OutputMode", 0);

            AddRow("DFWAutoLogicScn", 5, 6);

            SetValue("DFWAutoLogicScn", "Id", 6, "ParentId", 6);
            SetValue("DFWAutoLogicScn", "Id", 6, "Type", 1);
            SetStringValue("DFWAutoLogicScn", "Id", 6, "Actions", "A6");
            SetStringValue("DFWAutoLogicScn", "Id", 6, "Delay",
                (protectionDelay + 0.05 + recloserDelay).ToString());
            SetValue("DFWAutoLogicScn", "Id", 6, "OutputMode", 0);

            AddRow("DFWAutoLogicScn", 6, 7);

            SetValue("DFWAutoLogicScn", "Id", 7, "ParentId", 7);
            SetValue("DFWAutoLogicScn", "Id", 7, "Type", 1);
            SetStringValue("DFWAutoLogicScn", "Id", 7, "Actions", "A7");
            SetStringValue("DFWAutoLogicScn", "Id", 7, "Delay",
                (protectionDelay + 0.05 +
                recloserDelay + protectionDelay).ToString());
            SetValue("DFWAutoLogicScn", "Id", 7, "OutputMode", 0);
            #endregion
        }

        /// <summary>
        /// Формирование сценариев возмущений III группы
        /// </summary>
        /// <param name="shunt">шунт КЗ</param>
        /// <param name="shuntRecloser">Шунт после АПВ</param>
        /// <param name="protectionDelay">Выдержка времени РЗА</param>
        /// <param name="recloserDelay">Выдержка времени АПВ</param>
        /// <param name="faultNode">Узел КЗ</param>
        /// <param name="newFaultNode">Новый узел КЗ</param>
        /// <param name="line1">Линия 1</param>
        /// <param name="line2">Линия 2</param>
        /// <param name="line3">Линия 3</param>
        /// <param name="line4">Линия 4</param>
        public static void MakeScnIII(double shunt, double protectionDelay,
            double CBFPDelay, int faultNode, string line1, string line2,
            string line3, string line4, string line5)
        {
            #region Формирование действий сценария
            AddRow("DFWAutoActionScn", 0, 1);

            SetValue("DFWAutoActionScn", "Id", 1, "ParentId", 1);
            SetValue("DFWAutoActionScn", "Id", 1, "Type", 6);
            SetStringValue("DFWAutoActionScn", "Id", 1, "Formula",
                shunt.ToString());
            SetValue("DFWAutoActionScn", "Id", 1, "ObjectKey",
                faultNode);
            SetValue("DFWAutoActionScn", "Id", 1, "OutputMode", 0);
            SetValue("DFWAutoActionScn", "Id", 1, "RunsCount", 1);

            AddRow("DFWAutoActionScn", 1, 2);

            SetValue("DFWAutoActionScn", "Id", 2, "ParentId", 2);
            SetValue("DFWAutoActionScn", "Id", 2, "Type", 3);
            SetValue("DFWAutoActionScn", "Id", 2, "Formula", 0);
            SetStringValue("DFWAutoActionScn", "Id", 2, "ObjectKey", line1);
            SetValue("DFWAutoActionScn", "Id", 2, "OutputMode", 0);
            SetValue("DFWAutoActionScn", "Id", 2, "RunsCount", 1);

            AddRow("DFWAutoActionScn", 2, 3);

            SetValue("DFWAutoActionScn", "Id", 3, "ParentId", 3);
            SetValue("DFWAutoActionScn", "Id", 3, "Type", 3);
            SetValue("DFWAutoActionScn", "Id", 3, "Formula", 0);
            SetStringValue("DFWAutoActionScn", "Id", 3, "ObjectKey",
                line2);
            SetValue("DFWAutoActionScn", "Id", 3, "OutputMode", 0);
            SetValue("DFWAutoActionScn", "Id", 3, "RunsCount", 1);

            AddRow("DFWAutoActionScn", 3, 4);

            SetValue("DFWAutoActionScn", "Id", 4, "ParentId", 3);
            SetValue("DFWAutoActionScn", "Id", 4, "Type", 3);
            SetValue("DFWAutoActionScn", "Id", 4, "Formula", 0);
            SetStringValue("DFWAutoActionScn", "Id", 4, "ObjectKey",
                line3);
            SetValue("DFWAutoActionScn", "Id", 4, "OutputMode", 0);
            SetValue("DFWAutoActionScn", "Id", 4, "RunsCount", 1);

            AddRow("DFWAutoActionScn", 4, 5);

            SetValue("DFWAutoActionScn", "Id", 5, "ParentId", 4);
            SetValue("DFWAutoActionScn", "Id", 5, "Type", 3);
            SetValue("DFWAutoActionScn", "Id", 5, "Formula", 0);
            SetStringValue("DFWAutoActionScn", "Id", 5, "ObjectKey",
                line4);
            SetValue("DFWAutoActionScn", "Id", 5, "OutputMode", 0);
            SetValue("DFWAutoActionScn", "Id", 5, "RunsCount", 1);

            AddRow("DFWAutoActionScn", 5, 6);

            SetValue("DFWAutoActionScn", "Id", 6, "ParentId", 4);
            SetValue("DFWAutoActionScn", "Id", 6, "Type", 3);
            SetValue("DFWAutoActionScn", "Id", 6, "Formula", 0);
            SetStringValue("DFWAutoActionScn", "Id", 6, "ObjectKey",
                line5);
            SetValue("DFWAutoActionScn", "Id", 6, "OutputMode", 0);
            SetValue("DFWAutoActionScn", "Id", 6, "RunsCount", 1);
            #endregion

            #region Формирование логики сценария
            AddRow("DFWAutoLogicScn", 0, 1);

            SetValue("DFWAutoLogicScn", "Id", 1, "Formula", 1);
            SetValue("DFWAutoLogicScn", "Id", 1, "Type", 1);
            SetStringValue("DFWAutoLogicScn", "Id", 1, "Actions", "A1");
            SetValue("DFWAutoLogicScn", "Id", 1, "Delay", 0);
            SetValue("DFWAutoLogicScn", "Id", 1, "OutputMode", 0);

            AddRow("DFWAutoLogicScn", 1, 2);

            SetValue("DFWAutoLogicScn", "Id", 2, "Formula", 1);
            SetValue("DFWAutoLogicScn", "Id", 2, "Type", 1);
            SetStringValue("DFWAutoLogicScn", "Id", 2, "Actions", "A2");
            SetStringValue("DFWAutoLogicScn", "Id", 2, "Delay",
                protectionDelay.ToString());
           SetValue("DFWAutoLogicScn", "Id", 2, "OutputMode", 0);

            AddRow("DFWAutoLogicScn", 2, 3);

            SetValue("DFWAutoLogicScn", "Id", 3, "Formula", 1);
            SetValue("DFWAutoLogicScn", "Id", 3, "Type", 1);
            SetStringValue("DFWAutoLogicScn", "Id", 3, "Actions", "A3");
            SetStringValue("DFWAutoLogicScn", "Id", 3, "Delay",
                (protectionDelay + 0.05).ToString());
            SetValue("DFWAutoActionScn", "Id", 3, "OutputMode", 0);

            AddRow("DFWAutoLogicScn", 3, 4);

            SetValue("DFWAutoLogicScn", "Id", 4, "Formula", 1);
            SetValue("DFWAutoLogicScn", "Id", 4, "Type", 1);
            SetStringValue("DFWAutoLogicScn", "Id", 4, "Actions", "A4");
            SetStringValue("DFWAutoLogicScn", "Id", 4, "Delay",
                (protectionDelay + 0.05 + CBFPDelay).ToString());
           SetValue("DFWAutoActionScn", "Id", 4, "OutputMode", 0);
            #endregion
        }

        /// <summary>
        /// Запуск расчета динамики
        /// </summary>
        /// <param name="inputWidth">Размер окна расчета</param>
        public static void CalculateDynamic(double inputWidth,
            string savePath) 
        {
            Console.WriteLine("Начался расчет динамики");

            try
            {
                FWDynamic FWDynamic = _rastr.FWDynamic();

                ITable table = _rastr.Tables.Item("com_dynamics");

                ICol columnItemsTras = table.Cols.Item("Tras");
                columnItemsTras.set_ZN(0, inputWidth);

                FWDynamic.Run();
                //FWDynamic.RunEMSmode();

                Console.WriteLine("Расчет динамики завершен");
            }
            catch (Exception ex) 
            {
                throw new Exception(ex.Message);
            }
            
        }

        /// <summary>
        /// Подготовка файла для расчета переходного процесса
        /// </summary>
        /// <param name="rstFile">Файл режима</param>
        /// <param name="scnFile">Файл сценария</param>
        /// <param name="dfwfile">Файл автоматики</param>
        /// <param name="shablScn">Шаблон сценария</param>
        /// <param name="kprFile">Файл КВ</param>
        /// <param name="name">Наименование целевого генератора</param>
        /// <param name="set">Выборка целевого генератора</param>
        /// <param name="nameSupport">Наименование опорного генератора</param>
        /// <param name="setSupport">Выборка опорного генератора</param>
        public static void FilePrepair(string rstFile, string scnFile,
            string dfwfile, string shablScn, string kprFile, string name,
            string set, string nameSupport, string setSupport) 
        {
            LoadFile(rstFile, "");
            LoadFile(scnFile, shablScn);
            LoadFile(dfwfile, dfwfile);
            LoadFile(kprFile, kprFile);
            CreateKpr(name, set, nameSupport, setSupport);
        }

        /// <summary>
        /// Создание файлов КВ
        /// </summary>
        /// <param name="name">Наименование целевого генератора</param>
        /// <param name="set">Выборка целевого генератора</param>
        /// <param name="nameSupport">Наименование опорного генератора</param>
        /// <param name="setSupport">Выборка опорного генератора</param>
        public static void CreateKpr(string name, string set, 
            string nameSupport, string setSupport) 
        {
            #region Формирование параметров целевого генератора
            AddKpr("ots_val", 0, 1);

            SetStringValue("ots_val", "Num", 1, "name", name);
            SetValue("ots_val", "Num", 1, "tip", 0);
            SetStringValue("ots_val", "Num", 1, "tabl", "Generator");
            SetStringValue("ots_val", "Num", 1, "vibork", set);
            SetStringValue("ots_val", "Num", 1, "formula", "Delta");
            SetValue("ots_val", "Num", 1, "prec", 2);
            SetValue("ots_val", "Num", 1, "mash", 57);
            #endregion

            #region Формирование параметров опорного генератора
            AddKpr("ots_val", 1, 2);

            SetStringValue("ots_val", "Num", 2, "name", nameSupport);
            SetValue("ots_val", "Num", 2, "tip", 0);
            SetStringValue("ots_val", "Num", 2, "tabl", "Generator");
            SetStringValue("ots_val", "Num", 2, "vibork", setSupport);
            SetStringValue("ots_val", "Num", 2, "formula", "Delta");
            SetValue("ots_val", "Num", 2, "prec", 2);
            SetValue("ots_val", "Num", 2, "mash", 57);
            #endregion
        }

        /// <summary>
        /// Запись результата в CSV
        /// </summary>
        /// <param name="name">Наименование таблицы</param>
        /// <param name="parameter">Наименование параметра</param>
        /// <param name="key">Номер гененератора</param>
        /// <param name="snapIndex">Номер файла результата расчета</param>
        /// <param name="station">Наименование станции</param>
        /// <param name="index">Порядковый номер результата</param>
        public static void GetTransient(string name, string parameter,
            int key, int snapIndex, string station, int index)
        {
            int genId = GetIndexByNumber(name, "Num", key);
            var plot = _rastr.GetChainedGraphSnapshot(name, parameter,
                genId, snapIndex);

            string savePath = $"C:\\Users\\geu2\\Desktop\\res\\{station}_{index}.csv";

            using (var writer = new StreamWriter(savePath))
            using (var csv = new CsvWriter(writer,
                new CsvHelper.Configuration.CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = ";"
            }))
            {
                csv.WriteField("delta");
                csv.WriteField("t");
                csv.NextRecord();

                for (int i = 0; i < plot.GetLength(0); i++)
                {
                    for (int j = 0; j < plot.GetLength(1); j++)
                    {
                        csv.WriteField(plot[i, j].
                            ToString("F",
                            CultureInfo.CreateSpecificCulture("ru-RU")));
                    }

                    csv.NextRecord();
                }
            }

            Console.WriteLine($"Данные записаны в файл {savePath}");
        }
    }
}