using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_comparator
{
    /*
     * 
     * Только я и бог знаем, как это работает 
     * Через месяц только бог будет знать как он работает
     * 
     */

    class ExcelWorker : IExcelWorker
    {

        private Excel.Worksheet workSheet_1;
        private Excel.Worksheet workSheet_2;

        public Excel._Application excelApp_1;
        public Excel._Application excelApp_2;

        private Excel.Workbook workBook_1;
        private Excel.Workbook workBook_2;

        public void GetFiles(string file_1, string file_2)
        {
            // Получить объект приложения Excel.
            excelApp_1 = new Excel.ApplicationClass();

            // Сделать Excel невидимым (необязательно).
            excelApp_1.Visible = false;

            // Откройте рабочую книгу только для чтения.
            workBook_1 = excelApp_1.Workbooks.Open(
                file_1,
                Type.Missing, true, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            workSheet_1 = (Microsoft.Office.Interop.Excel.Worksheet)workBook_1.Sheets[1];

            excelApp_2 = new Excel.ApplicationClass();

            excelApp_2.Visible = false;

            workBook_2 = excelApp_2.Workbooks.Open(
                file_2,
                Type.Missing, true, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            workSheet_2 = (Microsoft.Office.Interop.Excel.Worksheet)workBook_2.Sheets[1];
        }

        /// <summary>
        /// Закрывает открытые excel файлы
        /// </summary>
        public void CloseFiles()
        {
            workBook_1.Close(false, Type.Missing, Type.Missing);
            workBook_2.Close(false, Type.Missing, Type.Missing);
            excelApp_1.Quit();
            excelApp_2.Quit();
        }

        /// <summary>
        /// возвращает длину строки
        /// </summary>
        /// <param name="workSheet"></param>
        /// <returns></returns>
        private int GetLastRow(Excel.Worksheet workSheet)
        {
            int counter = 0;
            int lastRow = 0;
            while (counter < 2500) //таки костыль
            {
                try
                {
                    lastRow = workSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                    Thread.Sleep(250);
                    return lastRow;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    counter++;
                }
            }

            throw new IndexOutOfRangeException();
        }

        /// <summary>
        /// Возвращает высоту столбца
        /// </summary>
        /// <param name="workSheet"></param>
        /// <returns></returns>
        private int GetLastColumn(Excel.Worksheet workSheet)
        {
            int counter = 0;
            int lastRow = 0;
            while (counter < 2500) 
            {
                try
                {
                    lastRow = workSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;
                    Thread.Sleep(250);
                    return lastRow;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    counter++;
                }
            }

            throw new IndexOutOfRangeException();
        }

        /// <summary>
        /// возвращает номер столбца с нужным именем
        /// </summary>
        /// <param name="lengthOfSheet"></param>
        /// <param name="rowName"></param>
        /// <param name="workSheet"></param>
        /// <returns></returns>
        private int GetRow(int lengthOfSheet, string rowName, Excel.Worksheet workSheet)
        {
            Excel.Range xlRange = workSheet.UsedRange;

            int num = 0;
            for (int i = 1; i < lengthOfSheet + 1; i++)
            {
                if (((Excel.Range)xlRange.Cells[1, i]).Value2 != null && ((Excel.Range)xlRange.Cells[1, i]).Value2.ToString().ToLower() == rowName.ToLower())
                {
                    num = i;
                    break;
                }
            }
            return num;
        }

        /// <summary>
        /// заполняет список людей по их ФИО
        /// </summary>
        /// <param name="listPeople"></param>
        /// <param name="workSheet"></param>
        /// <param name="lastRow"></param>
        /// <param name="surenameNum"></param>
        /// <param name="nameNum"></param>
        /// <param name="patronymicNum"></param>
        private void FillSNPList(ref List<Human> listPeople, Excel.Worksheet workSheet, int lastRow, int surenameNum, int nameNum, int patronymicNum)
        {
            string surename = "";
            string name = "";
            string patronymic = "";

            Excel.Range xlRange = workSheet.UsedRange;

            for (int i = 2; i < lastRow + 1; i++)
            {
                surename = "";
                name = "";
                patronymic = "";
                try
                {
                    surename = ((Excel.Range)xlRange.Cells[i, surenameNum]).Value2.ToString().ToLower();
                    name = ((Excel.Range)xlRange.Cells[i, nameNum]).Value2.ToString().ToLower();
                    patronymic = ((Excel.Range)xlRange.Cells[i, patronymicNum]).Value2.ToString().ToLower();
                }
                catch (System.NullReferenceException)
                {

                }

                if (name != "" && surename != "")
                {
                    Human currentHuman = new Human(name, surename, patronymic);
                    listPeople.Add(currentHuman);
                }
                else
                {
                    break;
                }
            }
        }

        /// <summary>
        /// заполняет списки людей из двух книг
        /// </summary>
        /// <param name="firstList"></param>
        /// <param name="secondList"></param>
        private void FillListsOfPeople(ref List<Human> firstList, ref List<Human> secondList, int lastRow_1, int lastRow_2)
        {
            int lengthOfFirst = GetLastColumn(workSheet_1);
            int lengthOfSecond = GetLastColumn(workSheet_2);

            int firstSurenameNumber = GetRow(lengthOfFirst, "фамилия", workSheet_1);
            int firstNameNumber = GetRow(lengthOfFirst, "имя", workSheet_1);
            int firstPatronymicNumber = GetRow(lengthOfFirst, "отчество", workSheet_1);

            int secondSurenameNumber = GetRow(lengthOfSecond, "фамилия", workSheet_2);
            int secondNameNumber = GetRow(lengthOfSecond, "имя", workSheet_2);
            int secondPatronymicNumber = GetRow(lengthOfSecond, "отчество", workSheet_2);

            FillSNPList(ref firstList, workSheet_1, lastRow_1, firstSurenameNumber, firstNameNumber, firstPatronymicNumber);

            FillSNPList(ref secondList, workSheet_2, lastRow_2, secondSurenameNumber, secondNameNumber, secondPatronymicNumber);
        }

        public List<string> MissingPeople()
        {
            List<Human> firstSheetPeople = new List<Human>();
            List<Human> secondSheetPeople = new List<Human>();

            int lastRow_1 = GetLastRow(workSheet_1);
            int lastRow_2 = GetLastRow(workSheet_2);

            FillListsOfPeople(ref firstSheetPeople, ref secondSheetPeople, lastRow_1, lastRow_2);

            List<Human> missingPeople = new List<Human>();

            bool exist = false;
            for (int i = 0; i < firstSheetPeople.Count; i++)
            {
                exist = false;
                for (int j = 0; j < secondSheetPeople.Count; j++)
                {
                    if (firstSheetPeople[i] == secondSheetPeople[j])
                    {
                        exist = true;
                    }
                }

                if (!exist)
                {
                    missingPeople.Add(firstSheetPeople[i]);
                }
            }

            int amountOfMissing = missingPeople.Count;

            List<string> returnListOfPeople = new List<string>();

            for (int i = 0; i < amountOfMissing; i++)
            {
                returnListOfPeople.Add(missingPeople[i].GetFullName());
            }

            return returnListOfPeople;
        }

        public async Task<List<string>> MissingPeopleAsync()
        {
            return await Task.Run(() => MissingPeople());
        }

        public List<string> NewPeople()
        {
            List<Human> firstSheetPeople = new List<Human>();
            List<Human> secondSheetPeople = new List<Human>();

            int lastRow_1 = GetLastRow(workSheet_1);
            int lastRow_2 = GetLastRow(workSheet_2);

            FillListsOfPeople(ref firstSheetPeople, ref secondSheetPeople, lastRow_1, lastRow_2);

            List<Human> missingPeople = new List<Human>();

            bool exist = false;
            for (int i = 0; i < secondSheetPeople.Count; i++)
            {
                exist = false;
                for (int j = 0; j < firstSheetPeople.Count; j++)
                {
                    if (firstSheetPeople[j] == secondSheetPeople[i])
                    {
                        exist = true;
                    }
                }

                if (!exist)
                {
                    missingPeople.Add(firstSheetPeople[i]);
                }
            }

            int amountOfMissing = missingPeople.Count;

            List<string> returnListOfPeople = new List<string>();

            for (int i = 0; i < amountOfMissing; i++)
            {
                returnListOfPeople.Add(missingPeople[i].GetFullName());
            }

            return returnListOfPeople;
        }

        public async Task<List<string>> NewPeopleAsync()
        {
            return await Task.Run(() => MissingPeople());
        }
    }
}

