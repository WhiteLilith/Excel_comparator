using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_comparator
{

    class ExcelWorker : IExcelWorker
    {

        private Microsoft.Office.Interop.Excel.Worksheet workSheet_1;
        private Microsoft.Office.Interop.Excel.Worksheet workSheet_2;

        private Excel._Application excelApp_1;
        private Excel._Application excelApp_2;

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

            // Сделать Excel невидимым (необязательно).
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
            int num = 0;
            for (int i = 1; i < lengthOfSheet + 1; i++)
            {
                if (workSheet.Cells[1, i].ToString().ToLower() == rowName)
                {
                    num = i;
                    break;
                }
            }
            return num;
        }

        public List<string> MissingPeople()
        {
            List<Human> firstSheetPeople = new List<Human>();
            List<Human> secondSheetPeople = new List<Human>();

            int lastRow_1 = GetLastRow(workSheet_1);
            int lastRow_2 = GetLastRow(workSheet_2);

            int lengthOfFirst = GetLastColumn(workSheet_2);
            int lengthOfSecond = GetLastColumn(workSheet_2);


            int firstSurenameNumber = GetRow(lengthOfFirst, "фамилия", workSheet_1);
            int firstNameNumber = GetRow(lengthOfFirst, "имя", workSheet_1);
            int firstPatronymicNumber = GetRow(lengthOfFirst, "отчество", workSheet_1);

            int secondSurenameNumber = GetRow(lengthOfSecond, "фамилия", workSheet_2);
            int secondNameNumber = GetRow(lengthOfSecond, "имя", workSheet_2);
            int secondPatronymicNumber = GetRow(lengthOfSecond, "отчество", workSheet_2);


            for (int i = 2; i < lastRow_1 + 1; i++)
            {
                string surename = workSheet_1.Cells[i, firstSurenameNumber].ToString();
                string name = workSheet_1.Cells[i, firstNameNumber].ToString();
                string patronymic = workSheet_1.Cells[i, firstPatronymicNumber].ToString();
                Human currentHuman = new Human(name, surename, patronymic);
                firstSheetPeople.Add(currentHuman);
            }

            for (int i = 2; i < lastRow_2 + 1; i++)
            {
                string surename = workSheet_2.Cells[i, secondSurenameNumber].ToString();
                string name = workSheet_2.Cells[i, secondNameNumber].ToString();
                string patronymic = workSheet_2.Cells[i, secondPatronymicNumber].ToString();
                Human currentHuman = new Human(name, surename, patronymic);
                secondSheetPeople.Add(currentHuman);
            }

            List<Human> missingPeople = new List<Human>();

            bool exist = false;
            for (int i = 0; i < lastRow_1; i++)
            {
                exist = false;
                for (int j = 0; j < lastRow_2; j++)
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

        public List<string> NewPeople()
        {
            List<Human> firstSheetPeople = new List<Human>();
            List<Human> secondSheetPeople = new List<Human>();

            int lastRow_1 = GetLastRow(workSheet_1);
            int lastRow_2 = GetLastRow(workSheet_2);

            int lengthOfFirst = GetLastColumn(workSheet_2);
            int lengthOfSecond = GetLastColumn(workSheet_2);


            int firstSurenameNumber = GetRow(lengthOfFirst, "фамилия", workSheet_1);
            int firstNameNumber = GetRow(lengthOfFirst, "имя", workSheet_1);
            int firstPatronymicNumber = GetRow(lengthOfFirst, "отчество", workSheet_1);

            int secondSurenameNumber = GetRow(lengthOfSecond, "фамилия", workSheet_2);
            int secondNameNumber = GetRow(lengthOfSecond, "имя", workSheet_2);
            int secondPatronymicNumber = GetRow(lengthOfSecond, "отчество", workSheet_2);


            for (int i = 2; i < lastRow_1 + 1; i++)
            {
                string surename = workSheet_1.Cells[i, firstSurenameNumber].ToString();
                string name = workSheet_1.Cells[i, firstNameNumber].ToString();
                string patronymic = workSheet_1.Cells[i, firstPatronymicNumber].ToString();
                Human currentHuman = new Human(name, surename, patronymic);
                firstSheetPeople.Add(currentHuman);
            }

            for (int i = 2; i < lastRow_2 + 1; i++)
            {
                string surename = workSheet_2.Cells[i, secondSurenameNumber].ToString();
                string name = workSheet_2.Cells[i, secondNameNumber].ToString();
                string patronymic = workSheet_2.Cells[i, secondPatronymicNumber].ToString();
                Human currentHuman = new Human(name, surename, patronymic);
                secondSheetPeople.Add(currentHuman);
            }

            List<Human> missingPeople = new List<Human>();

            bool exist = false;
            for (int i = 0; i < lastRow_2; i++)
            {
                exist = false;
                for (int j = 0; j < lastRow_1; j++)
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
    }
}

