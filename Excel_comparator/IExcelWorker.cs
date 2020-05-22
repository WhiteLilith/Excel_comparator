using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_comparator
{
    interface IExcelWorker
    {
        /// <summary>
        /// Загрузка фйайлов для сравнения
        /// </summary>
        /// <param name="file_1"></param>
        /// <param name="file_2"></param>
        void GetFiles(string file_1, string file_2);

        /// <summary>
        /// Закрытие всех файлов
        /// </summary>
        void CloseFiles();

        /// <summary>
        /// новые люди в файле 2
        /// </summary>
        /// <returns></returns>
        List<string> NewPeople();

        /// <summary>
        /// Асинхронный аналог метода для поиска новых людей
        /// </summary>
        /// <returns></returns>
        Task<List<string>> NewPeopleAsync();

        /// <summary>
        /// отсутствующие люди в файле 2
        /// </summary>
        /// <returns></returns>
        List<string> MissingPeople();

        /// <summary>
        /// Асинхронный аналог метода для поиска отсутствующих людей
        /// </summary>
        /// <returns></returns>
        Task<List<string>> MissingPeopleAsync();
    }
}
