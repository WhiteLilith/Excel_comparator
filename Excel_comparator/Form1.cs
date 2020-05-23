using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Globalization;

namespace Excel_comparator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private string FileDialog()
        {
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.Title = "Выберете файл";
            ofd.Filter = "excel файл (*.xlsx, *.xls)| *.xlsx; *.xls";

            ofd.ShowDialog();

            return ofd.FileName;
        }

        /// <summary>
        /// запуск сравнения
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void button1_Click(object sender, EventArgs e)
        {
            listNewPeople.Items.Clear();
            listMissingPeople.Items.Clear();

            buttonCompare.Enabled = false;
            string path1 = textFilePath1.Text;
            string path2 = textFilePath2.Text;

            if (!File.Exists(path1))
            {
                MessageBox.Show("Задан несуществующий первый файл", "Ошибка");
                buttonCompare.Enabled = true;
                return;
            }
            if (!File.Exists(path2))
            {
                MessageBox.Show("Задан несуществующий второй файл", "Ошибка");
                buttonCompare.Enabled = true;
                return;
            }
            ExcelWorker ew = new ExcelWorker();
            ew.GetFiles(path1, path2);

            string[] newPeople = { };
            string[] missingPeople = { };

            try
            {
                List<string> newPeopleList = await ew.NewPeopleAsync();
                List<string> missingPeopleList = await ew.MissingPeopleAsync();

                var textInfo = new CultureInfo("ru-RU").TextInfo;

                newPeople = newPeopleList.ToArray();
                for(int i = 0; i < newPeople.Length; i++)
                {
                    newPeople[i] = textInfo.ToTitleCase(newPeople[i]);
                }
                listNewPeople.Items.AddRange(newPeople);

                missingPeople = missingPeopleList.ToArray();
                for (int i = 0; i < missingPeople.Length; i++)
                {
                    missingPeople[i] = textInfo.ToTitleCase(missingPeople[i]);
                }
                listMissingPeople.Items.AddRange(missingPeople);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show("Произошла неизвестная ошибка. Попробуйте еще раз", "Ошибка");
                buttonCompare.Enabled = true;
            }
            finally 
            {
                ew.CloseFiles();
                Thread.Sleep(50);
                buttonCompare.Enabled = true;
            }
        }

        /// <summary>
        /// выбор 1 файла из выпадающего списка
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ChooseFile1_Click(object sender, EventArgs e)
        {
            textFilePath1.Text = FileDialog();
        }

        /// <summary>
        /// выбор 2 файла из выпадающего списка
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ChooseFile2_Click(object sender, EventArgs e)
        {
            textFilePath2.Text = FileDialog();
        }

        // это случайно добавилось, но без него ничего не работает ¯\_(ツ)_/¯
        private void labelFile2_Click(object sender, EventArgs e)
        {

        }

        private void labelNewPeople_Click(object sender, EventArgs e)
        {

        }
    }
}
