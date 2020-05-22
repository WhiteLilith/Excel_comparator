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
            ofd.Filter = "excel файл (*.xlsx)| *.xlsx";

            ofd.ShowDialog();

            return ofd.FileName;
        }

        /// <summary>
        /// запуск сравнения
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            string path1 = textFilePath1.Text;
            string path2 = textFilePath2.Text;

            if (!File.Exists(path1))
            {
                MessageBox.Show("Задан несуществующий первый файл", "Ошибка");
                return;
            }
            if (!File.Exists(path2))
            {
                MessageBox.Show("Задан несуществующий второй файл", "Ошибка");
                return;
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
    }
}
