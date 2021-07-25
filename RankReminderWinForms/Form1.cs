using System;
using System.Windows.Forms;

namespace RankReminderWinForms
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form2 newForm = new Form2();
            newForm.Show(); // показать главное окно программы
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            Application.Exit(); // выход из приложения
        }

        private void button3_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Кадровый калькулятор. Версия: 1.0.\nПрограмма для автоматического подсчета даты присвоения очередных специальных званий сотрудников.\nАвтор программы: Банков Михаил.");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form3 newForm2 = new Form3();
            newForm2.Show(); // показать главное окно программы
        }
    }
}
