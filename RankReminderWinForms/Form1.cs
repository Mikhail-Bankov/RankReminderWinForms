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
            newForm.ShowDialog(); // показать главное окно программы
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form3 newForm2 = new Form3();
            newForm2.Show(); // показать главное окно программы
        }

        private void button3_Click(object sender, EventArgs e)
        {
            MessageBox.Show("АИС \"Кадры ГФС\". Версия: 1.0.\n\nПрограмма для учета сведений о сотрудниках\nГосударственной фельдъегерской службы\nи автоматизации их обработки.\n\nАвтор: Банков Михаил aka PC_USER\nE-mail: pcuser@internet.ru\nGithub: https://github.com/Mikhail-Bankov", "О программе",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            Application.Exit(); // выход из приложения
        }
    }
}