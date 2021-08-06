using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.IO;

namespace RankReminderWinForms
{
    public partial class Form3 : Form
    {
        DataSet dataSet1 = new DataSet();
        public Form3()
        {
            InitializeComponent();            
        }

        private void Button_CloseSettings_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Button_UnloadDB_Click(object sender, EventArgs e)
        {            
            String date = DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year;
            String fileName = "BaseLichSost_Backup (" + date + ")";
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Title = "Выберите место для сохранения базы данных";
            saveFileDialog1.InitialDirectory = "c:\\";
            saveFileDialog1.FileName = fileName;
            saveFileDialog1.Filter = "xml файлы (*.xml)|*.xml";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {        
                File.Copy(XMLDB.Path, saveFileDialog1.FileName);
                MessageBox.Show("Резервная копия базы данных создана по пути: " + saveFileDialog1.FileName);
            }
        }

        private void Button_LoadDB_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Title = "Выберите резервную копию базы данных";
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "xml файлы (*.xml)|*.xml";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var result = MessageBox.Show("Заменить текущую базу данных?", "Вы уверены?",
                                 MessageBoxButtons.YesNo,
                                 MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    File.Copy(openFileDialog1.FileName, XMLDB.Path, true);
                    MessageBox.Show("База данных восстановлена по пути: " + XMLDB.Path);
                }

            }
        }

        private void Button_RecreateDB_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Пересоздать текущую базу данных? Все записи будут навсегда утеряны!", "Внимание!",
                                 MessageBoxButtons.YesNo,
                                 MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                File.WriteAllBytes(XMLDB.Path, Convert.FromBase64String(XMLDB.DefaultXMLDBBase64)); //Декодируем строку с шаблоном базы данных из Base64 и создаем файл
                dataSet1.ReadXml(XMLDB.Path); // считываем в dataSet1 созданную нами базу в формате XML
                MessageBox.Show("База данных пересоздана по пути: " + XMLDB.Path);
            }
        }
    }
}
