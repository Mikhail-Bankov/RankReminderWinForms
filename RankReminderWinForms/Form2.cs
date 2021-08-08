using System;
using System.Linq;
using System.Xml;
using System.Text;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using System.Windows.Forms.VisualStyles;
using System.IO;

namespace RankReminderWinForms
{

    public partial class Form2 : Form
    {
        //string filePath = @"C:\C#_Projects\Rank_Reminder\BaseLichSost.xml";

        public DataSet dataSet1 = new DataSet(); // создаем DataSet с именем dataSet1
        string Required_PersonalFileNum, Required_PersonalNum, Required_Surname, Required_Name, Required_MiddleName, Required_DateOfBirth; // переменные для однозначного поиска сотрудника в Datatable

        List<string> ZvanieList = new List<string>() //словарь "Звания"
            {
                "Рядовой", "Мл. сержант", "Сержант", "Ст. сержант", "Старшина",
                "Прапорщик", "Ст. прапорщик", "Мл. лейтенант", "Лейтенант",
                "Ст. лейтенант", "Капитан", "Майор", "Подполковник", "Полковник",
                "Генерал-майор", "Генерал-лейтенант", "Генерал-полковник", "Генерал"
            };

        List<string> KlassnostList = new List<string>() //словарь "Классность"
            {
                "Отсутствует", "Специалист 3 класса", "Специалист 2 класса", "Специалист 1 класса", "Мастер"
            };

        // ###############  ЧИСЛОВЫЕ ПЕРЕМЕННЫЕ  ############### 

        // Общие
        int WantToDeleteRow = 0; // маркер намерения удалить строку. 0 - нет, 1 - да.
        int SomeRowsWasHidden = 0; // маркер, указывающий, нужно ли пересчитать порядковые номера у строку. 0 - нет, 1 - да.
        int StillResizing = 0; // маркер, сообщающий, закончилось ли изменение размера формы. 0 - нет, 1 - да.
        int Card1to9WasLoaded = 0; // маркер, по которому мы определяем, прогрузилась ли вкладка "Карточка 1-9". Это нужно, так
                                   // как при подгрузке textbox'ов срабатывает событие их изменения и база данных без необходимости
                                   // переписывается множество раз. 0 - по умолчанию, 1 - карточка прогрузилась. 
        int Card16to18WasLoaded = 0; // маркер, по которому мы определяем, прогрузилась ли вкладка "Карточка 16-18". Это нужно, так
                                     // как при подгрузке dateTimePicker'а срабатывает событие его изменения и база данных без необходимости
                                     // переписывается. 0 - по умолчанию, 1 - карточка прогрузилась. 
        int Card21and22WasLoaded = 0; // маркер, по которому мы определяем, прогрузилась ли вкладка "Карточка 21-22". Это нужно, так
                                      // как при подгрузке textbox'ов срабатывает событие их изменения и база данных без необходимости
                                      // переписывается множество раз. 0 - по умолчанию, 1 - карточка прогрузилась. 
        int Card23to25WasLoaded = 0; // маркер, по которому мы определяем, прогрузилась ли вкладка "Карточка 23-25".
        int Card26to29WasLoaded = 0; // маркер, по которому мы определяем, прогрузилась ли вкладка "Карточка 26-29". Это нужно, так
                                     // как при подгрузке textbox'ов срабатывает событие их изменения и база данных без необходимости
                                     // переписывается множество раз. 0 - по умолчанию, 1 - карточка прогрузилась. 

        int LastEditedCellRow; // индекс строки редактируемой ячейки
        int LastEditedCellCol; // индекс столбца редактируемой ячейки


        // Индексы колонок dataGridView1
        int IndexCnum; // Порядковый номер
        int IndexPersonalFileNum; // Номер личного дела
        int IndexPersonalNum; // Личный номер
        int IndexSurname; // Фамилия
        int IndexName; // Имя
        int IndexMiddleName; // Отчество
        int IndexGender; // Пол
        int IndexDateOfBirth; // Дата рождения
        int IndexPlaceOfBirth; // Место рождения
        int IndexRegistration; // Прописан
        int IndexPlaceOfLiving; // Место жительства
        int IndexPhoneRegistration; // Телефон по прописке
        int IndexPhonePlaceOfLiving; // Телефон по месту жительства
        int IndexPost; // Должность
        int IndexRank; // Звание
        int IndexRankDate; // Дата присвоения звания
        int IndexRankLimit; // Потолок по званию
        int IndexNextRankDate; // Следующая дата присвоения звания
        int IndexKlassnost; // Квалификационное звание
        int IndexKlassnostDate; // Дата присвоения квалиф. звания
        int IndexNextKlassnostDate; // Следующая дата присвоения квалиф. звания
        int IndexStudy; // Образование
        int IndexUchStepen; // Ученая степень
        int IndexPrisvZvaniy; // Дата присвоения званий и чинов
        int IndexMarried; // Семейное положение
        int IndexFamily; // Члены семьи
        int IndexTrudDeyat; // Трудовая деятельность до прихода
        int IndexStazhVysluga; // Стаж и выслуга до прихода
        int IndexDataPrisyagi; // Дата принятия присяги
        int IndexRabotaGFS; // Прохождение службы (работа) в ГФС России
        int IndexAttestaciya; // Аттестация
        int IndexNextAttestaciyaDate; // Дата следующей аттестации
        int IndexProfPodg; // Профессиональная подготовка
        int IndexKlassnostCheyPrikaz; // Чей приказ о присвоении квалиф. звания
        int IndexKlassnostNomerPrikaza; // Номер приказа о присвоении квалиф. звания
        int IndexKlassnostOld; // Предыдущие квалификационные звания
        int IndexNagrady; // Награды и поощрения
        int IndexProdlenie; // Продление службы
        int IndexBoevye; // Участие в боевых действиях
        int IndexRezerv; // Состояние в резерве
        int IndexVzyskaniya; // Взыскания
        int IndexUvolnenie; // Увольнение
        int IndexZapolnil; // Карточку заполнил
        int IndexDataZapolneniya; // Дата заполнения карточки
        int IndexImageString; // Изображение в виде текста

        // Индекс открытой личной карточки
        int IndexRowLichnayaKarta = 0;

        int IndexOfRowToExport;


        // ###############  СТРОКОВЫЕ ПЕРЕМЕННЫЕ  ############### 

        string CellValueToCompare; // переменная для проверки изменения в редактируемой ячейке

        string CurrentDataTableName = "Kadry";
        string OtherDataTableName = "Archive";


        //Создаем колонки для грида "Общий список"
        DataGridViewTextBoxColumn cnum = new DataGridViewTextBoxColumn(); // Порядковый номер
        DataGridViewTextBoxColumn personalfilenum = new DataGridViewTextBoxColumn(); // Номер личного дела
        DataGridViewTextBoxColumn personalnum = new DataGridViewTextBoxColumn(); // Личный номер
        DataGridViewTextBoxColumn surname = new DataGridViewTextBoxColumn(); // Фамилия
        DataGridViewTextBoxColumn name = new DataGridViewTextBoxColumn(); // Имя
        DataGridViewTextBoxColumn middleName = new DataGridViewTextBoxColumn(); // Отчество
        DataGridViewTextBoxColumn gender = new DataGridViewTextBoxColumn(); // Пол
        CalendarColumn dateofbirth = new CalendarColumn(); // Дата рождения
        DataGridViewTextBoxColumn placeofbirth = new DataGridViewTextBoxColumn(); // Место рождения
        DataGridViewTextBoxColumn registration = new DataGridViewTextBoxColumn(); // Прописан
        DataGridViewTextBoxColumn placeofliving = new DataGridViewTextBoxColumn(); // Место жительства
        DataGridViewTextBoxColumn phoneregistration = new DataGridViewTextBoxColumn(); // Телефон по прописке
        DataGridViewTextBoxColumn phoneplaceofliving = new DataGridViewTextBoxColumn(); // Телефон по месту жительства
        DataGridViewTextBoxColumn post = new DataGridViewTextBoxColumn(); // Должность
        DataGridViewComboBoxColumn rank = new DataGridViewComboBoxColumn(); // Звание
        CalendarColumn rankdate = new CalendarColumn(); // Дата присвоения звания
        DataGridViewComboBoxColumn ranklimit = new DataGridViewComboBoxColumn(); // Потолок по званию
        CalendarColumn nextrankdate = new CalendarColumn(); // Следующая дата присвоения звания
        DataGridViewComboBoxColumn klassnost = new DataGridViewComboBoxColumn(); // Квалификационное звание (Классность)
        CalendarColumn klassnostdate = new CalendarColumn(); // Дата присвоения квалиф. звания
        CalendarColumn nextklassnostdate = new CalendarColumn(); // Следующая дата присвоения квалиф. звания
        DataGridViewTextBoxColumn study = new DataGridViewTextBoxColumn(); // Образование
        DataGridViewTextBoxColumn uchstepen = new DataGridViewTextBoxColumn(); // Ученая степень
        DataGridViewTextBoxColumn prisvzvaniy = new DataGridViewTextBoxColumn(); // Дата присвоения званий и чинов        
        DataGridViewTextBoxColumn married = new DataGridViewTextBoxColumn(); // Семейное положение
        DataGridViewTextBoxColumn family = new DataGridViewTextBoxColumn(); // Члены семьи
        DataGridViewTextBoxColumn truddeyat = new DataGridViewTextBoxColumn(); // Трудовая деятельность до прихода
        DataGridViewTextBoxColumn stazhvysluga = new DataGridViewTextBoxColumn(); // Стаж и выслуга до прихода
        DataGridViewTextBoxColumn dataprisyagi = new DataGridViewTextBoxColumn(); // Дата принятия присяги
        DataGridViewTextBoxColumn rabotagfs = new DataGridViewTextBoxColumn(); // Прохождение службы (работа) в ГФС России
        DataGridViewTextBoxColumn attestaciya = new DataGridViewTextBoxColumn(); // Аттестация
        CalendarColumn nextattestaciyadate = new CalendarColumn(); // Дата следующей аттестации
        DataGridViewTextBoxColumn profpodg = new DataGridViewTextBoxColumn(); // Профессиональная подготовка
        DataGridViewTextBoxColumn klassnostcheyprikaz = new DataGridViewTextBoxColumn(); // Чей приказ о присвоении квалиф. звания
        DataGridViewTextBoxColumn klassnostnomerprikaza = new DataGridViewTextBoxColumn(); // Номер приказа о присвоении квалиф. звания
        DataGridViewTextBoxColumn klassnostold = new DataGridViewTextBoxColumn(); // Предыдущие квалификационные звания
        DataGridViewTextBoxColumn nagrady = new DataGridViewTextBoxColumn(); // Награды и поощрения
        DataGridViewTextBoxColumn prodlenie = new DataGridViewTextBoxColumn(); // Продление службы
        DataGridViewTextBoxColumn boevye = new DataGridViewTextBoxColumn(); // Участие в боевых действиях
        DataGridViewTextBoxColumn rezerv = new DataGridViewTextBoxColumn(); // Состояние в резерве
        DataGridViewTextBoxColumn vzyskaniya = new DataGridViewTextBoxColumn(); // Взыскания
        DataGridViewTextBoxColumn uvolnenie = new DataGridViewTextBoxColumn(); // Увольнение
        DataGridViewTextBoxColumn zapolnil = new DataGridViewTextBoxColumn(); // Увольнение
        DataGridViewTextBoxColumn datazapolneniya = new DataGridViewTextBoxColumn(); // Дата заполнения карточки     
        DataGridViewTextBoxColumn imagestring = new DataGridViewTextBoxColumn(); // Изображение в виде текста


        //Наименования столбцов dataGridView1
        const string Cnum_HeaderText = "№"; // Порядковый номер
        const string PersonalFileNum_HeaderText = "№ личного дела"; // Номер личного дела
        const string PersonalNum_HeaderText = "Личный номер"; // Личный номер
        const string Surname_HeaderText = "Фамилия"; // Фамилия
        const string Name_HeaderText = "Имя"; // Имя
        const string MiddleName_HeaderText = "Отчество"; // Отчество
        const string Gender_HeaderText = "Пол"; // Пол
        const string DateOfBirth_HeaderText = "Дата рождения"; // Дата рождения
        const string PlaceOfBirth_HeaderText = "Место рождения"; // Место рождения
        const string Registration_HeaderText = "Прописан"; // Прописан
        const string PlaceOfLiving_HeaderText = "Проживает"; // Место жительства
        const string PhoneRegistration_HeaderText = "Телефон по прописке"; // Телефон по прописке
        const string PhonePlaceOfLiving_HeaderText = "Тел. по месту жительства"; // Телефон по месту жительства
        const string Post_HeaderText = "Должность"; // Должность
        const string Rank_HeaderText = "Специальное звание"; // Звание
        const string RankDate_HeaderText = "Дата присвоения спец. звания"; // Дата присвоения звания
        const string RankLimit_HeaderText = "Потолок по спец. званию"; // Потолок по званию
        const string NextRankDate_HeaderText = "Следующая дата присвоения спец. звания"; // Следующая дата присвоения звания
        const string Klassnost_HeaderText = "Квалификационное звание"; // Квалификационное звание
        const string KlassnostDate_HeaderText = "Дата присвоения квалиф. звания"; // Дата присвоения квалиф. звания
        const string NextKlassnostDate_HeaderText = "Следующая дата присвоения квалиф. звания"; // Следующая дата присвоения квалиф. звания
        const string Study_HeaderText = "Образование"; // Образование
        const string UchStepen_HeaderText = "Ученая степень"; // Ученая степень
        const string PrisvZvaniy_HeaderText = "Присвоение званий"; // Дата присвоения званий и чинов
        const string Married_HeaderText = "Семейное положение"; // Семейное положение
        const string Family_HeaderText = "Члены семьи"; // Члены семьи
        const string TrudDeyat_HeaderText = "Трудовая деятельность до прихода"; // Трудовая деятельность до прихода
        const string StazhVysluga_HeaderText = "Стаж и выслуга до прихода"; // Стаж и выслуга до прихода
        const string DataPrisyagi_HeaderText = "Дата принятия присяги"; // Дата принятия присяги
        const string RabotaGFS_HeaderText = "Прохождение службы в ГФС России"; // Прохождение службы (работа) в ГФС России
        const string Attestaciya_HeaderText = "Аттестация"; // Аттестация
        const string NextAttestaciyaDate_HeaderText = "Дата следующей аттестации"; // Дата следующей аттестации
        const string ProfPodg_HeaderText = "Профессиональная подготовка"; // Профессиональная подготовка
        const string KlassnostCheyPrikaz_HeaderText = "Чей приказ о присвоении квалиф. звания"; // Чей приказ о присвоении квалиф. звания
        const string KlassnostNomerPrikaza_HeaderText = "Номер приказа о присвоении квалиф. звания"; // Номер приказа о присвоении квалиф. звания
        const string KlassnostOld_HeaderText = "Предыдущие квалификационные звания"; // Предыдущие квалификационные звания
        const string Nagrady_HeaderText = "Награды и поощрения"; // Награды и поощрения
        const string Prodlenie_HeaderText = "Продление службы"; // Продление службы
        const string Boevye_HeaderText = "Участие в боевых действиях"; // Участие в боевых действиях
        const string Rezerv_HeaderText = "Состояние в резерве"; // Состояние в резерве
        const string Vzyskaniya_HeaderText = "Взыскания"; // Взыскания
        const string Uvolnenie_HeaderText = "Увольнение"; // Увольнение
        const string Zapolnil_HeaderText = "Карточку заполнил"; // Карточку заполнил
        const string DataZapolneniya_HeaderText = "Дата заполнения карточки"; // Дата заполнения карточки
        const string ImageString_HeaderText = "Фото"; // Изображение в виде текста


        public Form2()
        {
            InitializeComponent();
            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            if (!File.Exists(XMLDB.Path)) // Если база данных в формате XML не существует...
            {
                MessageBox.Show("Похоже, что Вы запустили программу впервые, либо переместили файл базы данных. База будет создана заново.", "Внимание!",
                                 MessageBoxButtons.OK,
                                 MessageBoxIcon.Information);
                File.WriteAllBytes(XMLDB.Path, Convert.FromBase64String(XMLDB.DefaultXMLDBBase64)); //Декодируем строку с шаблоном базы данных из Base64 и создаем файл

                if (File.Exists(XMLDB.Path)) // Еще раз проверяем, создалась ли база данных
                {
                    MessageBox.Show("База данных успешно создана по пути:\n" + XMLDB.Path, "Внимание!",
                                 MessageBoxButtons.OK,
                                 MessageBoxIcon.Information);
                }
                else 
                {
                    MessageBox.Show("Произошла неизвестная ошибка при создании базы данных. Попробуйте запустить программу от имени администратора.", "Ошибка!",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
                }
            }

            dataSet1.ReadXml(XMLDB.Path); // считываем XML базу данных в dataSet1

            if (dataSet1.Tables["Kadry"] == null) // Если DataTable "Kadry" отсутствует
            {
                dataSet1.Tables.Add("Kadry");

                foreach (DataColumn column in dataSet1.Tables["BackUp"].Columns) // Заполняем новый DataTable необходимыми колонками
                {
                    dataSet1.Tables["Kadry"].Columns.Add(column.ColumnName);
                }
            }

            if (dataSet1.Tables["Archive"] == null) // Если DataTable "Archive" отсутствует
            {
                dataSet1.Tables.Add("Archive");

                foreach (DataColumn column in dataSet1.Tables["BackUp"].Columns) // Заполняем новый DataTable необходимыми колонками
                {
                    dataSet1.Tables["Archive"].Columns.Add(column.ColumnName);
                }
            }

            dataGridView1.DataSource = dataSet1.Tables[CurrentDataTableName]; // присваиваем источник данных для dataGridView1

            this.DrawDatagrid(); // формируем DataGrid
            this.CheckColumnsIndex(); // сверяем индексы просчитываемых столбцов
            this.PereschetZvanie(); // пересчитываем звания
            this.PereschetKlassnost(); // пересчитываем классность
            this.PereschetCnum(); // пересчитываем порядковые номера 
        }

        // ###############  ДЕЙСТВИЯ ПОСЛЕ ЗАГРУЗКИ ФОРМЫ  ###############
        private void Form2_Load(object sender, EventArgs e)
        {
            radioButton1.Checked = true; // изначально показывать все колонки
            Cards_groupBox.Visible = false; // изначально не показывать кнопки карточек

            // ###############  ОБРАБОТЧИКИ СОБЫТИЙ УДАЛЕНИЯ СТРОК В ТАБЛИЦАХ  ###############
            dataSet1.Tables["Kadry"].RowDeleting += new System.Data.DataRowChangeEventHandler(RowDeleting); // обработчик события попытки удаления строки
            dataSet1.Tables["Kadry"].RowDeleted += new System.Data.DataRowChangeEventHandler(RowDeleted); // обработчик события удаления строки                                                                                         
            dataSet1.Tables["Archive"].RowDeleting += new System.Data.DataRowChangeEventHandler(RowDeleting); // обработчик события попытки удаления строки
            dataSet1.Tables["Archive"].RowDeleted += new System.Data.DataRowChangeEventHandler(RowDeleted); // обработчик события удаления строки     
            //dataGridView_Family.UserDeletedRow += new System.Windows.Forms.DataGridViewRowEventHandler(WhichRowDeleted); // обработчик события удаления строки
            


            // ###############  РАЗНЫЕ ОБРАБОТЧИКИ СОБЫТИЙ  ###############
            dataGridView1.Sorted += new System.EventHandler(dataGridView1_Sorted); // обработчик события сортировки колонки
            tabControl1.Selecting += tabControl1_Selecting; // обработчик события перед сменой активной вкладки
            tabControl1.SelectedIndexChanged += tabControl1_SelectedIndexChanged; // обработчик события смены активной вкладки


            // ###############  ОБРАБОТЧИК ComboBox'ов ДЛЯ ЦЕНТРОВКИ В РЕЖИМЕ РЕДАКТИРОВАНИЯ  ###############
            Gender_comboBox.DrawItem += new DrawItemEventHandler(CenteredComboBox.ComboBox_DrawItem_Centered); // Пол
            Klassnost_comboBox.DrawItem += new DrawItemEventHandler(CenteredComboBox.ComboBox_DrawItem_Centered); // Текущее квалификационное звание
            Rank_comboBox.DrawItem += new DrawItemEventHandler(CenteredComboBox.ComboBox_DrawItem_Centered); // Текущее звание
            RankLimit_comboBox.DrawItem += new DrawItemEventHandler(CenteredComboBox.ComboBox_DrawItem_Centered); // Потолок по званию
                                                                                                                  // Test_comboBox.DrawItem += new DrawItemEventHandler(OnDrawItem); // Потолок по званию

            // ###############  ОБРАБОТЧИК ComboBox'ов В dataGridView ДЛЯ ЦЕНТРОВКИ В РЕЖИМЕ РЕДАКТИРОВАНИЯ  ###############
            dataGridView_Prodlenie.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(CenteredComboBox.MyDGV_EditingControlShowing); // Продление службы
            dataGridView_KlassnostOld.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(CenteredComboBox.MyDGV_EditingControlShowing); // Классность
            dataGridView_Attestaciya.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(CenteredComboBox.MyDGV_EditingControlShowing); // Аттестация
            dataGridView_Family.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(CenteredComboBox.MyDGV_EditingControlShowing); // Члены семьи
            dataGridView_Married.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(CenteredComboBox.MyDGV_EditingControlShowing); // Семейное положение
            dataGridView_ProfPodg.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(CenteredComboBox.MyDGV_EditingControlShowing); // Профессиональная подготовка

            // ###############  ОБРАБОТЧИКИ СОБЫТИЙ ИЗМЕНЕНИЯ ТАБЛИЦ  ###############
            dataGridView_Married.CellValueChanged += new DataGridViewCellEventHandler(MarriedAdd_Click);
            dataGridView_Family.CellValueChanged += new DataGridViewCellEventHandler(FamilyAddPerson_Click);
            dataGridView_Study.CellValueChanged += new DataGridViewCellEventHandler(StudyAdd_Click);
            dataGridView_UchStepen.CellValueChanged += new DataGridViewCellEventHandler(UchStepenAdd_Click);
            dataGridView_PrisvZvaniy.CellValueChanged += new DataGridViewCellEventHandler(ZvanieAdd_Click);
            dataGridView_TrudDeyat.CellValueChanged += new DataGridViewCellEventHandler(TrudDeyatAdd_Click);
            dataGridView_StazhVysluga.CellValueChanged += new DataGridViewCellEventHandler(SaveChangesToDataGridView_StazhVysluga);
            dataGridView_RabotaGFS.CellValueChanged += new DataGridViewCellEventHandler(RabotaGFSAdd_Click);
            dataGridView_Attestaciya.CellValueChanged += new DataGridViewCellEventHandler(AttestaciyaAdd_Click);
            dataGridView_ProfPodg.CellValueChanged += new DataGridViewCellEventHandler(ProfPodgAdd_Click);
            dataGridView_KlassnostOld.CellValueChanged += new DataGridViewCellEventHandler(KlassnostAdd_Click);
            dataGridView_Nagrady.CellValueChanged += new DataGridViewCellEventHandler(NagradyAdd_Click);
            dataGridView_Prodlenie.CellValueChanged += new DataGridViewCellEventHandler(Prodlenie_checkBox_CheckedChanged);
            dataGridView_Boevye.CellValueChanged += new DataGridViewCellEventHandler(BoevyeAdd_Click);
            dataGridView_Rezerv.CellValueChanged += new DataGridViewCellEventHandler(RezervAdd_Click);
            dataGridView_Vzyskaniya.CellValueChanged += new DataGridViewCellEventHandler(VzyskaniyaAdd_Click);
            dataGridView_Uvolnenie.CellValueChanged += new DataGridViewCellEventHandler(UvolnenieAdd_Click);

            // ###############  ОБРАБОТЧИКИ СОБЫТИЙ ИЗМЕНЕНИЯ РАЗМЕРА ТАБЛИЦ  ###############
            ResizeBegin += this.FormResizeBegin;
            Resize += this.FormResize;
            ResizeEnd += this.FormResizeEnd;

            // ###############  ПЕРЕПИСЫВАЕМ ОТРИСОВКУ ComboBox'ов В НЕОБХОДИМЫХ ТАБЛИЦАХ  ###############

            // Образование
            Study_FormaObucheniya.CellTemplate = new DataGridViewComboBoxCellEx();//CellTemplate
            Study_FormaObucheniya.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            Study_FormaObucheniya.Items.AddRange(new string[] { "Высшее (очное)", "Высшее (заочное)", "Среднее (очное)", "Среднее (заочное)" });
            //Study_FormaObucheniya.FlatStyle = FlatStyle.Flat;

            // Профессиональная подготовка
            ProfPodg_VidObuch.CellTemplate = new DataGridViewComboBoxCellEx();//CellTemplate
            ProfPodg_VidObuch.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            ProfPodg_VidObuch.Items.AddRange(new string[] { "Первоначальное обучение", "Стажировка", "Повышение квалификации" });

            // Трудовая деятельность до прихода в ГФС (сокращение)
            TrudDeyat_Sokrash.CellTemplate = new DataGridViewComboBoxCellEx();//CellTemplate
            TrudDeyat_Sokrash.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            TrudDeyat_Sokrash.Items.AddRange(new string[] { "У", "УВ", "Р", "ВС", "СЗ", "ГС" });

            Cnum_label.Text = (IndexRowLichnayaKarta + 1).ToString() + " из " + dataGridView1.RowCount.ToString(); // Порядковый номер личной карточки

            label1.Text = "Сегодня: " + DateTime.Today.ToShortDateString(); //ставим текущую дату внизу формы
        }

        // ###############  ОТРИСОВКА dataGridView1  ###############
        private void DrawDatagrid()
        {
            DataGridViewCellStyle style = dataGridView1.ColumnHeadersDefaultCellStyle;
            style.Alignment = DataGridViewContentAlignment.MiddleCenter; // выравниваем текст заголовков по центру

            //Столбец "Порядковый номер"
            cnum.HeaderText = Cnum_HeaderText;
            cnum.DataPropertyName = "Cnum";
            cnum.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells; //автоширина по содержимому ячеек и заголовка
            cnum.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; //выравниваем содержимое столбца по центру
            cnum.SortMode = DataGridViewColumnSortMode.NotSortable; //запрещаем сортировку данной колонки
            cnum.ReadOnly = true; // запрещаем редактирование данной колонки

            //Столбец "Номер личного дела"
            personalfilenum.HeaderText = PersonalFileNum_HeaderText;
            personalfilenum.DataPropertyName = "PersonalFileNum";
            personalfilenum.MinimumWidth = 150;
            personalfilenum.FillWeight = 250;

            //Столбец "Личный номер"
            personalnum.HeaderText = PersonalNum_HeaderText;
            personalnum.DataPropertyName = "PersonalNum";
            personalnum.MinimumWidth = 150;
            personalnum.FillWeight = 250;

            //Столбец "Фамилия"
            surname.HeaderText = Surname_HeaderText;
            surname.DataPropertyName = "Surname";
            surname.MinimumWidth = 150;
            surname.FillWeight = 250;

            //Столбец "Имя"
            name.HeaderText = Name_HeaderText;
            name.DataPropertyName = "Name";
            name.MinimumWidth = 150;
            name.FillWeight = 250;

            //Столбец "Отчество"
            middleName.HeaderText = MiddleName_HeaderText;
            middleName.DataPropertyName = "MiddleName";
            middleName.MinimumWidth = 150;
            middleName.FillWeight = 250;

            //Столбец "Пол"
            gender.HeaderText = Gender_HeaderText;
            gender.DataPropertyName = "Gender";
            gender.MinimumWidth = 150;
            gender.FillWeight = 250;

            //Столбец "Дата рождения"
            dateofbirth.HeaderText = DateOfBirth_HeaderText;
            dateofbirth.DataPropertyName = "DateOfBirth";
            dateofbirth.MinimumWidth = 120;
            dateofbirth.FillWeight = 130;
            dateofbirth.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            //Столбец "Место рождения"
            placeofbirth.HeaderText = PlaceOfBirth_HeaderText;
            placeofbirth.DataPropertyName = "PlaceOfBirth";
            placeofbirth.MinimumWidth = 150;
            placeofbirth.FillWeight = 250;

            //Столбец "Прописан"
            registration.HeaderText = Registration_HeaderText;
            registration.DataPropertyName = "Registration";
            registration.MinimumWidth = 150;
            registration.FillWeight = 250;

            //Столбец "Место жительства"
            placeofliving.HeaderText = PlaceOfLiving_HeaderText;
            placeofliving.DataPropertyName = "PlaceOfLiving";
            placeofliving.MinimumWidth = 150;
            placeofliving.FillWeight = 250;

            //Столбец "Телефон по прописке"
            phoneregistration.HeaderText = PhoneRegistration_HeaderText;
            phoneregistration.DataPropertyName = "PhoneRegistration";
            phoneregistration.MinimumWidth = 150;
            phoneregistration.FillWeight = 250;

            //Столбец "Телефон по месту жительства"
            phoneplaceofliving.HeaderText = PhonePlaceOfLiving_HeaderText;
            phoneplaceofliving.DataPropertyName = "PhonePlaceOfLiving";
            phoneplaceofliving.MinimumWidth = 150;
            phoneplaceofliving.FillWeight = 250;

            //Столбец "Должность"
            post.HeaderText = Post_HeaderText;
            post.DataPropertyName = "Post";
            post.MinimumWidth = 150;
            post.FillWeight = 250;
            post.SortMode = DataGridViewColumnSortMode.NotSortable; //запрещаем сортировку данной колонки

            //Столбец "Звание"
            rank.DataSource = ZvanieList;
            rank.FlatStyle = FlatStyle.Flat;
            rank.HeaderText = Rank_HeaderText;
            rank.DataPropertyName = "Rank";
            rank.MinimumWidth = 150;
            rank.FillWeight = 160;

            //Столбец "Дата присвоения звания"
            rankdate.HeaderText = RankDate_HeaderText;
            rankdate.DataPropertyName = "RankDate";
            rankdate.MinimumWidth = 130;
            rankdate.FillWeight = 140;
            rankdate.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            //Столбец "Потолок по званию"
            ranklimit.DataSource = ZvanieList;
            ranklimit.FlatStyle = FlatStyle.Flat;
            ranklimit.HeaderText = RankLimit_HeaderText;
            ranklimit.DataPropertyName = "RankLimit";
            ranklimit.MinimumWidth = 150;
            ranklimit.FillWeight = 160;

            //Столбец "Следующая дата присвоения звания"
            nextrankdate.HeaderText = NextRankDate_HeaderText;
            nextrankdate.DataPropertyName = "NextRankDate";
            nextrankdate.MinimumWidth = 130;
            nextrankdate.FillWeight = 140;
            nextrankdate.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            //Столбец "Квалификационное звание"
            klassnost.DataSource = KlassnostList;
            klassnost.FlatStyle = FlatStyle.Flat;
            klassnost.HeaderText = Klassnost_HeaderText;
            klassnost.DataPropertyName = "Klassnost";
            klassnost.MinimumWidth = 200;
            klassnost.FillWeight = 210;

            //Столбец "Дата присвоения квалиф. звания"
            klassnostdate.HeaderText = KlassnostDate_HeaderText;
            klassnostdate.DataPropertyName = "KlassnostDate";
            klassnostdate.MinimumWidth = 130;
            klassnostdate.FillWeight = 140;
            klassnostdate.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            //Столбец "Следующая дата присвоения квалиф. звания"
            nextklassnostdate.HeaderText = NextKlassnostDate_HeaderText;
            nextklassnostdate.DataPropertyName = "NextKlassnostDate";
            nextklassnostdate.MinimumWidth = 130;
            nextklassnostdate.FillWeight = 140;
            nextklassnostdate.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            //Столбец "Образование"
            study.HeaderText = Study_HeaderText;
            study.DataPropertyName = "Study";
            study.MinimumWidth = 130;
            study.FillWeight = 140;

            //Столбец "Ученая степень"
            uchstepen.HeaderText = UchStepen_HeaderText;
            uchstepen.DataPropertyName = "UchStepen";
            uchstepen.MinimumWidth = 130;
            uchstepen.FillWeight = 140;

            //Столбец "Присвоение званий, классных чинов"
            prisvzvaniy.HeaderText = PrisvZvaniy_HeaderText;
            prisvzvaniy.DataPropertyName = "PrisvZvaniy";
            prisvzvaniy.MinimumWidth = 130;
            prisvzvaniy.FillWeight = 140;

            //Столбец "Семейное положение"
            married.HeaderText = Married_HeaderText;
            married.DataPropertyName = "Married";
            married.MinimumWidth = 130;
            married.FillWeight = 140;

            //Столбец "Члены семьи"
            family.HeaderText = Family_HeaderText;
            family.DataPropertyName = "Family";
            family.MinimumWidth = 130;
            family.FillWeight = 140;

            //Столбец "Трудовая деятельность до прихода"
            truddeyat.HeaderText = TrudDeyat_HeaderText;
            truddeyat.DataPropertyName = "TrudDeyat";
            truddeyat.MinimumWidth = 130;
            truddeyat.FillWeight = 140;

            //Столбец "Трудовая деятельность до прихода"
            stazhvysluga.HeaderText = StazhVysluga_HeaderText;
            stazhvysluga.DataPropertyName = "StazhVysluga";
            stazhvysluga.MinimumWidth = 130;
            stazhvysluga.FillWeight = 140;

            //Столбец "Дата принятия присяги"
            dataprisyagi.HeaderText = DataPrisyagi_HeaderText;
            dataprisyagi.DataPropertyName = "DataPrisyagi";
            dataprisyagi.MinimumWidth = 130;
            dataprisyagi.FillWeight = 140;

            //Столбец "Прохождение службы (работа) в ГФС России"
            rabotagfs.HeaderText = RabotaGFS_HeaderText;
            rabotagfs.DataPropertyName = "RabotaGFS";
            rabotagfs.MinimumWidth = 130;
            rabotagfs.FillWeight = 140;

            //Столбец "Аттестация"
            attestaciya.HeaderText = Attestaciya_HeaderText;
            attestaciya.DataPropertyName = "Attestaciya";
            attestaciya.MinimumWidth = 130;
            attestaciya.FillWeight = 140;

            //Столбец "Дата следующей аттестации"
            nextattestaciyadate.HeaderText = NextAttestaciyaDate_HeaderText;
            nextattestaciyadate.DataPropertyName = "nextattestaciyadate";
            nextattestaciyadate.MinimumWidth = 130;
            nextattestaciyadate.FillWeight = 140;
            nextattestaciyadate.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            nextattestaciyadate.ReadOnly = true;

            //Столбец "Профессиональная подготовка"
            profpodg.HeaderText = ProfPodg_HeaderText;
            profpodg.DataPropertyName = "ProfPodg";
            profpodg.MinimumWidth = 130;
            profpodg.FillWeight = 140;

            //Столбец "Чей приказ о присвоении квалиф. звания"
            klassnostcheyprikaz.HeaderText = KlassnostCheyPrikaz_HeaderText;
            klassnostcheyprikaz.DataPropertyName = "KlassnostCheyPrikaz";
            klassnostcheyprikaz.MinimumWidth = 130;
            klassnostcheyprikaz.FillWeight = 140;

            //Столбец "Номер приказа о присвоении квалиф. звания"
            klassnostnomerprikaza.HeaderText = KlassnostNomerPrikaza_HeaderText;
            klassnostnomerprikaza.DataPropertyName = "KlassnostNomerPrikaza";
            klassnostnomerprikaza.MinimumWidth = 130;
            klassnostnomerprikaza.FillWeight = 140;

            //Столбец "Предыдущие квалификационные звания"
            klassnostold.HeaderText = KlassnostOld_HeaderText;
            klassnostold.DataPropertyName = "KlassnostOld";
            klassnostold.MinimumWidth = 130;
            klassnostold.FillWeight = 140;

            //Столбец "Награды и поощрения"
            nagrady.HeaderText = Nagrady_HeaderText;
            nagrady.DataPropertyName = "Nagrady";
            nagrady.MinimumWidth = 130;
            nagrady.FillWeight = 140;

            //Столбец "Продление службы"
            prodlenie.HeaderText = Prodlenie_HeaderText;
            prodlenie.DataPropertyName = "Prodlenie";
            prodlenie.MinimumWidth = 130;
            prodlenie.FillWeight = 140;

            //Столбец "Участие в боевых действиях"
            boevye.HeaderText = Boevye_HeaderText;
            boevye.DataPropertyName = "Boevye";
            boevye.MinimumWidth = 130;
            boevye.FillWeight = 140;

            //Столбец "Состояние в резерве"
            rezerv.HeaderText = Rezerv_HeaderText;
            rezerv.DataPropertyName = "Rezerv";
            rezerv.MinimumWidth = 130;
            rezerv.FillWeight = 140;

            //Столбец "Взыскания"
            vzyskaniya.HeaderText = Vzyskaniya_HeaderText;
            vzyskaniya.DataPropertyName = "Vzyskaniya";
            vzyskaniya.MinimumWidth = 130;
            vzyskaniya.FillWeight = 140;

            //Столбец "Увольнение"
            uvolnenie.HeaderText = Uvolnenie_HeaderText;
            uvolnenie.DataPropertyName = "Uvolnenie";
            uvolnenie.MinimumWidth = 130;
            uvolnenie.FillWeight = 140;

            //Столбец "Карточку заполнил"
            zapolnil.HeaderText = Zapolnil_HeaderText;
            zapolnil.DataPropertyName = "Zapolnil";
            zapolnil.MinimumWidth = 130;
            zapolnil.FillWeight = 140;

            //Столбец "Дата заполнения карточки"
            datazapolneniya.HeaderText = DataZapolneniya_HeaderText;
            datazapolneniya.DataPropertyName = "DataZapolneniya";
            datazapolneniya.MinimumWidth = 130;
            datazapolneniya.FillWeight = 140;

            //Столбец "Фото"
            imagestring.HeaderText = ImageString_HeaderText;
            imagestring.DataPropertyName = "Image";
            imagestring.MinimumWidth = 150;
            imagestring.FillWeight = 250;
            imagestring.SortMode = DataGridViewColumnSortMode.NotSortable; //запрещаем сортировку данной колонки

            //Выводим столбцы в нужном нам порядке
            dataGridView1.Columns.AddRange(cnum, personalfilenum, personalnum, surname, name, middleName, gender, dateofbirth, placeofbirth, registration, placeofliving, phoneregistration, phoneplaceofliving, post, rank,
            rankdate, ranklimit, nextrankdate, klassnost, klassnostdate, nextklassnostdate, study, uchstepen, prisvzvaniy, married, family, truddeyat, stazhvysluga, dataprisyagi, rabotagfs, attestaciya, nextattestaciyadate, profpodg, klassnostcheyprikaz,
            klassnostnomerprikaza, klassnostold, nagrady, prodlenie, boevye, rezerv, vzyskaniya, uvolnenie, zapolnil, datazapolneniya, imagestring);
        }


    // ###############  ОСНОВНОЙ МЕТОД СЧИТЫВАНИЯ ИНДЕКСОВ У НЕОБХОДИМЫХ КОЛОНОК  ###############
    private void CheckColumnsIndex()
        {
            foreach (DataGridViewColumn currColumn in dataGridView1.Columns) //пробегаем по всем колонкам в dataGridView1
            {
                switch (currColumn.HeaderText)
                {
                    case Cnum_HeaderText: // Порядковый номер
                        IndexCnum = currColumn.DisplayIndex;
                        break;
                    case PersonalFileNum_HeaderText: // Номер личного дела
                        IndexPersonalFileNum = currColumn.DisplayIndex;
                        break;
                    case PersonalNum_HeaderText: // Личный номер
                        IndexPersonalNum = currColumn.DisplayIndex;
                        break;
                    case Surname_HeaderText: // Фамилия
                        IndexSurname = currColumn.DisplayIndex;
                        break;
                    case Name_HeaderText: // Имя
                        IndexName = currColumn.DisplayIndex;
                        break;
                    case MiddleName_HeaderText: // Отчество
                        IndexMiddleName = currColumn.DisplayIndex;
                        break;
                    case Gender_HeaderText: // Пол
                        IndexGender = currColumn.DisplayIndex;
                        break;
                    case DateOfBirth_HeaderText: // Дата рождения
                        IndexDateOfBirth = currColumn.DisplayIndex;
                        break;
                    case PlaceOfBirth_HeaderText: // Место рождения
                        IndexPlaceOfBirth = currColumn.DisplayIndex;
                        break;
                    case Registration_HeaderText: // Прописан
                        IndexRegistration = currColumn.DisplayIndex;
                        break;
                    case PlaceOfLiving_HeaderText: // Место жительства
                        IndexPlaceOfLiving = currColumn.DisplayIndex;
                        break;
                    case PhoneRegistration_HeaderText: // Телефон по прописке
                        IndexPhoneRegistration = currColumn.DisplayIndex;
                        break;
                    case PhonePlaceOfLiving_HeaderText: // Телефон по месту жительства
                        IndexPhonePlaceOfLiving = currColumn.DisplayIndex;
                        break;
                    case Post_HeaderText: // Должность
                        IndexPost = currColumn.DisplayIndex;
                        break;
                    case Rank_HeaderText: // Звание
                        IndexRank = currColumn.DisplayIndex;
                        break;
                    case RankDate_HeaderText: // Дата присвоения звания
                        IndexRankDate = currColumn.DisplayIndex;
                        break;
                    case RankLimit_HeaderText: // Потолок по званию
                        IndexRankLimit = currColumn.DisplayIndex;
                        break;
                    case NextRankDate_HeaderText: // Следующая дата присвоения звания
                        IndexNextRankDate = currColumn.DisplayIndex;
                        break;
                    case Klassnost_HeaderText: // Квалификационное звание
                        IndexKlassnost = currColumn.DisplayIndex;
                        break;
                    case KlassnostDate_HeaderText: // Дата присвоения квалиф. звания
                        IndexKlassnostDate = currColumn.DisplayIndex;
                        break;
                    case NextKlassnostDate_HeaderText: // Следующая дата присвоения квалиф. звания
                        IndexNextKlassnostDate = currColumn.DisplayIndex;
                        break;
                    case Study_HeaderText: // Образование
                        IndexStudy = currColumn.DisplayIndex;
                        break;
                    case UchStepen_HeaderText: // Ученая степень
                        IndexUchStepen = currColumn.DisplayIndex;
                        break;
                    case PrisvZvaniy_HeaderText: // Присвоение званий, классных чинов
                        IndexPrisvZvaniy = currColumn.DisplayIndex;
                        break;
                    case Married_HeaderText: //Семейное положение
                        IndexMarried = currColumn.DisplayIndex;
                        break;
                    case Family_HeaderText: // Члены семьи
                        IndexFamily = currColumn.DisplayIndex;
                        break;
                    case TrudDeyat_HeaderText: // Трудовая деятельность до прихода
                        IndexTrudDeyat = currColumn.DisplayIndex;
                        break;
                    case StazhVysluga_HeaderText: // Стаж и выслуга до прихода
                        IndexStazhVysluga = currColumn.DisplayIndex;
                        break;
                    case DataPrisyagi_HeaderText: // Дата принятия присяги
                        IndexDataPrisyagi = currColumn.DisplayIndex;
                        break;
                    case RabotaGFS_HeaderText: // Прохождение службы (работа) в ГФС России
                        IndexRabotaGFS = currColumn.DisplayIndex;
                        break;
                    case Attestaciya_HeaderText: // Аттестация
                        IndexAttestaciya = currColumn.DisplayIndex;
                        break;
                    case NextAttestaciyaDate_HeaderText: // Дата следующей аттестации
                        IndexNextAttestaciyaDate = currColumn.DisplayIndex;
                        break;
                    case ProfPodg_HeaderText: // Профессиональная подготовка
                        IndexProfPodg = currColumn.DisplayIndex;
                        break;
                    case KlassnostCheyPrikaz_HeaderText: // Чей приказ о присвоении квалиф. звания
                        IndexKlassnostCheyPrikaz = currColumn.DisplayIndex;
                        break;
                    case KlassnostNomerPrikaza_HeaderText: // Номер приказа о присвоении квалиф. звания
                        IndexKlassnostNomerPrikaza = currColumn.DisplayIndex;
                        break;
                    case KlassnostOld_HeaderText: // Предыдущие квалификационные звания
                        IndexKlassnostOld = currColumn.DisplayIndex;
                        break;
                    case Nagrady_HeaderText: // Награды и поощрения
                        IndexNagrady = currColumn.DisplayIndex;
                        break;
                    case Prodlenie_HeaderText: // Продление службы
                        IndexProdlenie = currColumn.DisplayIndex;
                        break;
                    case Boevye_HeaderText: // Участие в боевых действиях
                        IndexBoevye = currColumn.DisplayIndex;
                        break;
                    case Rezerv_HeaderText: // Состояние в резерве
                        IndexRezerv = currColumn.DisplayIndex;
                        break;
                    case Vzyskaniya_HeaderText: // Взыскания
                        IndexVzyskaniya = currColumn.DisplayIndex;
                        break;
                    case Uvolnenie_HeaderText: // Увольнение
                        IndexUvolnenie = currColumn.DisplayIndex;
                        break;
                    case Zapolnil_HeaderText: // Карточку заполнил
                        IndexZapolnil = currColumn.DisplayIndex;
                        break;
                    case DataZapolneniya_HeaderText: // Дата заполнения карточки
                        IndexDataZapolneniya = currColumn.DisplayIndex;
                        break;
                    case ImageString_HeaderText: // Фото сотрудника
                        IndexImageString = currColumn.DisplayIndex;
                        break;
                }
            }
        }



        // ###############  ПЕРЕСЧЁТ КОЛОНКИ С ПОРЯДКОВЫМИ НОМЕРАМИ СТРОК  ###############
        // Пересчет колонки с порядковыми номерами строк
        private void PereschetCnum()
        {
            if (dataGridView1.Rows.Count != 0) // Проверка dataGridView1 на пустоту
            {
                for (int i = 0; i < dataGridView1.RowCount; i++) // Заполняем колонку с порядковыми номерами строк
                {
                    dataGridView1[IndexCnum, i].Value = i + 1; // Увеличиваем порядковый номер в каждой последующей строке на единицу
                }
            }
        }


        // ###############  ОСНОВНОЙ МЕТОД ПРОВЕРКИ И ПЕРЕСЧЁТА СРОКОВ ВЫСЛУГИ  ###############
        private void PereschetZvanie()
        {
            if (dataGridView1.Rows.Count != 0) // Проверка dataGridView1 на пустоту
            {
                int RankVariable = 0;
                int CurrentRankToCompare = 0; //"вес" текущего звания
                int RankLimitToCompare = 0; //"вес" потолка по званию
                int NumberOfYears = 1; //срок выслуги до следующего звания
                int pYear2, pMonth2, pDay2; //переменные для парсинга текстовой даты в год, месяц, день
                string peremennaya2; //переменная для хранения даты из ячеек

                foreach (DataGridViewRow currRow in dataGridView1.Rows) // проходим по каждой строке в таблице
                {
                    string elemStr = currRow.Cells[IndexRank].Value.ToString(); // Переменная с текущим званием
                    string elemStr2 = currRow.Cells[IndexRankLimit].Value.ToString(); // Переменная с "потолком" по званию


                    if (elemStr == elemStr2) // Чтобы лишний раз не гонять циклы, первым делом проверяем,
                                             // равно ли текущее звание "потолку" звания по должности.
                                             // Если равно, то пишем "роста нет" и переходим к следующей строке.
                    {
                        currRow.Cells[IndexNextRankDate].Value = "роста нет";
                    }
                    else // Если же звания отличаются, то запускаем их обработку и высчитываем дату присвоения следующего звания
                    
                    {
                        string elemVariable = "unknown";

                        for (int i = 0; i < 2; i++) // Прогоняем цикл два раза:
                                                    // 1. Для определения текущего звания сотрудника и установки его числового "веса"
                                                    // 2. Для определения "потолка" по званию и установки его числового "веса"
                        {
                            if (i == 0) // если первый проход цикла...
                            {
                                elemVariable = elemStr; // ...то работаем с текущим званием
                            }
                            else if (i == 1) // если второй проход цикла...
                            {
                                elemVariable = elemStr2; // ...то работаем с "потолком" по званию
                            }

                            switch (elemVariable) // непосредственно проверка званий и установка их числового "веса" для сравнения 
                            {
                                case "Рядовой":             // 1 год
                                    RankVariable = 1;
                                    break;
                                case "Мл. сержант":         // 1 год
                                    RankVariable = 2;
                                    break;
                                case "Сержант":             // 2 года
                                    RankVariable = 3;
                                    break;
                                case "Ст. сержант":         // 3 года
                                    RankVariable = 4;
                                    break;
                                case "Старшина":            // НЕ УСТАНОВЛЕН
                                    RankVariable = 5;
                                    break;
                                case "Прапорщик":           // 5 лет
                                    RankVariable = 6;
                                    break;
                                case "Ст. прапорщик":       // НЕ УСТАНОВЛЕН
                                    RankVariable = 7;
                                    break;
                                case "Мл. лейтенант":       // 1 год
                                    RankVariable = 8;
                                    break;
                                case "Лейтенант":           // 2 года
                                    RankVariable = 9;
                                    break;
                                case "Ст. лейтенант":       // 3 года
                                    RankVariable = 10;
                                    break;
                                case "Капитан":             // 3 года
                                    RankVariable = 11;
                                    break;
                                case "Майор":               // 4 года
                                    RankVariable = 12;
                                    break;
                                case "Подполковник":        // 5 лет
                                    RankVariable = 13;
                                    break;
                                case "Полковник":           // НЕ УСТАНОВЛЕН
                                    RankVariable = 14;
                                    break;
                                case "Генерал-майор":       // НЕ УСТАНОВЛЕН
                                    RankVariable = 15;
                                    break;
                                case "Генерал-лейтенант":   // НЕ УСТАНОВЛЕН
                                    RankVariable = 16;
                                    break;
                                case "Генерал-полковник":   // НЕ УСТАНОВЛЕН
                                    RankVariable = 17;
                                    break;
                                case "Генерал":             // НЕ УСТАНОВЛЕН
                                    RankVariable = 18;
                                    break;
                            }

                            if (i == 0) // если первый проход цикла...
                            {
                                CurrentRankToCompare = RankVariable; // ...присваиваем числовой "вес" переменной CurrentRankToCompare (текущее звание)              
                            }
                            else if (i == 1) // если второй проход цикла...
                            {
                                RankLimitToCompare = RankVariable; // ...присваиваем числовой "вес" переменной RankLimitToCompare ("потолок" по званию)
                            }
                        }

                        //НАЧИНАЕМ СВЕРЯТЬ ЧИСЛОВОЙ "ВЕС" ЗВАНИЙ И ВЫСЧИТЫВАТЬ ДАТУ ПРИСВОЕНИЯ СЛЕДУЮЩЕГО ЗВАНИЯ
     
                        if (CurrentRankToCompare > RankLimitToCompare) // Если текущее звание выше звания по должности...
                        {                        
                            currRow.Cells[IndexNextRankDate].Value = "роста нет"; // ...значит расти дальше некуда
                        }
                        else if ((CurrentRankToCompare < RankLimitToCompare) && (CurrentRankToCompare == 5 || CurrentRankToCompare == 7 || CurrentRankToCompare == 14 || CurrentRankToCompare == 15 || CurrentRankToCompare == 16 || CurrentRankToCompare == 17 || CurrentRankToCompare == 18))
                        {                        
                            currRow.Cells[IndexNextRankDate].Value = "не установлена"; // для перечисленных званий срок выслуги не установлен
                        }
                        else if (CurrentRankToCompare < RankLimitToCompare) // Если есть куда расти, то определяем количество лет до следующего звания
                        {              
                            // Звания со сроком выслуги в один год:
                            if ((CurrentRankToCompare == 1) || (CurrentRankToCompare == 2) || (CurrentRankToCompare == 8)) NumberOfYears = 1;
                            // Звания со сроком выслуги в два года:
                            if ((CurrentRankToCompare == 3) || (CurrentRankToCompare == 9)) NumberOfYears = 2;
                            // Звания со сроком выслуги в три года:
                            if ((CurrentRankToCompare == 4) || (CurrentRankToCompare == 10) || (CurrentRankToCompare == 11)) NumberOfYears = 3;
                            // Звания со сроком выслуги в четыре года:
                            if (CurrentRankToCompare == 12) NumberOfYears = 4;
                            // Звания со сроком выслуги в пять лет:
                            if ((CurrentRankToCompare == 6) || (CurrentRankToCompare == 13)) NumberOfYears = 5;

                            peremennaya2 = currRow.Cells[IndexRankDate].Value.ToString(); // считываем значение даты из ячейки в peremennaya2 типа string
                                                                                          //MessageBox.Show(peremennaya2);
                            pYear2 = Convert.ToInt32(peremennaya2.Substring(6, 4)); // парсим peremennaya2 с 7го символа, длина - 4 символа
                                                                                    //MessageBox.Show(pYear.ToString());
                            pMonth2 = Convert.ToInt32(peremennaya2.Substring(3, 2)); // парсим peremennaya2 с 4го символа, длина - 2 символа
                                                                                     //MessageBox.Show(pMonth.ToString());
                            pDay2 = Convert.ToInt32(peremennaya2.Substring(0, 2)); // парсим peremennaya2 с 1го символа, длина - 2 символа
                                                                                   //MessageBox.Show(pDay.ToString());
                            DateTime proverka2 = new DateTime(pYear2, pMonth2, pDay2);
                            currRow.Cells[IndexNextRankDate].Value = proverka2.AddYears(NumberOfYears).ToString("dd.MM.yyyy");
                        }
                    }
                }
            }
        }


        // ###############  ОСНОВНОЙ МЕТОД ПРОВЕРКИ И ПЕРЕСЧЁТА КВАЛИФИКАЦИОННЫХ ЗВАНИЙ  ###############
        private void PereschetKlassnost()
        {
            if (dataGridView1.Rows.Count != 0) // Проверка dataGridView1 на пустоту
            {
                int NumberOfYearsKlassnost = 3; // срок выслуги до следующего квалификационного звания
                int pYear3, pMonth3, pDay3; // переменные для парсинга текстовой даты в год, месяц, день
                string peremennaya3; // переменная для хранения даты из ячеек

                foreach (DataGridViewRow currRow in dataGridView1.Rows) // проходим по каждой строке в таблице
                {
                    string elemStr = currRow.Cells[IndexKlassnost].Value.ToString(); // Переменная с текущим значением классности

                    // Если классность отсутствует
                    if (elemStr == "Отсутствует") 
                    {
                        currRow.Cells[IndexKlassnostDate].Value = "--.--.----"; // убираем дату присвоения классности
                        currRow.Cells[IndexNextKlassnostDate].Value = "--.--.----"; // убираем следующую дату присвоения классности
                    }

                    // Если сотрудник является специалистом 3, 2 или 1 класса
                    if ((elemStr == "Специалист 3 класса") || (elemStr == "Специалист 2 класса") || (elemStr == "Специалист 1 класса")) //(KlassnostCompare == 3)
                    {
                        if (currRow.Cells[IndexKlassnostDate].Value.ToString() == "--.--.----") currRow.Cells[IndexKlassnostDate].Value = DateTime.Now.ToString("dd.MM.yyyy");
                        peremennaya3 = currRow.Cells[IndexKlassnostDate].Value.ToString(); // считываем значение даты из ячейки в peremennaya2 типа string
                                                                                            //MessageBox.Show(peremennaya3);
                        pYear3 = Convert.ToInt32(peremennaya3.Substring(6, 4)); // парсим peremennaya2 с 7го символа, длина - 4 символа
                                                                                //MessageBox.Show(pYear.ToString());
                        pMonth3 = Convert.ToInt32(peremennaya3.Substring(3, 2)); // парсим peremennaya2 с 4го символа, длина - 2 символа
                                                                                 //MessageBox.Show(pMonth.ToString());
                        pDay3 = Convert.ToInt32(peremennaya3.Substring(0, 2)); // парсим peremennaya2 с 1го символа, длина - 2 символа
                                                                               //MessageBox.Show(pDay.ToString());
                        DateTime proverka3 = new DateTime(pYear3, pMonth3, pDay3);
                        currRow.Cells[IndexNextKlassnostDate].Value = proverka3.AddYears(NumberOfYearsKlassnost).ToString("dd.MM.yyyy");
                    }

                    // Если сотрудник имеет высшее квалификационное звание "Мастер"
                    if (elemStr == "Мастер") 
                    {
                        if (currRow.Cells[IndexKlassnostDate].Value.ToString() == "--.--.----") // Если дата присвоения отсутствует... 
                        {
                            currRow.Cells[IndexKlassnostDate].Value = DateTime.Now.ToString("dd.MM.yyyy"); // ...ставим текущую дату
                        }
                        currRow.Cells[IndexNextKlassnostDate].Value = "высшее звание"; // в ячейке со следующей датой пишем "высшее звание"
                            
                    }
                }
            }
        }

        // ###############  ФИЛЬТРАЦИЯ  ############### 

        private void ShowAllRows() // показать все строки dataGridView1
        {
            dataGridView1.CurrentCell = null;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.Visible = true; // показываем строку
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (SomeRowsWasHidden == 1) // Если есть скрытые строки
            {
                this.ShowAllRows(); // Показываем все строки
                this.PereschetCnum(); // Пересчитываем порядковые номера строк
                SomeRowsWasHidden = 0; // Переводим маркер в состояние "Нет скрытых строк"
            }
            this.ShowAllColumns(); // Показываем "стандартные" колонки dataGridView1

        }

        private void ShowAllColumns() // показать все "стандартные" колонки dataGridView1
        {
            personalfilenum.Visible = false;
            personalnum.Visible = false;
            gender.Visible = false;
            dateofbirth.Visible = true;
            placeofbirth.Visible = false;
            registration.Visible = false;
            placeofliving.Visible = false;
            phoneregistration.Visible = false;
            phoneplaceofliving.Visible = false;
            rank.Visible = true;
            rankdate.Visible = true;
            ranklimit.Visible = true;
            nextrankdate.Visible = true;
            klassnost.Visible = true;
            klassnostdate.Visible = true;
            nextklassnostdate.Visible = true;
            study.Visible = false;
            uchstepen.Visible = false;
            prisvzvaniy.Visible = false;
            married.Visible = false;
            family.Visible = false;
            truddeyat.Visible = false;
            stazhvysluga.Visible = false;
            dataprisyagi.Visible = false;
            rabotagfs.Visible = false;
            attestaciya.Visible = false;
            nextattestaciyadate.Visible = false;
            profpodg.Visible = false;
            klassnostcheyprikaz.Visible = false;
            klassnostnomerprikaza.Visible = false;
            klassnostold.Visible = false;
            nagrady.Visible = false;
            prodlenie.Visible = false;
            boevye.Visible = false;
            rezerv.Visible = false;
            vzyskaniya.Visible = false;
            uvolnenie.Visible = false;
            zapolnil.Visible = false;
            datazapolneniya.Visible = false;
            imagestring.Visible = false;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (SomeRowsWasHidden == 1) // Если есть скрытые строки
            {
                this.ShowAllRows(); // Показываем все строки
                this.PereschetCnum(); // Пересчитываем порядковые номера строк
                SomeRowsWasHidden = 0; // Переводим маркер в состояние "Нет скрытых строк"
            }
            this.ShowVysluga(); // Показываем колонки dataGridView1, связанные со специальным званием
        }

        private void ShowVysluga() // показать только выслугу
        {
            personalfilenum.Visible = false;
            personalnum.Visible = false;
            gender.Visible = false;
            dateofbirth.Visible = false;
            placeofbirth.Visible = false;
            registration.Visible = false;
            placeofliving.Visible = false;
            phoneregistration.Visible = false;
            phoneplaceofliving.Visible = false;
            rank.Visible = true;
            rankdate.Visible = true;
            ranklimit.Visible = true;
            nextrankdate.Visible = true;
            klassnost.Visible = false;
            klassnostdate.Visible = false;
            nextklassnostdate.Visible = false;
            study.Visible = false;
            uchstepen.Visible = false;
            prisvzvaniy.Visible = false;
            married.Visible = false;
            family.Visible = false;
            truddeyat.Visible = false;
            stazhvysluga.Visible = false;
            dataprisyagi.Visible = false;
            rabotagfs.Visible = false;
            attestaciya.Visible = false;
            nextattestaciyadate.Visible = false;
            profpodg.Visible = false;
            klassnostcheyprikaz.Visible = false;
            klassnostnomerprikaza.Visible = false;
            klassnostold.Visible = false;
            nagrady.Visible = false;
            prodlenie.Visible = false;
            boevye.Visible = false;
            rezerv.Visible = false;
            vzyskaniya.Visible = false;
            uvolnenie.Visible = false;
            zapolnil.Visible = false;
            datazapolneniya.Visible = false;
            imagestring.Visible = false;
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (SomeRowsWasHidden == 1) // Если есть скрытые строки
            {
                this.ShowAllRows(); // Показываем все строки
                this.PereschetCnum(); // Пересчитываем порядковые номера строк
                SomeRowsWasHidden = 0; // Переводим маркер в состояние "Нет скрытых строк"
            }
            this.ShowKlassnost(); // Показываем колонки dataGridView1, связанные с классным званием
        }

        private void ShowKlassnost() // Показать только классность
        {
            personalfilenum.Visible = false;
            personalnum.Visible = false;
            gender.Visible = false;
            dateofbirth.Visible = false;
            placeofbirth.Visible = false;
            registration.Visible = false;
            placeofliving.Visible = false;
            phoneregistration.Visible = false;
            phoneplaceofliving.Visible = false;
            rank.Visible = false;
            rankdate.Visible = false;
            ranklimit.Visible = false;
            nextrankdate.Visible = false;
            klassnost.Visible = true;
            klassnostdate.Visible = true;
            nextklassnostdate.Visible = true;
            study.Visible = false;
            uchstepen.Visible = false;
            prisvzvaniy.Visible = false;
            married.Visible = false;
            family.Visible = false;
            truddeyat.Visible = false;
            stazhvysluga.Visible = false;
            dataprisyagi.Visible = false;
            rabotagfs.Visible = false;
            attestaciya.Visible = false;
            nextattestaciyadate.Visible = false;
            profpodg.Visible = false;
            klassnostcheyprikaz.Visible = false;
            klassnostnomerprikaza.Visible = false;
            klassnostold.Visible = false;
            nagrady.Visible = false;
            prodlenie.Visible = false;
            boevye.Visible = false;
            rezerv.Visible = false;
            vzyskaniya.Visible = false;
            uvolnenie.Visible = false;
            zapolnil.Visible = false;
            datazapolneniya.Visible = false;
            imagestring.Visible = false;
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            this.ShowAttestaciya();
        }

        private void ShowAttestaciya() // показать только тех, у кого аттестация в следующем году (проводится каждые 4 года)
        {
            if (radioButton4.Checked == true) //без этой проверки, при radioButton1.Checked = true, метод срабатывает второй раз (позже разобраться почему)
            {
                int NumberOfHiddenRows = 0; // Переменная для храненения количества скрытых строк
                int VisibleCnum = 0; // Переменная для пересчета порядковых номеров в видимых строках
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (dataGridView1[IndexNextAttestaciyaDate, row.Index].Value.ToString() != "") // если значение даты аттестации в строке не пустое
                    {
                        DateTime DateFromCurrentRow = DateTime.Parse(dataGridView1[IndexNextAttestaciyaDate, row.Index].Value.ToString()); // парсим значение даты аттестации в формат DateTime
                        DateTime CheckedYear = DateTime.Now.AddYears(1); // прибавляем к текущей дате один год в формат DateTime

                        if (DateFromCurrentRow.Year != CheckedYear.Year) // Если год из ячейки с датой следующей аттестации не является следующим годом
                        {
                            dataGridView1.CurrentCell = null;
                            row.Visible = false; // скрываем строку
                            NumberOfHiddenRows++;
                        }
                        else 
                        {
                            VisibleCnum++;
                            dataGridView1[IndexCnum, row.Index].Value = VisibleCnum;
                        }
                            
                    }
                    else
                    {
                        dataGridView1.CurrentCell = null;
                        row.Visible = false; // скрываем строку
                        NumberOfHiddenRows++;
                    }
                }
                if (NumberOfHiddenRows == dataGridView1.Rows.Count) // Если по итогу перебора строк, все строки оказались скрыты
                {
                    MessageBox.Show("Сотрудники, подпадающие под данный фильтр отсутствуют!");
                    radioButton1.Checked = true; // сбрасываем выбор фильтра
                    this.ShowAllRows(); // Показываем все строки
                }
                else // если есть, что выводить - скрываем лишние столбцы
                {
                    personalfilenum.Visible = false;
                    personalnum.Visible = false;
                    gender.Visible = false;
                    dateofbirth.Visible = true;
                    placeofbirth.Visible = false;
                    registration.Visible = false;
                    placeofliving.Visible = false;
                    phoneregistration.Visible = false;
                    phoneplaceofliving.Visible = false;
                    rank.Visible = false;
                    rankdate.Visible = false;
                    ranklimit.Visible = false;
                    nextrankdate.Visible = false;
                    klassnost.Visible = false;
                    klassnostdate.Visible = false;
                    nextklassnostdate.Visible = false;
                    study.Visible = false;
                    uchstepen.Visible = false;
                    prisvzvaniy.Visible = false;
                    married.Visible = false;
                    family.Visible = false;
                    truddeyat.Visible = false;
                    stazhvysluga.Visible = false;
                    dataprisyagi.Visible = false;
                    rabotagfs.Visible = false;
                    attestaciya.Visible = false;
                    nextattestaciyadate.Visible = true;
                    profpodg.Visible = false;
                    klassnostcheyprikaz.Visible = false;
                    klassnostnomerprikaza.Visible = false;
                    klassnostold.Visible = false;
                    nagrady.Visible = false;
                    prodlenie.Visible = false;
                    boevye.Visible = false;
                    rezerv.Visible = false;
                    vzyskaniya.Visible = false;
                    uvolnenie.Visible = false;
                    zapolnil.Visible = false;
                    datazapolneniya.Visible = false;
                    imagestring.Visible = false;

                    //this.PereschetCnum(); // Пересчитываем порядковые номера строк
                    SomeRowsWasHidden = 1; // Переводим маркер в состояние "Есть скрытые строки"
                }
            }
        }


        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Tab && dataGridView1.CurrentCell.ColumnIndex == 1)
            {
                e.Handled = true;
                DataGridViewCell cell = dataGridView1.Rows[0].Cells[0];
                dataGridView1.CurrentCell = cell;
                dataGridView1.BeginEdit(true);
            }
        }





        // #################################################
        // ##  КНОПКА "ЗАКРЫТЬ" НА ВКЛАДКЕ "ОБЩИЙ СПИСОК" ##
        // #################################################
        private void Close1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // ####################################
        // ##  КНОПКА "ДОБАВИТЬ СОТРУДНИКА"  ##
        // ####################################
        private void AddPerson_Click(object sender, EventArgs e)
        {
            int id;

            if (dataGridView1.Rows.Count == 0) // Проверка dataGridView1 на пустоту
            {
                id = 0;
            }
            else
            {
                this.PereschetCnum();
                id = Convert.ToInt32(dataGridView1[IndexCnum, dataGridView1.RowCount - 1].Value); //присваиваем переменной ID последний порядковый номер
            }

            dataSet1.Tables[CurrentDataTableName].Rows.Add(id + 1, "не указан", "не указан", "Фамилия", "Имя", "Отчество", "М", DateTime.Now.ToString("dd.MM.yyyy")/* Дата рождения */,
                "Город", "Прописка", "Адрес проживания", "не указан"/* Телефон 1 */, "не указан"/* Телефон 2 */, "Должность", "Рядовой"/* Спец. звание */,
                DateTime.Now.ToString("dd.MM.yyyy")/* Дата присвоения звания */, "Рядовой"/* Потолок по званию */, "роста нет"/* Дата след. звания */,
                "Отсутствует"/* Классность */, "--.--.----"/* Дата классности */, "--.--.----"/* След. дата классности */,
                ""/* 10.Образование */,
                ""/* 11. Ученая степень */, ""/* 12.Присвоение званий, чинов */,
                ""/* 13.Семейное положение */, ""/* 14.Члены семьи */, ""/* 15.Труд. деят. до прихода */,
                "Общий трудовой стаж^0^0^0$Льготная выслуга^0^0^0$Стаж для государственных служащих^0^0^0$Половина периода обучения в высш. и сред. спец. учебных заведениях (для лиц начальствующего состава)^0^0^0$Календарная выслуга^0^0^0"/* 16.Стаж и выслуга до прихода */,
                DateTime.Now.ToString("dd.MM.yyyy")/* 17.Дата принятия присяги */,
                ""/* 18.Прохождение службы (работа) в ГФС России */,
                ""/* 19.Аттестация */,
                ""/* 20.Дата следующей аттестации */,
                ""/* 21.Профессиональная подготовка */,
                "---"/* 22.Чей приказ о присвоении квалиф. звания */, "---"/* 23.Дата приказа о присвоении квалиф. звания */, ""/* 24.Сведения о присвоенных ранее квалиф. званиях  */,
                ""/* 25.Награды и поощрения */, ""/* 26.Продление службы */, ""/* 27.Участие в боевых действиях */, ""/* 28.Состояние в резерве */,
                ""/* 29.Взыскания */, ""/* 30.Увольнение */, ""/* 31.Карточку заполнил */, DateTime.Now.ToString("dd.MM.yyyy")/* 32.Дата заполнения карточки */,
                ""/* 33.Фото */);
            this.AcceptAndWriteChanges();
            Cnum_label.Text = (IndexRowLichnayaKarta + 1).ToString() + " из " + dataGridView1.RowCount.ToString(); // Порядковый номер личной карточки         
        }

        // ###########################################################
        // ##  КНОПКА "АРХИВНЫЕ СОТРУДНИКИ/ДЕЙСТВУЮЩИЕ СОТРУДНИКИ"  ##
        // ###########################################################
        private void Archive_Click(object sender, EventArgs e)
        {
            if (CurrentDataTableName == "Kadry")
            {
                Archive.Text = "Действующие сотрудники";
                CurrentDataTableName = "Archive";
                OtherDataTableName = "Kadry";
                CurrentBase_label.Text = "Текущая БД: Архивные сотрудники";
                MessageBox.Show("Перешли в архивную базу данных");
            }
            else if (CurrentDataTableName == "Archive")
            {
                Archive.Text = "Архивные сотрудники";
                CurrentDataTableName = "Kadry";
                OtherDataTableName = "Archive";
                CurrentBase_label.Text = "Текущая БД: Действующие сотрудники";
                MessageBox.Show("Перешли в текущую базу данных");
            }
            this.RefreshDataGridView1(); // обновляем DataGridView1

            if (dataGridView1.Rows.Count != 0) // Проверка dataGridView1 на пустоту
            {
                IndexRowLichnayaKarta = 0;
                this.PereschetZvanie();
                this.PereschetKlassnost();
                this.PereschetCnum(); // пересчитываем порядковые номера 
                Cnum_label.Text = (IndexRowLichnayaKarta + 1).ToString() + " из " + dataGridView1.RowCount.ToString(); // Порядковый номер личной карточки
                CardsFIO_label.Text = dataGridView1[IndexSurname, IndexRowLichnayaKarta].Value.ToString() + " "
        + dataGridView1[IndexName, IndexRowLichnayaKarta].Value.ToString() + " "
        + dataGridView1[IndexMiddleName, IndexRowLichnayaKarta].Value.ToString(); // Прописываем ФИО над стрелками в карточке
            }
            // При необходимости, добавить сюда события при пустом гриде
        }


        // ##################################
        // ##  КНОПКА "ВЫГРУЗИТЬ В EXCEL"  ##
        // ##################################
        private void ExportToExcel_Click(object sender, EventArgs e)
        {
            this.ExportDataGridToExcel();
        }

        // ###############  ВЫГРУЗКА dataGridView1 В EXCEL ФАЙЛ  ###############
        public void ExportDataGridToExcel()
        {
            //Формируем новый список listVisibleColumns, состоящий только из видимых столбцов
            List<DataGridViewColumn> listVisibleColumns = new List<DataGridViewColumn>();
            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                if (col.Visible)
                {
                    listVisibleColumns.Add(col);
                }
            }

            //Формируем новый список listVisibleRows, состоящий только из видимых строк
            List<DataGridViewRow> listVisibleRows = new List<DataGridViewRow>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Visible)
                {
                    listVisibleRows.Add(row);
                }
            }

            /*==============================================================================================================*/

            // Подготавливаем Excel для экспорта dataGridView1
            Excel.Application ExcelApp = new Excel.Application();
            ExcelApp.Application.Workbooks.Add(Type.Missing); //Создаем новую книгу
            ExcelApp.Columns.ColumnWidth = 15; // устанавливаем ширину столбцов
            ExcelApp.Cells.WrapText = "true"; // устанавливаем перенос по словам

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)ExcelApp.Worksheets.get_Item(1); //Создаем новый лист
            xlWorkSheet.Name = "Сведения о личном составе"; // именуем лист

            var _with1 = xlWorkSheet.PageSetup; // Блок параметров листа
            _with1.PaperSize = Excel.XlPaperSize.xlPaperA4; // размер А4
            _with1.Orientation = Excel.XlPageOrientation.xlLandscape; // ландшафтная ориентация
            _with1.Zoom = false;
            // Ужимаем всё при выводе на печать
            _with1.FitToPagesWide = 1;
            _with1.FitToPagesTall = 1;

            /*==============================================================================================================*/

            // Заполняем заголовки
            for (int i = 0; i < listVisibleColumns.Count; i++) // Проходим только по видимым столбцам
            {
                ExcelApp.Cells[1, i + 1] = listVisibleColumns[i].HeaderText; // Заполняем первую строку Excel заголовками видимых столбцов
            }

            // Украшаем заголовки
            Excel.Range range_zagolovki = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, listVisibleColumns.Count]); // диапазон заголовка в файле Excel
            range_zagolovki.Cells.Font.Bold = true; // жирный шрифт
            range_zagolovki.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // горизонтальное выравнивание по центру
            range_zagolovki.BorderAround(Type.Missing, Excel.XlBorderWeight.xlThick, Excel.XlColorIndex.xlColorIndexAutomatic); // увеличиваем толщину внешних границ
            range_zagolovki.Borders.Color = Color.Black; // черный цвет границ
            range_zagolovki.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter; // вертикальное выравнивание по центру

            /*==============================================================================================================*/

            //Заполняем лист Excel видимыми строками, столбцами и центрируем столбцы с датами
            for (int col = 0; col < listVisibleColumns.Count; col++)
            {
                if (listVisibleColumns[col] is CalendarColumn) // Центрируем столбцы с датами в Excel
                {
                    Excel.Range range_col_with_date = xlWorkSheet.get_Range(xlWorkSheet.Cells[2, col + 1], xlWorkSheet.Cells[listVisibleRows.Count + 1, col + 1]); // диапазон столбца, где обнаружена дата (без заголовка)
                    range_col_with_date.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // горизонтальное выравнивание по центру
                }
                for (int row = 0; row < listVisibleRows.Count; row++)
                {
                    ExcelApp.Cells[row + 2, col + 1] = dataGridView1.Rows[listVisibleRows[row].Index].Cells[listVisibleColumns[col].Index].Value.ToString(); // Наполняем лист Excel видимыми ячейками, начиная с первой строки после заголовка
                }
            }

            //Украшаем все, кроме заголовка
            Excel.Range range_all_cells_without_headers = xlWorkSheet.get_Range(xlWorkSheet.Cells[2, 1], xlWorkSheet.Cells[listVisibleRows.Count + 1, listVisibleColumns.Count]); // диапазон всех ячеек, кроме заголовка
            range_all_cells_without_headers.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter; // вертикальное выравнивание по центру
            range_all_cells_without_headers.Borders[Excel.XlBordersIndex.xlInsideVertical].Color = Color.LightGray; //внутренние вертикальные границы области с данными
            range_all_cells_without_headers.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Color = Color.Black; //внутренние горизонтальные границы области с данными
            range_all_cells_without_headers.Borders[Excel.XlBordersIndex.xlEdgeRight].Color = Color.Black; //крайняя правая граница области с данными
            range_all_cells_without_headers.Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = Color.Black; //крайняя левая граница области с данными
            range_all_cells_without_headers.Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = Color.Black; //крайняя нижняя граница области с данными

            Excel.Range range_Cnum = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[listVisibleRows.Count + 1, 1]); // Диапазон ячеек с порядковым номером
            range_Cnum.ColumnWidth = 5; // уменьшаем ширину
            range_Cnum.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // горизонтальное выравнивание по центру

            /*
                        // В ДАЛЬНЕЙШЕМ ВОЗМОЖНО ПОНАДОБИТСЯ ИМЕНОВАТЬ ЛИСТ EXCEL В ЗАВИСИМОСТИ ОТ ВЫБРАННОГО ФИЛЬТРА
                        if (radioButton1.Checked == true) //если выбран определенный фильтр
                        {
                            xlWorkSheet.Name = "Название листа"; // именуем лист
                        }
            */

            ExcelApp.Visible = true; // Показываем Excel
        }


        // #########################################################
        // ##  КНОПКА "ПОДСЧЕТ ВЫСЛУГИ" НА ВКЛАДКЕ "ОБЩИЙ СПИСОК" ##
        // #########################################################
        private void VyslugaCalc_Click(object sender, EventArgs e)
        {

        }

        // ###############  ДЕЙСТВИЯ ПРИ СРАБАТЫВАНИИ СОБЫТИЯ СОРТИРОВКИ  ###############
        private void dataGridView1_Sorted(object sender, EventArgs e) //отработка события изменения сортировки
        {
            this.PereschetCnum();
        }

        // ###############  ДЕЙСТВИЯ, ЕСЛИ БЫЛИ КАКИЕ-ЛИБО ИЗМЕНЕНИЯ В dataGridView1  ###############
        public void DataGridWasChanged()
        {
            // MessageBox.Show("grid изменен");  // позже будет закомментировано 
            this.PereschetZvanie(); // пересчитываем звание
            this.PereschetKlassnost(); // пересчитываем классность
            this.AcceptAndWriteChanges(); // сохраняем изменения в XML
            this.RefreshDataGridView1(); // обновляем DataGridView1
        }

        // ###############  ОБНОВЛЕНИЕ dataGridView1  ###############
        public void RefreshDataGridView1()
        {
            dataSet1.Clear(); // очищаем dataSet1
            dataGridView1.DataSource = null; // очищаем DataSource
            dataSet1.ReadXml(XMLDB.Path); // считываем XML
            dataGridView1.DataSource = dataSet1.Tables[CurrentDataTableName]; // присваиваем DataSource
        }

        // ###############  ПРИМЕНИТЬ ВСЕ ИЗМЕНЕНИЯ И СОХРАНИТЬ XML  ###############
        public void AcceptAndWriteChanges()
        {
            // MessageBox.Show("Произошло сохранение базы данных"); // позже будет закомментировано 
            dataSet1.AcceptChanges(); // применяем изменения в dataSet1
            dataSet1.WriteXml(XMLDB.Path); // сохраняем изменения в XML          
        }


        // ###############  НАЧАЛО РЕДАКТИРОВАНИЯ ЯЧЕЙКИ  ###############
        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            CellValueToCompare = dataGridView1.CurrentCell.Value.ToString(); // присваиваем переменной CellValueToCompare текущее значение ячейки до редактирования
            LastEditedCellRow = dataGridView1.CurrentCell.RowIndex;
            LastEditedCellCol = dataGridView1.CurrentCell.ColumnIndex;
        }

        // ###############  ЗАВЕРШЕНИЕ РЕДАКТИРОВАНИЯ ЯЧЕЙКИ  ###############
        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentCell.Value.ToString() != CellValueToCompare) // сравниваем CellValueToCompare со значением ячейки после редактирования
            {
                DataGridWasChanged();
                dataGridView1.CurrentCell = dataGridView1[LastEditedCellCol, LastEditedCellRow];
            }
        }

        // ###############  ДЕЙСТВИЯ ПРИ СРАБАТЫВАНИИ СОБЫТИЯ RowDeleting (ПЕРЕД УДАЛЕНИЕМ СТРОКИ)  ###############
        private void RowDeleting(object sender, DataRowChangeEventArgs e)
        {
            if (tabControl1.SelectedTab.Text == "Общий список")
            {
                var result = MessageBox.Show("Удалить данную запись?", "Вы уверены?",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (result == DialogResult.No) // если была нажать кнопка "Нет"
                {
                    WantToDeleteRow = 0; // сбрасываем маркер удаления строки в ноль
                    dataSet1.Tables[CurrentDataTableName].RejectChanges(); // отменяем изменения
                    this.RefreshDataGridView1(); // обновляем DataGridView1
                }
                else
                {
                    WantToDeleteRow = 1; // маркер, что пользователь все-таки хочет удалить строку
                }
            }
        }

        // ###############  ДЕЙСТВИЯ ПРИ СРАБАТЫВАНИИ СОБЫТИЯ RowDeleted (ПОСЛЕ УДАЛЕНИЯ СТРОКИ)  ###############
        private void RowDeleted(object sender, DataRowChangeEventArgs e)
        {
            if (WantToDeleteRow == 1) // если пользователь хочет удалить строку
            {
                this.AcceptAndWriteChanges(); // сохраняем изменения
                WantToDeleteRow = 0; // сбрасываем маркер удаления строки в ноль

                if (dataGridView1.Rows.Count != 0) // Проверка dataGridView1 на пустоту
                {
                    IndexRowLichnayaKarta = 0;
                    this.PereschetCnum(); // пересчитываем порядковые номера 
                    Cnum_label.Text = (IndexRowLichnayaKarta + 1).ToString() + " из " + dataGridView1.RowCount.ToString(); // Порядковый номер личной карточки
                    CardsFIO_label.Text = dataGridView1[IndexSurname, IndexRowLichnayaKarta].Value.ToString() + " "
            + dataGridView1[IndexName, IndexRowLichnayaKarta].Value.ToString() + " "
            + dataGridView1[IndexMiddleName, IndexRowLichnayaKarta].Value.ToString(); // Прописываем ФИО над стрелками в карточке
                }
                // Позже, при необходимости, описать события для ситуации, когда таблица остается пустой

            }
        }

        // ###############  СОБЫТИЕ, ПРИ СМЕНЕ АКТИВНОЙ ВКЛАДКИ ###############
        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (dataGridView1.Rows.Count == 0) // Проверка dataGridView1 на пустоту. Если грид пустой - не даем уйти с вкладки "Общий список"
            {
                e.Cancel = true;
                MessageBox.Show("Сначала добавьте хотя бы одного сотрудника!");
            }
        }
        // ###############  СОБЫТИЕ, ПОСЛЕ СМЕНЫ АКТИВНОЙ ВКЛАДКИ ###############
        private void tabControl1_SelectedIndexChanged(Object sender, EventArgs e)
        {
            if (IndexRowLichnayaKarta > dataGridView1.RowCount - 1) // Проверка на выход за пределы диапазона личных карточек.
                                                                    // Такое может произойти, если была активной последняя карточка,
                                                                    // после чего её удалили и снова "вышли" из "Общего списка"
            {
                IndexRowLichnayaKarta = 0;
                Cnum_label.Text = (IndexRowLichnayaKarta + 1).ToString() + " из " + dataGridView1.RowCount.ToString(); // Порядковый номер личной карточки
                CardsFIO_label.Text = dataGridView1[IndexSurname, IndexRowLichnayaKarta].Value.ToString() + " "
        + dataGridView1[IndexName, IndexRowLichnayaKarta].Value.ToString() + " "
        + dataGridView1[IndexMiddleName, IndexRowLichnayaKarta].Value.ToString(); // Прописываем ФИО над стрелками в карточке
            }
            this.NeedToUpdateCard();
        }

        // ###############  ОПРЕДЕЛЯЕМ, КАКУЮ ВКЛАДКУ НУЖНО ОБНОВИТЬ ###############
        public void NeedToUpdateCard()
        {
            if (tabControl1.SelectedTab.Text == "Общий список") // скрыть нижнюю панель со стрелками на вкладке "Общий список" и показать все колонки
            {
                Cards_groupBox.Visible = false; // скрываем панель со стрелками
                radioButton1.Checked = true; // сбрасываем выбор фильтра
            }
            else Cards_groupBox.Visible = true; // отображаем панель со стрелками

            switch (tabControl1.SelectedTab.Text) // сверяем название активной вкладки
            {
                case "Карточка 1-9": // 
                    this.UpdateCard1to9();
                    break;
                case "Карточка 10-11": // 
                    this.UpdateCard10and11();
                    break;
                case "Карточка 12": //
                    this.UpdateCard12();
                    break;
                case "Карточка 13-14": // 
                    this.UpdateCard13and14();
                    break;
                case "Карточка 15": // 
                    this.UpdateCard15();
                    break;
                case "Карточка 16-18": // 
                    this.UpdateCard16to18();
                    break;
                case "Карточка 19-20": // 
                    this.UpdateCard19and20();
                    break;
                case "Карточка 21-22": // 
                    this.UpdateCard21and22();
                    break;
                case "Карточка 23-25": // 
                    this.UpdateCard23to25();
                    break;
                case "Карточка 26-29": // 
                    this.UpdateCard26to29();
                    break;
            }
        }


        //               //""""""""""""""""""""""""\\ 
        // ###############  ВКЛАДКА "КАРТОЧКА 1-9"  ############################################################
        public void UpdateCard1to9()
        {
            Card1to9WasLoaded = 0;

            if (dataGridView1[IndexImageString, IndexRowLichnayaKarta].Value.ToString() == "") // Если картинка отсутствует
            {
                dataGridView1[IndexImageString, IndexRowLichnayaKarta].Value = XMLDB.DefaultImageBase64; // присваиваем pictureBox1 стандартную картинку с крестиком
                Bitmap bmp = new Bitmap(new MemoryStream(Convert.FromBase64String(dataGridView1[IndexImageString, IndexRowLichnayaKarta].Value.ToString()))); // собираем изображение
                pictureBox1.Image = bmp;
            }
            else
            {
                Bitmap bmp = new Bitmap(new MemoryStream(Convert.FromBase64String(dataGridView1[IndexImageString, IndexRowLichnayaKarta].Value.ToString()))); // собираем изображение
                pictureBox1.Image = bmp; //присваиваем pictureBox1 собранную ячейку
            }

            // ЗАПОЛНЯЕМ textBox'ы:
            PersonalFileNum_textBox.Text = dataGridView1[IndexPersonalFileNum, IndexRowLichnayaKarta].Value.ToString();
            PersonalNum_textBox.Text = dataGridView1[IndexPersonalNum, IndexRowLichnayaKarta].Value.ToString();
            Surname_textBox.Text = dataGridView1[IndexSurname, IndexRowLichnayaKarta].Value.ToString();
            Name_textBox.Text = dataGridView1[IndexName, IndexRowLichnayaKarta].Value.ToString();
            MiddleName_textBox.Text = dataGridView1[IndexMiddleName, IndexRowLichnayaKarta].Value.ToString();
            Gender_comboBox.Text = dataGridView1[IndexGender, IndexRowLichnayaKarta].Value.ToString();
            DateOfBirth_dateTimePicker.Text = dataGridView1[IndexDateOfBirth, IndexRowLichnayaKarta].Value.ToString();
            RankDate_dateTimePicker.Text = dataGridView1[IndexRankDate, IndexRowLichnayaKarta].Value.ToString();
            PlaceOfBirth_textBox.Text = dataGridView1[IndexPlaceOfBirth, IndexRowLichnayaKarta].Value.ToString();
            Registration_textBox.Text = dataGridView1[IndexRegistration, IndexRowLichnayaKarta].Value.ToString();
            PlaceOfLiving_textBox.Text = dataGridView1[IndexPlaceOfLiving, IndexRowLichnayaKarta].Value.ToString();
            PhoneRegistration_textBox.Text = dataGridView1[IndexPhoneRegistration, IndexRowLichnayaKarta].Value.ToString();
            PhonePlaceOfLiving_textBox.Text = dataGridView1[IndexPhonePlaceOfLiving, IndexRowLichnayaKarta].Value.ToString();
            Post_textBox.Text = dataGridView1[IndexPost, IndexRowLichnayaKarta].Value.ToString();
            NextRankDate_textBox.Text = dataGridView1[IndexNextRankDate, IndexRowLichnayaKarta].Value.ToString();

            Rank_comboBox.BindingContext = new BindingContext();   //создаем новый контекст, иначе в определенный момент получаем null в одном из comboBox'ов
            Rank_comboBox.DataSource = ZvanieList;
            Rank_comboBox.Text = dataGridView1[IndexRank, IndexRowLichnayaKarta].Value.ToString();
            RankLimit_comboBox.BindingContext = new BindingContext();   //создаем новый контекст, иначе в определенный момент получаем null в одном из comboBox'ов
            RankLimit_comboBox.DataSource = ZvanieList;
            RankLimit_comboBox.Text = dataGridView1[IndexRankLimit, IndexRowLichnayaKarta].Value.ToString();

            Card1to9WasLoaded = 1; // карточка прогрузилась
        }


        // ##################################################################################
        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В TextBox'ах НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        // ##################################################################################

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В Surname_textBox НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void Surname_textBox_Leave(object sender, EventArgs e)
        {
            if (Surname_textBox.Text != dataGridView1[IndexSurname, IndexRowLichnayaKarta].Value.ToString())
            {
                dataGridView1[IndexSurname, IndexRowLichnayaKarta].Value = Surname_textBox.Text;
                this.AcceptAndWriteChanges(); // применить изменения
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В Name_textBox НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void Name_textBox_Leave(object sender, EventArgs e)
        {
            if (Name_textBox.Text != dataGridView1[IndexName, IndexRowLichnayaKarta].Value.ToString())
            {
                dataGridView1[IndexName, IndexRowLichnayaKarta].Value = Name_textBox.Text;
                this.AcceptAndWriteChanges(); // применить изменения
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В MiddleName_textBox НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void MiddleName_textBox_Leave(object sender, EventArgs e)
        {
            if (MiddleName_textBox.Text != dataGridView1[IndexMiddleName, IndexRowLichnayaKarta].Value.ToString())
            {
                dataGridView1[IndexMiddleName, IndexRowLichnayaKarta].Value = MiddleName_textBox.Text;
                this.AcceptAndWriteChanges(); // применить изменения
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В PlaceOfBirth_textBox НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void PlaceOfBirth_textBox_Leave(object sender, EventArgs e)
        {
            if (PlaceOfBirth_textBox.Text != dataGridView1[IndexPlaceOfBirth, IndexRowLichnayaKarta].Value.ToString())
            {
                dataGridView1[IndexPlaceOfBirth, IndexRowLichnayaKarta].Value = PlaceOfBirth_textBox.Text;
                this.AcceptAndWriteChanges(); // применить изменения
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В PersonalFileNum_textBox НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void PersonalFileNum_textBox_Leave(object sender, EventArgs e)
        {
            if (PersonalFileNum_textBox.Text != dataGridView1[IndexPersonalFileNum, IndexRowLichnayaKarta].Value.ToString())
            {
                dataGridView1[IndexPersonalFileNum, IndexRowLichnayaKarta].Value = PersonalFileNum_textBox.Text;
                this.AcceptAndWriteChanges(); // применить изменения
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В PersonalNum_textBox НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void PersonalNum_textBox_Leave(object sender, EventArgs e)
        {
            if (PersonalNum_textBox.Text != dataGridView1[IndexPersonalNum, IndexRowLichnayaKarta].Value.ToString())
            {
                dataGridView1[IndexPersonalNum, IndexRowLichnayaKarta].Value = PersonalNum_textBox.Text;
                this.AcceptAndWriteChanges(); // применить изменения
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В Registration_textBox НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void Registration_textBox_Leave(object sender, EventArgs e)
        {
            if (Registration_textBox.Text != dataGridView1[IndexRegistration, IndexRowLichnayaKarta].Value.ToString())
            {
                dataGridView1[IndexRegistration, IndexRowLichnayaKarta].Value = Registration_textBox.Text;
                this.AcceptAndWriteChanges(); // применить изменения
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В PlaceOfLiving_textBox НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void PlaceOfLiving_textBox_Leave(object sender, EventArgs e)
        {
            if (PlaceOfLiving_textBox.Text != dataGridView1[IndexPlaceOfLiving, IndexRowLichnayaKarta].Value.ToString())
            {
                dataGridView1[IndexPlaceOfLiving, IndexRowLichnayaKarta].Value = PlaceOfLiving_textBox.Text;
                this.AcceptAndWriteChanges(); // применить изменения
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В PhoneRegistration_textBox НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void PhoneRegistration_textBox_Leave(object sender, EventArgs e)
        {
            if (PhoneRegistration_textBox.Text != dataGridView1[IndexPhoneRegistration, IndexRowLichnayaKarta].Value.ToString())
            {
                dataGridView1[IndexPhoneRegistration, IndexRowLichnayaKarta].Value = PhoneRegistration_textBox.Text;
                this.AcceptAndWriteChanges(); // применить изменения
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В PhonePlaceOfLiving_textBox НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void PhonePlaceOfLiving_textBox_Leave(object sender, EventArgs e)
        {
            if (PhonePlaceOfLiving_textBox.Text != dataGridView1[IndexPhonePlaceOfLiving, IndexRowLichnayaKarta].Value.ToString())
            {
                dataGridView1[IndexPhonePlaceOfLiving, IndexRowLichnayaKarta].Value = PhonePlaceOfLiving_textBox.Text;
                this.AcceptAndWriteChanges(); // применить изменения
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В Post_textBox НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void Post_textBox_Leave(object sender, EventArgs e)
        {
            if (Post_textBox.Text != dataGridView1[IndexPost, IndexRowLichnayaKarta].Value.ToString())
            {
                dataGridView1[IndexPost, IndexRowLichnayaKarta].Value = Post_textBox.Text;
                this.AcceptAndWriteChanges(); // применить изменения
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В DateOfBirth__dateTimePicker НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void DateOfBirth_dateTimePicker_ValueChanged(object sender, EventArgs e) // dateTimePicker "Дата рождения"
        {
            if (Card1to9WasLoaded == 1)
            {
                dataGridView1[IndexDateOfBirth, IndexRowLichnayaKarta].Value = DateOfBirth_dateTimePicker.Value.ToString("dd.MM.yyyy");
                this.AcceptAndWriteChanges(); // применить изменения
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В RankDate_dateTimePicker НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void RankDate_dateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            if (Card1to9WasLoaded == 1)
            {
                dataGridView1[IndexRankDate, IndexRowLichnayaKarta].Value = RankDate_dateTimePicker.Value.ToString("dd.MM.yyyy");
                this.PereschetZvanie();
                this.AcceptAndWriteChanges(); // применить изменения
                NextRankDate_textBox.Text = dataGridView1[IndexNextRankDate, IndexRowLichnayaKarta].Value.ToString();
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В Rank_comboBox НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void Rank_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Card1to9WasLoaded == 1)
            {
                dataGridView1[IndexRank, IndexRowLichnayaKarta].Value = Rank_comboBox.Text;
                this.PereschetZvanie();
                this.AcceptAndWriteChanges(); // применить изменения
                NextRankDate_textBox.Text = dataGridView1[IndexNextRankDate, IndexRowLichnayaKarta].Value.ToString();
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В RankLimit_comboBox НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void RankLimit_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Card1to9WasLoaded == 1)
            {
                dataGridView1[IndexRankLimit, IndexRowLichnayaKarta].Value = RankLimit_comboBox.Text;
                this.PereschetZvanie();
                this.AcceptAndWriteChanges(); // применить изменения
                NextRankDate_textBox.Text = dataGridView1[IndexNextRankDate, IndexRowLichnayaKarta].Value.ToString();
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В Gender_comboBox НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void Gender_comboBox_SelectedIndexChanged(object sender, EventArgs e) // ComboBox "Пол"
        {
            if (Card1to9WasLoaded == 1)
            {
                dataGridView1[IndexGender, IndexRowLichnayaKarta].Value = Gender_comboBox.Text;
                this.AcceptAndWriteChanges(); // применить изменения
            }
        }


        // ######################################################
        // ##  КНОПКА "ВЫБРАТЬ ФОТО" НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##
        // ######################################################
        private void ChooseImage_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Title = "Выберите новую фотографию сотрудника";
            openFileDialog2.InitialDirectory = "c:\\";
            openFileDialog2.Filter = "Все изображения|*.bmp; *.jpg; *.jpeg; *.png; *.gif";
            if (openFileDialog2.ShowDialog() == DialogResult.OK) // если пользователь выбрал файл изображения
            {
                Bitmap bmp = new Bitmap(openFileDialog2.FileName); // присваиваем переменной bmp выбранный файл
                TypeConverter converter = TypeDescriptor.GetConverter(typeof(Bitmap));
                string ImageBase64 = Convert.ToBase64String((byte[])converter.ConvertTo(bmp, typeof(byte[]))); // конвертируем изображение в текст
                dataGridView1[IndexImageString, IndexRowLichnayaKarta].Value = ImageBase64; // записываем результат в соответствующую ячейку
                pictureBox1.Image = bmp; //присваиваем pictureBox1 собранную ячейку
                this.AcceptAndWriteChanges(); // применить изменения
            }
        }


        // ######################################################
        // ##  КНОПКА "УДАЛИТЬ ФОТО" НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##
        // ######################################################
        private void RemoveImage_Click(object sender, EventArgs e)
        {
            dataGridView1[IndexImageString, IndexRowLichnayaKarta].Value = XMLDB.DefaultImageBase64;
            Bitmap bmp = new Bitmap(new MemoryStream(Convert.FromBase64String(dataGridView1[IndexImageString, IndexRowLichnayaKarta].Value.ToString()))); // собираем изображение
            pictureBox1.Image = bmp;
            this.AcceptAndWriteChanges(); // применить изменения
        }



        //               //""""""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 10-11"  ############################################################
        public void UpdateCard10and11()
        {
            dataGridView_Study.Rows.Clear();
            dataGridView_Study.AutoGenerateColumns = false;

            this.Draw_dataGridView_All(IndexStudy, dataGridView_Study); // Отрисовываем таблицу dataGridView_Study

            Study_FormaObucheniya.MinimumWidth = 120;
            Study_Naimenovanie.MinimumWidth = 130;
            Study_DataPost.MinimumWidth = 120;
            Study_DataPost.Width = 120;
            Study_DataOkonch.MinimumWidth = 120;
            Study_DataOkonch.Width = 120;
            Study_Document.MinimumWidth = 140;

            //Чтобы не обрезался текст, расчитываем ширину выпадающего списка, когда ComboBox в режиме редактирования
            Study_FormaObucheniya.DropDownWidth = Study_FormaObucheniya.Items.Cast<Object>().Select(x => x.ToString())
    .Max(x => TextRenderer.MeasureText(x, Study_FormaObucheniya.InheritedStyle.Font, Size.Empty, TextFormatFlags.Default).Width);

            /*==============================================================================================================*/

            dataGridView_UchStepen.Rows.Clear();
            dataGridView_UchStepen.AutoGenerateColumns = false;

            this.Draw_dataGridView_All(IndexUchStepen, dataGridView_UchStepen); // Отрисовываем таблицу dataGridView_UchStepen
        }        


        // ###################################################################
        // ##  КНОПКА "ДОБАВИТЬ УЧЕНУЮ СТЕПЕНЬ" НА ВКЛАДКЕ "КАРТОЧКА 10-11" ##
        // ###################################################################
        private void UchStepenAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                this.SaveChangesToDataGridView_All(IndexUchStepen, dataGridView_UchStepen);
            }
            else //Если метод вызван нажатием кнопки
            {
                dataGridView_UchStepen.Rows.Add("---", DateTime.Now.ToString("dd.MM.yyyy")); // добавить ученую степень
                this.SaveChangesToDataGridView_All(IndexUchStepen, dataGridView_UchStepen);
            }
        }


        // ################################################################
        // ##  КНОПКА "ДОБАВИТЬ ОБРАЗОВАНИЕ" НА ВКЛАДКЕ "КАРТОЧКА 10-11" ##
        // ################################################################
        private void StudyAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                this.SaveChangesToDataGridView_All(IndexStudy, dataGridView_Study);
            }
            else //Если метод вызван нажатием кнопки
            {
                dataGridView_Study.Rows.Add("Высшее (очное)", "---", DateTime.Now.ToString("dd.MM.yyyy"), DateTime.Now.ToString("dd.MM.yyyy"), "---", "---", "---"); // добавить образование
                this.SaveChangesToDataGridView_All(IndexStudy, dataGridView_Study);
            }
        }



        //               //"""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 12"  ############################################################
        public void UpdateCard12()
        {
            dataGridView_PrisvZvaniy.Rows.Clear();
            dataGridView_PrisvZvaniy.AutoGenerateColumns = false;

            this.Draw_dataGridView_All(IndexPrisvZvaniy, dataGridView_PrisvZvaniy); // Отрисовываем таблицу dataGridView_PrisvZvaniy
        }
        

        // ######################################################################
        // ##  КНОПКА "ДОБАВИТЬ ЗВАНИЕ, КЛАССНЫЙ ЧИН" НА ВКЛАДКЕ "КАРТОЧКА 12" ##
        // ######################################################################
        private void ZvanieAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                this.SaveChangesToDataGridView_All(IndexPrisvZvaniy, dataGridView_PrisvZvaniy);
            }
            else //Если метод вызван нажатием кнопки
            {
                dataGridView_PrisvZvaniy.Rows.Add("---", DateTime.Now.ToString("dd.MM.yyyy"), "---", "---", DateTime.Now.ToString("dd.MM.yyyy")); // добавить звание, классный чин
                this.SaveChangesToDataGridView_All(IndexPrisvZvaniy, dataGridView_PrisvZvaniy);
            }
        }



        //               //""""""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 13-14"  ############################################################
        public void UpdateCard13and14()
        {
            dataGridView_Married.Rows.Clear();
            dataGridView_Married.AutoGenerateColumns = false;

            this.Draw_dataGridView_All(IndexMarried, dataGridView_Married); // Отрисовываем таблицу dataGridView_Married

            /*==============================================================================================================*/

            dataGridView_Family.Rows.Clear();
            dataGridView_Family.AutoGenerateColumns = false;

            this.Draw_dataGridView_All(IndexFamily, dataGridView_Family); // Отрисовываем таблицу dataGridView_Family

            Family_StepenRodstva.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // Выравнивание по центру в колонке "Степень родства"
            Family_DateOfBirth.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // Выравнивание по центру в колонке "Дата рождения"
            Family_DateOfBirth.MinimumWidth = 120;
        }

        
        // ############################################################
        // ##  КНОПКА "ДОБАВИТЬ СОБЫТИЕ" НА ВКЛАДКЕ "КАРТОЧКА 13-14" ##
        // ############################################################
        private void MarriedAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                this.SaveChangesToDataGridView_All(IndexMarried, dataGridView_Married);
            }
            else //Если метод вызван нажатием кнопки
            {
                dataGridView_Married.Rows.Add("Женат", DateTime.Now.ToString("yyyy")); // добавить событие (свадьба, развод)
                this.SaveChangesToDataGridView_All(IndexMarried, dataGridView_Married);
            }
        }


        // ################################################################
        // ##  КНОПКА "ДОБАВИТЬ ЧЛЕНА СЕМЬИ" НА ВКЛАДКЕ "КАРТОЧКА 13-14" ##
        // ################################################################
        private void FamilyAddPerson_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                this.SaveChangesToDataGridView_All(IndexFamily, dataGridView_Family);
            }
            else //Если метод вызван нажатием кнопки
            {
                dataGridView_Family.Rows.Add("Мать", DateTime.Now.ToString("dd.MM.yyyy"), "---"); // добавить члена семьи
                this.SaveChangesToDataGridView_All(IndexFamily, dataGridView_Family);
            }
        }



        //               //"""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 15"  ############################################################
        public void UpdateCard15()
        {
            dataGridView_TrudDeyat.Rows.Clear();
            dataGridView_TrudDeyat.AutoGenerateColumns = false;

            this.Draw_dataGridView_All(IndexTrudDeyat, dataGridView_TrudDeyat); // Отрисовываем таблицу dataGridView_TrudDeyat

            TrudDeyat_DataNaznach.MinimumWidth = 130;
            TrudDeyat_DataNaznach.Width = 130;
            TrudDeyat_DataOsvobozhd.MinimumWidth = 130;
            TrudDeyat_DataOsvobozhd.Width = 130;
            TrudDeyat_LgotKoeff.Width = 130;
            TrudDeyat_LgotKoeff.MinimumWidth = 130;
            TrudDeyat_Sokrash.MinimumWidth = 120;
            TrudDeyat_Sokrash.Width = 120;
        }



        // ##############################################################
        // ##  КНОПКА "ДОБАВИТЬ МЕСТО РАБОТЫ" НА ВКЛАДКЕ "КАРТОЧКА 15" ##
        // ##############################################################
        private void TrudDeyatAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                this.SaveChangesToDataGridView_All(IndexTrudDeyat, dataGridView_TrudDeyat);
            }
            else //Если метод вызван нажатием кнопки
            {
                dataGridView_TrudDeyat.Rows.Add(DateTime.Now.ToString("dd.MM.yyyy"), DateTime.Now.ToString("dd.MM.yyyy"), "1", "---", "У"); // добавить место работы
                this.SaveChangesToDataGridView_All(IndexTrudDeyat, dataGridView_TrudDeyat);
            }
        }



        private void FormResizeBegin(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab.Text == "Карточка 16-18")
            {
                StillResizing = 1;
            }
        }

        private void FormResize(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab.Text == "Карточка 16-18")
            {
                if ((StillResizing == 0) & (WindowState != FormWindowState.Minimized)) StazhVysluga_Resize(); // дополнительно отключаем подстройку высоты строк при сворачивании окна
            }
        }

        private void FormResizeEnd(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab.Text == "Карточка 16-18")
            {
                StazhVysluga_Resize(); // Перерисовываем таблицу под новый размер
                StillResizing = 0;
            }
        }



        //               //""""""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 16-18"  ############################################################
        public void UpdateCard16to18()
        {
            Card16to18WasLoaded = 0;

            dataGridView_StazhVysluga.Rows.Clear();
            dataGridView_StazhVysluga.AutoGenerateColumns = false;

            this.Draw_dataGridView_All(IndexStazhVysluga, dataGridView_StazhVysluga); // Отрисовываем таблицу dataGridView_StazhVysluga

            dataGridView_StazhVysluga.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            StazhVysluga_Resize();

            /*==============================================================================================================*/

            DataPrisyagi_dateTimePicker.Text = dataGridView1[IndexDataPrisyagi, IndexRowLichnayaKarta].Value.ToString();

            /*==============================================================================================================*/

            dataGridView_RabotaGFS.Rows.Clear();
            dataGridView_RabotaGFS.AutoGenerateColumns = false;

            this.Draw_dataGridView_All(IndexRabotaGFS, dataGridView_RabotaGFS); // Отрисовываем таблицу dataGridView_RabotaGFS

            RabotaGFS_DataNaznach.MinimumWidth = 120;
            RabotaGFS_DataNaznach.Width = 120;
            RabotaGFS_DataOsvobozhd.MinimumWidth = 120;
            RabotaGFS_DataOsvobozhd.Width = 120;
            RabotaGFS_Dolzhnost.MinimumWidth = 200;
            RabotaGFS_CheyPrikaz.Width = 80;
            RabotaGFS_NomerPrikaza.MinimumWidth = 80;
            RabotaGFS_NomerPrikaza.Width = 80;
            RabotaGFS_DataPrikaza.Width = 120;
            RabotaGFS_Stavka.Width = 80;
            RabotaGFS_LgotKoeff.Width = 80;

            Card16to18WasLoaded = 1; // карточка прогрузилась
        }


        // ##########  ПОДСТРОЙКА РАЗМЕРА СТРОК В DataGridView_StazhVysluga НА ВКЛАДКЕ "КАРТОЧКА 16-18" ##########
        public void StazhVysluga_Resize()
        {
            dataGridView_StazhVysluga.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells; // Включаем свойство AutoSizeRowsMode, чтобы оно автоматически подстроило высоту строк в таблице
            StazhVysluga_Poyasnenie.DefaultCellStyle.WrapMode = DataGridViewTriState.True; // Перенос слов в колонке "Пояснение" 
            StazhVysluga_Poyasnenie.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            StazhVysluga_Let.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // Выравнивание в колонке "Лет"
            StazhVysluga_Mesyacev.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // Выравнивание в колонке "Месяцев"
            StazhVysluga_Dney.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // Выравнивание в колонке "Дней"

            int Stroka4 = dataGridView_StazhVysluga[0, 3].OwningRow.Height; // Записываем в переменную высоту строки, присвоенную AutoSizeRowsMode
            int Stroka_proverka = dataGridView_StazhVysluga[0, 2].OwningRow.Height; // Записываем в переменную высоту "стандартной" строки для дальнейшего сравнения
            int Zagolovok = dataGridView_StazhVysluga.ColumnHeadersHeight; // Записываем в переменную высоту заголовка, присвоенную AutoSizeRowsMode
            dataGridView_StazhVysluga.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None; // Отключаем свойство AutoSizeRowsMode, чтобы далее можно было программно присвоить высоту строк в таблице
            dataGridView_StazhVysluga.ColumnHeadersHeight = Zagolovok;

            if (Stroka4 == Stroka_proverka) // Если размер окна программы поволяет уместить текст в четвертой строке без переносов, значит ставим всем одну высоту  
                foreach (DataGridViewRow row in dataGridView_StazhVysluga.Rows)
                {
                    row.Height = (dataGridView_StazhVysluga.Height - Zagolovok) / (dataGridView_StazhVysluga.Rows.Count);// Вычисляем высоту строк для заполнения всего свободного пространства
                }

            else // Если размер окна программы НЕ поволяет уместить текст в четвертой строке без переносов, значит высота этой строки должна быть больше, чем у других
                foreach (DataGridViewRow row in dataGridView_StazhVysluga.Rows)
                {
                    if (row != dataGridView_StazhVysluga[0, 3].OwningRow) // Для всех строк, кроме четвертой
                    {
                        row.Height = (dataGridView_StazhVysluga.Height - Stroka4 - Zagolovok) / (dataGridView_StazhVysluga.Rows.Count - 1); // Вычисляем высоту строк для заполнения всего свободного пространства
                    }
                    else row.Height = Stroka4; // Восстанавливаем высоту строки, присвоенную в самом начале свойством AutoSizeRowsMode
                }
        }


        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В DataGridView_StazhVysluga НА ВКЛАДКЕ "КАРТОЧКА 16-18" ##########
        private void SaveChangesToDataGridView_StazhVysluga(object sender, EventArgs e)
        {
            MessageBox.Show("Stazh");
            this.SaveChangesToDataGridView_All(IndexStazhVysluga, dataGridView_StazhVysluga);
        }


        // #################################################################
        // ##  КНОПКА "ДОБАВИТЬ МЕСТО СЛУЖБЫ" НА ВКЛАДКЕ "КАРТОЧКА 16-18" ##
        // #################################################################
        private void RabotaGFSAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                this.SaveChangesToDataGridView_All(IndexRabotaGFS, dataGridView_RabotaGFS);
            }
            else //Если метод вызван нажатием кнопки
            {
                dataGridView_RabotaGFS.Rows.Add(DateTime.Now.ToString("dd.MM.yyyy"), DateTime.Now.ToString("dd.MM.yyyy"), "---", "---", "---", DateTime.Now.ToString("dd.MM.yyyy"), "1", "0"); // добавить место службы
                this.SaveChangesToDataGridView_All(IndexRabotaGFS, dataGridView_RabotaGFS);
            }
        }



        //               //""""""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 19-20"  ############################################################
        public void UpdateCard19and20()
        {
            dataGridView_Attestaciya.Rows.Clear();
            dataGridView_Attestaciya.AutoGenerateColumns = false;

            this.Draw_dataGridView_All(IndexAttestaciya, dataGridView_Attestaciya); // Отрисовываем таблицу dataGridView_Attestaciya

            Attestaciya_Data.Width = 140;
            Attestaciya_Data.MinimumWidth = 140;
            Attestaciya_Prichina.Width = 180;
            Attestaciya_Prichina.MinimumWidth = 180;

            /*==============================================================================================================*/

            dataGridView_ProfPodg.Rows.Clear();
            dataGridView_ProfPodg.AutoGenerateColumns = false;

            this.Draw_dataGridView_All(IndexProfPodg, dataGridView_ProfPodg); // Отрисовываем таблицу dataGridView_ProfPodg

            ProfPodg_VidObuch.Width = 270;
            ProfPodg_VidObuch.MinimumWidth = 270;
            ProfPodg_DataNach.Width = 120;
            ProfPodg_DataNach.MinimumWidth = 120;
            ProfPodg_DataOkonch.Width = 120;
            ProfPodg_DataOkonch.MinimumWidth = 120;
        }


        // ###############################################################
        // ##  КНОПКА "ДОБАВИТЬ АТТЕСТАЦИЮ" НА ВКЛАДКЕ "КАРТОЧКА 19-20" ##
        // ###############################################################
        private void AttestaciyaAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                this.SaveChangesToDataGridView_All(IndexAttestaciya, dataGridView_Attestaciya);
            }
            else //Если метод вызван нажатием кнопки
            {
                dataGridView_Attestaciya.Rows.Add(DateTime.Now.ToString("dd.MM.yyyy"), "Плановая", "Cоответствует замещаемой должности"); // добавить аттестацию
                this.SaveChangesToDataGridView_All(IndexAttestaciya, dataGridView_Attestaciya);
            }
            Calculate_NextAttestaciyaDate(); // высчитываем дату следующей аттестации
        }


        // ###############################################################
        // ##  КНОПКА "ДОБАВИТЬ ПОДГОТОВКУ" НА ВКЛАДКЕ "КАРТОЧКА 19-20" ##
        // ###############################################################
        private void ProfPodgAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                this.SaveChangesToDataGridView_All(IndexProfPodg, dataGridView_ProfPodg);
            }
            else //Если метод вызван нажатием кнопки
            {
                dataGridView_ProfPodg.Rows.Add("Первоначальное обучение", DateTime.Now.ToString("dd.MM.yyyy"), DateTime.Now.ToString("dd.MM.yyyy"), "---", "---"); // добавить проф. подготовку
                this.SaveChangesToDataGridView_All(IndexProfPodg, dataGridView_ProfPodg);
            }
        }



        //               //""""""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 21-22"  ############################################################
        public void UpdateCard21and22()
        {
            Card21and22WasLoaded = 0;
            Klassnost_comboBox.Text = dataGridView1[IndexKlassnost, IndexRowLichnayaKarta].Value.ToString();
            KlassnostCheyPrikaz_textBox.Text = dataGridView1[IndexKlassnostCheyPrikaz, IndexRowLichnayaKarta].Value.ToString();
            KlassnostNomerPrikaza_textBox.Text = dataGridView1[IndexKlassnostNomerPrikaza, IndexRowLichnayaKarta].Value.ToString();
            KlassnostDate_textBox.Text = dataGridView1[IndexKlassnostDate, IndexRowLichnayaKarta].Value.ToString();

            /*==============================================================================================================*/

            dataGridView_KlassnostOld.Rows.Clear();
            dataGridView_KlassnostOld.AutoGenerateColumns = false;

            this.Draw_dataGridView_All(IndexKlassnostOld, dataGridView_KlassnostOld); // Отрисовываем таблицу dataGridView_KlassnostOld

            /*==============================================================================================================*/

            dataGridView_Nagrady.Rows.Clear();
            dataGridView_Nagrady.AutoGenerateColumns = false;

            this.Draw_dataGridView_All(IndexNagrady, dataGridView_Nagrady); // Отрисовываем таблицу dataGridView_Nagrady

            Card21and22WasLoaded = 1; // карточка прогрузилась
        }


        // ##########################################################################
        // ##  КНОПКА "ДОБАВИТЬ ПРЕДЫДУЩУЮ КЛАССНОСТЬ" НА ВКЛАДКЕ "КАРТОЧКА 21-22" ##
        // ##########################################################################
        private void KlassnostAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                this.SaveChangesToDataGridView_All(IndexKlassnostOld, dataGridView_KlassnostOld);
            }
            else //Если метод вызван нажатием кнопки
            {
                dataGridView_KlassnostOld.Rows.Add("Специалист 3 класса", "---", "---", DateTime.Now.ToString("dd.MM.yyyy")); // добавить предыдущую классность
                this.SaveChangesToDataGridView_All(IndexKlassnostOld, dataGridView_KlassnostOld);
            }
        }


        // ########################################################################
        // ##  КНОПКА "ДОБАВИТЬ НАГРАДЫ / ПООЩРЕНИЯ" НА ВКЛАДКЕ "КАРТОЧКА 21-22" ##
        // ########################################################################
        private void NagradyAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                this.SaveChangesToDataGridView_All(IndexNagrady, dataGridView_Nagrady);
            }
            else //Если метод вызван нажатием кнопки
            {
                dataGridView_Nagrady.Rows.Add("---", "---", "---", DateTime.Now.ToString("dd.MM.yyyy")); // добавить награды / поощрения
                this.SaveChangesToDataGridView_All(IndexNagrady, dataGridView_Nagrady);
            }
        }


        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В Klassnost_comboBox НА ВКЛАДКЕ "КАРТОЧКА 21-22" ##########
        private void Klassnost_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Card21and22WasLoaded == 1)
            {
                dataGridView1[IndexKlassnost, IndexRowLichnayaKarta].Value = Klassnost_comboBox.Text; // заполняем combobox значением текущей классности
                switch (Klassnost_comboBox.Text) // проверяем, какую классность выбрал пользователь
                {
                    case "Отсутствует":
                        dataGridView1[IndexKlassnostDate, IndexRowLichnayaKarta].Value = "--.--.----"; // "Обнуляем" дату текущей классности
                        dataGridView1[IndexNextKlassnostDate, IndexRowLichnayaKarta].Value = "--.--.----"; // "Обнуляем" дату следующей классности
                        dataGridView1[IndexKlassnostCheyPrikaz, IndexRowLichnayaKarta].Value = "---"; // "Обнуляем" чей приказ о присвоении классности
                        dataGridView1[IndexKlassnostNomerPrikaza, IndexRowLichnayaKarta].Value = "---"; // "Обнуляем" номер приказа о присвоении классности
                        KlassnostCheyPrikaz_textBox.Text = dataGridView1[IndexKlassnostCheyPrikaz, IndexRowLichnayaKarta].Value.ToString(); // Обновляем textbox "Чей приказ" 
                        KlassnostNomerPrikaza_textBox.Text = dataGridView1[IndexKlassnostNomerPrikaza, IndexRowLichnayaKarta].Value.ToString(); // Обновляем textbox "Номер приказа" 
                        KlassnostCheyPrikaz_textBox.ReadOnly = true; // Если классность отсутствует, окно для ввода должно быть неактивным 
                        KlassnostNomerPrikaza_textBox.ReadOnly = true; // Если классность отсутствует, окно для ввода должно быть неактивным
                        break;

                    case "Специалист 3 класса":
                    case "Специалист 2 класса":
                    case "Специалист 1 класса":
                        KlassnostCheyPrikaz_textBox.ReadOnly = false;
                        KlassnostNomerPrikaza_textBox.ReadOnly = false;
                        dataGridView1[IndexKlassnostDate, IndexRowLichnayaKarta].Value = DateTime.Now.ToString("dd.MM.yyyy"); // выводим дату присвоения классности 
                        dataGridView1[IndexNextKlassnostDate, IndexRowLichnayaKarta].Value = DateTime.Now.AddYears(3).ToString("dd.MM.yyyy"); // дата присвоения, плюс 3 года
                        break;

                    case "Мастер":
                        KlassnostCheyPrikaz_textBox.ReadOnly = false;
                        KlassnostNomerPrikaza_textBox.ReadOnly = false;
                        dataGridView1[IndexKlassnostDate, IndexRowLichnayaKarta].Value = DateTime.Now.ToString("dd.MM.yyyy"); // выводим дату присвоения классности 
                        dataGridView1[IndexNextKlassnostDate, IndexRowLichnayaKarta].Value = "высшее звание"; // высшая классность
                        break;
                }
                this.AcceptAndWriteChanges(); // применить изменения
                KlassnostDate_textBox.Text = dataGridView1[IndexKlassnostDate, IndexRowLichnayaKarta].Value.ToString(); //обновить окошко даты присвоения классности
            }
        }



        //               //""""""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 23-25"  ############################################################
        public void UpdateCard23to25()
        {
            Card23to25WasLoaded = 0;
            dataGridView_Prodlenie.Rows.Clear();
            dataGridView_Prodlenie.AutoGenerateColumns = false;

            this.Draw_dataGridView_All(IndexProdlenie, dataGridView_Prodlenie); // Отрисовываем таблицу dataGridView_Prodlenie

            if (dataGridView_Prodlenie.Rows.Count != 0) //проверка на существование данных в таблице
            {
                Prodlenie_checkBox.CheckState = CheckState.Checked;
            }
            else
            {
                Prodlenie_checkBox.CheckState = CheckState.Unchecked;
            }

            /*==============================================================================================================*/

            dataGridView_Boevye.Rows.Clear();
            dataGridView_Boevye.AutoGenerateColumns = false;

            this.Draw_dataGridView_All(IndexBoevye, dataGridView_Boevye); // Отрисовываем таблицу dataGridView_Boevye

            /*==============================================================================================================*/

            dataGridView_Rezerv.Rows.Clear();
            dataGridView_Rezerv.AutoGenerateColumns = false;

            this.Draw_dataGridView_All(IndexRezerv, dataGridView_Rezerv); // Отрисовываем таблицу dataGridView_Rezerv

            Card23to25WasLoaded = 1; // карточка прогрузилась
        }


        // ###############################################################################
        // ##  КНОПКА "ДОБАВИТЬ УЧАСТИЕ В БОЕВЫХ ДЕЙСТВИЯХ" НА ВКЛАДКЕ "КАРТОЧКА 23-25" ##
        // ###############################################################################
        private void BoevyeAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                this.SaveChangesToDataGridView_All(IndexBoevye, dataGridView_Boevye);
            }
            else //Если метод вызван нажатием кнопки
            {
                dataGridView_Boevye.Rows.Add("---", DateTime.Now.ToString("dd.MM.yyyy"), DateTime.Now.ToString("dd.MM.yyyy"), "1", "---"); // добавить участие в боевых действиях
                this.SaveChangesToDataGridView_All(IndexBoevye, dataGridView_Boevye);
            }
        }

        // ########################################################################
        // ##  КНОПКА "ДОБАВИТЬ СОСТОЯНИЕ В РЕЗЕРВЕ" НА ВКЛАДКЕ "КАРТОЧКА 23-25" ##
        // ########################################################################
        private void RezervAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                this.SaveChangesToDataGridView_All(IndexRezerv, dataGridView_Rezerv);
            }
            else //Если метод вызван нажатием кнопки
            {
                dataGridView_Rezerv.Rows.Add("---", DateTime.Now.ToString("yyyy"), "---", DateTime.Now.ToString("dd.MM.yyyy")); // добавить состояние в резерве
                this.SaveChangesToDataGridView_All(IndexRezerv, dataGridView_Rezerv);
            }
        }


        // ##########  ИЗМЕНЕНИЕ СОСТОЯНИЯ Prodlenie_checkBox НА ВКЛАДКЕ "КАРТОЧКА 23-25" ##########
        private void Prodlenie_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            if (Card23to25WasLoaded == 1)
            {
                if (Prodlenie_checkBox.CheckState == CheckState.Checked)
                {
                    if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
                    {
                        this.SaveChangesToDataGridView_All(IndexProdlenie, dataGridView_Prodlenie);
                    }
                    else //Если метод вызван нажатием кнопки
                    {
                        dataGridView_Prodlenie.Rows.Add(DateTime.Now.ToString("dd.MM.yyyy"), "1"); // добавить продление службы
                        this.SaveChangesToDataGridView_All(IndexProdlenie, dataGridView_Prodlenie);
                    }
                }
                else if (Prodlenie_checkBox.CheckState == CheckState.Unchecked)
                {
                    dataGridView_Prodlenie.Rows.Clear();
                    dataGridView1[IndexProdlenie, IndexRowLichnayaKarta].Value = "";
                    this.AcceptAndWriteChanges(); // Применить изменения
                }
                else if (Prodlenie_checkBox.CheckState == CheckState.Indeterminate)
                {
                    MessageBox.Show("Неизвестная ошибка. Проверьте еще раз отметку о продлении выслуги.");
                }
            }
        }



        //               //""""""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 26-29"  ############################################################
        public void UpdateCard26to29()
        {
            Card26to29WasLoaded = 0;
            if (CurrentDataTableName == "Kadry")
            {
                SendToAnotherDataTable.Text = "Переместить в архив";
            }
            else SendToAnotherDataTable.Text = "Восстановить из архива";

            dataGridView_Vzyskaniya.Rows.Clear();
            dataGridView_Vzyskaniya.AutoGenerateColumns = false;

            this.Draw_dataGridView_All(IndexVzyskaniya, dataGridView_Vzyskaniya); // Отрисовываем таблицу dataGridView_Vzyskaniya

            /*==============================================================================================================*/

            dataGridView_Uvolnenie.Rows.Clear();
            dataGridView_Uvolnenie.AutoGenerateColumns = false;

            this.Draw_dataGridView_All(IndexUvolnenie, dataGridView_Uvolnenie); // Отрисовываем таблицу dataGridView_Uvolnenie

            /*==============================================================================================================*/

            Zapolnil_textBox.Text = dataGridView1[IndexZapolnil, IndexRowLichnayaKarta].Value.ToString();

            /*==============================================================================================================*/

            DataZapolneniya_dateTimePicker.Text = dataGridView1[IndexDataZapolneniya, IndexRowLichnayaKarta].Value.ToString();

            Card26to29WasLoaded = 1; // карточка прогрузилась
        }



        // ##############################################################
        // ##  КНОПКА "ДОБАВИТЬ ВЗЫСКАНИЕ" НА ВКЛАДКЕ "КАРТОЧКА 26-29" ##
        // ##############################################################
        private void VzyskaniyaAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                this.SaveChangesToDataGridView_All(IndexVzyskaniya, dataGridView_Vzyskaniya);
            }
            else //Если метод вызван нажатием кнопки
            {
                dataGridView_Vzyskaniya.Rows.Add("---", "---", "---", "---", DateTime.Now.ToString("dd.MM.yyyy"), "---", "---", DateTime.Now.ToString("dd.MM.yyyy")); // добавить взыскание
                this.SaveChangesToDataGridView_All(IndexVzyskaniya, dataGridView_Vzyskaniya);
            }
        }

        // ###############################################################
        // ##  КНОПКА "ДОБАВИТЬ УВОЛЬНЕНИЕ" НА ВКЛАДКЕ "КАРТОЧКА 26-29" ##
        // ###############################################################
        private void UvolnenieAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                this.SaveChangesToDataGridView_All(IndexUvolnenie, dataGridView_Uvolnenie);
            }
            else //Если метод вызван нажатием кнопки
            {
                string UvolnenieProverka = dataGridView1[IndexUvolnenie, IndexRowLichnayaKarta].Value.ToString();
                if (UvolnenieProverka == "") //Если информация об увольнении отсутствует
                {
                    dataGridView_Uvolnenie.Rows.Add(DateTime.Now.ToString("dd.MM.yyyy"), "---", "---", DateTime.Now.ToString("dd.MM.yyyy"), "---"); // добавить увольнение
                    this.SaveChangesToDataGridView_All(IndexUvolnenie, dataGridView_Uvolnenie);                
                }
                else
                {
                    MessageBox.Show("Информация об увольнении уже существует!");
                }
            }
        }


        // ###########################################################################
        // ##########  ОБЩИЙ МЕТОД ОТРИСОВКИ dataGridView НА ВСЕХ ВКЛАДКАХ  ##########
        // ###########################################################################
        private void Draw_dataGridView_All(int Index, DataGridView dataGridView_Name)
        {
            string StringDataGrid = dataGridView1[Index, IndexRowLichnayaKarta].Value.ToString();
            if (StringDataGrid != "") //проверка на существование данных в таблице
            {
                string[] string_array = StringDataGrid.Split('$');

                foreach (string s in string_array)
                {
                    string[] Row = s.Split('^');
                    dataGridView_Name.Rows.Add(Row);
                }
            }
        }

        // #######################################################################################
        // ##########  ОБЩИЙ МЕТОД СОХРАНЕНИЯ ИЗМЕНЕНИЙ В DataGridView НА ВСЕХ ВКЛАДКАХ ##########
        // #######################################################################################
        private void SaveChangesToDataGridView_All(int Index, DataGridView dataGridView_Name)
        {
            StringBuilder sb = new StringBuilder(); // создаем строку для построения
            foreach (DataGridViewRow row in dataGridView_Name.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    // Если обрабатываемая ячейка из колонки типа CalendarColumn, то обрабатываем неверный формат даты.
                    if (cell.OwningColumn is CalendarColumn)
                    {
                        DateTime wrongdatetoconvert = DateTime.Parse(cell.Value.ToString()); // парсим её в формат DateTime
                        cell.Value = wrongdatetoconvert.ToString("dd.MM.yyyy"); // и конвертируем в нужный формат
                    }
                    sb.Append(cell.Value); // добавляем значение ячейки
                    sb.Append("^"); // ставим разделитель ячеек 
                }
                sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель ячеек
                sb.Append("$"); // ставим разделитель строки
            }
            sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель строки

            dataGridView1[Index, IndexRowLichnayaKarta].Value = sb.ToString(); // Заполняем ячейку результирующей строкой
            this.AcceptAndWriteChanges(); // Применить изменения
        }


        // ##########  ВЫСЧИТЫВАНИЕ ДАТЫ СЛЕДУЮЩЕЙ АТТЕСТАЦИИ ##########
        private void Calculate_NextAttestaciyaDate()
        {
            foreach (DataGridViewRow row in dataGridView_Attestaciya.Rows)
            {
                if (row.Index + 1 == dataGridView_Attestaciya.Rows.Count) //Находим последнюю строку в dataGridView_Attestaciya
                {
                    foreach (DataGridViewCell cell in row.Cells) //Пробегаем по ячейкам в найденной строке
                    {
                        // Обработка неверного формата даты и высчитывание даты следующей аттестации
                        if (cell.OwningColumn is CalendarColumn)
                        {
                            DateTime wrongdatetoconvert = DateTime.Parse(cell.Value.ToString()); // парсим её в формат DateTime
                            // Заполняем ячейку "Дата следующей аттестации", прибавив 4 года к последней аттестации
                            dataGridView1[IndexNextAttestaciyaDate, IndexRowLichnayaKarta].Value = wrongdatetoconvert.AddYears(4).ToString("dd.MM.yyyy");
                        }
                    }
                }             
            }
            this.AcceptAndWriteChanges(); // Применить изменения
        }


        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В DataPrisyagi_dateTimePicker НА ВКЛАДКЕ "КАРТОЧКА 16-18" ##########
        private void DataPrisyagi_dateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            if (Card16to18WasLoaded == 1)
            {
                dataGridView1[IndexDataPrisyagi, IndexRowLichnayaKarta].Value = DataPrisyagi_dateTimePicker.Value.ToString("dd.MM.yyyy");
                this.AcceptAndWriteChanges(); // применить изменения
            }
        }


        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В KlassnostCheyPrikaz_textBox НА ВКЛАДКЕ "КАРТОЧКА 21-22" ##########
        private void KlassnostCheyPrikaz_textBox_TextChanged(object sender, EventArgs e)
        {
            if (Card21and22WasLoaded == 1)
            {
                dataGridView1[IndexKlassnostCheyPrikaz, IndexRowLichnayaKarta].Value = KlassnostCheyPrikaz_textBox.Text;
                this.AcceptAndWriteChanges(); // применить изменения
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В KlassnostNomerPrikaza_textBox НА ВКЛАДКЕ "КАРТОЧКА 21-22" ##########
        private void KlassnostNomerPrikaza_textBox_TextChanged(object sender, EventArgs e)
        {
            if (Card21and22WasLoaded == 1)
            {
                dataGridView1[IndexKlassnostNomerPrikaza, IndexRowLichnayaKarta].Value = KlassnostNomerPrikaza_textBox.Text;
                this.AcceptAndWriteChanges(); // применить изменения
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В Zapolnil_textBox НА ВКЛАДКЕ "КАРТОЧКА 26-29" ##########
        private void Zapolnil_textBox_TextChanged(object sender, EventArgs e)
        {
            if (Card26to29WasLoaded == 1)
            {
                dataGridView1[IndexZapolnil, IndexRowLichnayaKarta].Value = Zapolnil_textBox.Text;
                this.AcceptAndWriteChanges(); // применить изменения
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В DataZapolneniya_dateTimePicker НА ВКЛАДКЕ "КАРТОЧКА 26-29" ##########
        private void DataZapolneniya_dateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            if (Card26to29WasLoaded == 1)
            {
                dataGridView1[IndexDataZapolneniya, IndexRowLichnayaKarta].Value = DataZapolneniya_dateTimePicker.Value.ToString("dd.MM.yyyy");
                this.AcceptAndWriteChanges(); // применить изменения
            }
        }


        // ###################################################
        // ##  КНОПКА "ПРЕДЫДУЩАЯ КАРТОЧКА" (СТРЕЛКА ВЛЕВО) ##
        // ###################################################
        private void PrevCard_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 1)
            {
                MessageBox.Show("Это единственная личная карточка");
            }
            else if (IndexRowLichnayaKarta == 0)
            {
                MessageBox.Show("Это первая личная карточка");
            }
            else
            {
                IndexRowLichnayaKarta = IndexRowLichnayaKarta - 1;
                this.NeedToUpdateCard(); // обновляем все поля личной карточки
            }
            Cnum_label.Text = (IndexRowLichnayaKarta + 1).ToString() + " из " + dataGridView1.RowCount.ToString(); // Порядковый номер личной карточки
            CardsFIO_label.Text = dataGridView1[IndexSurname, IndexRowLichnayaKarta].Value.ToString() + " "
    + dataGridView1[IndexName, IndexRowLichnayaKarta].Value.ToString() + " "
    + dataGridView1[IndexMiddleName, IndexRowLichnayaKarta].Value.ToString(); // Прописываем ФИО над стрелками в карточке
        }



        // ###################################################
        // ##  КНОПКА "СЛЕДУЮЩАЯ КАРТОЧКА" (СТРЕЛКА ВПРАВО) ##
        // ###################################################
        private void NextCard_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 1)
            {
                MessageBox.Show("Это единственная личная карточка");
            }
            else if (IndexRowLichnayaKarta == dataGridView1.RowCount - 1)
            {
                MessageBox.Show("Это последняя личная карточка");
            }
            else
            {
                IndexRowLichnayaKarta = IndexRowLichnayaKarta + 1;
                this.NeedToUpdateCard(); // обновляем все поля личной карточки
            }
            Cnum_label.Text = (IndexRowLichnayaKarta + 1).ToString() + " из " + dataGridView1.RowCount.ToString(); // Порядковый номер личной карточки
            CardsFIO_label.Text = dataGridView1[IndexSurname, IndexRowLichnayaKarta].Value.ToString() + " "
    + dataGridView1[IndexName, IndexRowLichnayaKarta].Value.ToString() + " "
    + dataGridView1[IndexMiddleName, IndexRowLichnayaKarta].Value.ToString(); // Прописываем ФИО над стрелками в карточке
        }



        // ####################################################
        // ##  КНОПКА "ЗАКРЫТЬ" НА ВКЛАДКЕ "ЛИЧНАЯ КАРТОЧКА" ##
        // ####################################################
        private void Close2_Click(object sender, EventArgs e)
        {
            Close();
        }

        // #######################################################################################
        // ##  КНОПКА "ПЕРЕМЕСТИТЬ В АРХИВ/ВОССТАНОВИТЬ ИЗ АРХИВА" НА ВКЛАДКЕ "ЛИЧНАЯ КАРТОЧКА" ##
        // #######################################################################################
        private void SendToAnotherDataTable_Click(object sender, EventArgs e)
        {
            Required_PersonalFileNum = dataGridView1[IndexPersonalFileNum, IndexRowLichnayaKarta].Value.ToString(); // Искомый № личного дела
            Required_PersonalNum = dataGridView1[IndexPersonalNum, IndexRowLichnayaKarta].Value.ToString(); // Искомый личный номер
            Required_Surname = dataGridView1[IndexSurname, IndexRowLichnayaKarta].Value.ToString(); // Искомая фамилия
            Required_Name = dataGridView1[IndexName, IndexRowLichnayaKarta].Value.ToString(); // Искомое имя
            Required_MiddleName = dataGridView1[IndexMiddleName, IndexRowLichnayaKarta].Value.ToString(); // Искомое отчество
            Required_DateOfBirth = dataGridView1[IndexDateOfBirth, IndexRowLichnayaKarta].Value.ToString(); // Искомая дата рождения
            int NumOfFinds = 0; // Сбрасываем количество найденных записей в ноль

            foreach (DataRow row in dataSet1.Tables[CurrentDataTableName].Rows) // Проходим по всем строкам активной DataTable
            {
                if ((row["PersonalFileNum"].ToString() == Required_PersonalFileNum) && (row["PersonalNum"].ToString() == Required_PersonalNum) && (row["Surname"].ToString() == Required_Surname)
                    && (row["Name"].ToString() == Required_Name) && (row["MiddleName"].ToString() == Required_MiddleName) && (row["DateOfBirth"].ToString() == Required_DateOfBirth))
                {
                    IndexOfRowToExport = dataSet1.Tables[CurrentDataTableName].Rows.IndexOf(row); // Записываем индекс искомой в Datatable строки
                    NumOfFinds++; // Увеличиваем количество найденных записей на единицу
                }
            }
            if (NumOfFinds == 0) MessageBox.Show("Похоже на ошибку в БД. Карточка не была перемещена в архив.");
            if (NumOfFinds == 1) // Если строка найдена и отсутствуют дубли
            {
                dataSet1.Tables[OtherDataTableName].ImportRow(dataSet1.Tables[CurrentDataTableName].Rows[IndexOfRowToExport]); // Импортируем найденную запись в другой DataTable
                dataSet1.Tables[CurrentDataTableName].Rows.RemoveAt(IndexOfRowToExport); // Удаляем найденную запись из текущего DataTable                
                AcceptAndWriteChanges(); // применяем изменения после перемещения строки из одного DataTable в другой
                if (CurrentDataTableName == "Kadry") MessageBox.Show("Карточка успешно перемещена в архив");
                else MessageBox.Show("Карточка успешно восстановлена из архива");

                this.RefreshDataGridView1(); // обновляем DataGridView1

                if (dataGridView1.Rows.Count != 0) // Проверка dataGridView1 на пустоту
                {
                    IndexRowLichnayaKarta = 0; // Делаем активной первую запись, дабы избежать проблемы с несуществующими индексами
                    this.PereschetCnum(); // пересчитываем порядковые номера
                    this.AcceptAndWriteChanges(); // сохраняем изменения
                    Cnum_label.Text = (IndexRowLichnayaKarta + 1).ToString() + " из " + dataGridView1.RowCount.ToString(); // Порядковый номер личной карточки
                    CardsFIO_label.Text = dataGridView1[IndexSurname, IndexRowLichnayaKarta].Value.ToString() + " "
            + dataGridView1[IndexName, IndexRowLichnayaKarta].Value.ToString() + " "
            + dataGridView1[IndexMiddleName, IndexRowLichnayaKarta].Value.ToString(); // Прописываем ФИО над стрелками в карточке
                }
                else Archive_Click(sender, e); // если грид пустой - переходим в другую базу
                tabControl1.SelectedTab = tabPage1; // Переходим на первую вкладку "Общий список"
            }
            // При необходимости, добавить сюда события при пустом гриде


            if (NumOfFinds > 1) MessageBox.Show("Существуют идентичные копии данной карточки. Измените их и попробуйте еще раз.");
        }
    }



    // ****************************************************************************************************
    // ****************************************************************************************************
    // ****************************************************************************************************
    // *****                                                                                          *****
    // *****    ПЕРЕПИСЫВАЕМ ОТРИСОВКУ ComboBox И DataGridViewComboBox, ДЛЯ ВЫРАВНИВАНИЯ ПО ЦЕНТРУ    *****
    // *****                                                                                          *****
    // ****************************************************************************************************
    // ****************************************************************************************************
    // ****************************************************************************************************
    class CenteredComboBox
    {

        public static void MyDGV_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control.GetType() == typeof(DataGridViewComboBoxEditingControl))
            {
                ComboBox cb = e.Control as ComboBox;
                if (cb != null)
                {
                    cb.DrawMode = DrawMode.OwnerDrawFixed;
                    cb.DropDownStyle = ComboBoxStyle.DropDownList;

                    cb.DrawItem += new DrawItemEventHandler(ComboBox_DrawItem_Centered);
                }
            }
        }

        // Центрируем ComboBox'ы
        public static void ComboBox_DrawItem_Centered(object sender, DrawItemEventArgs e)
        {
            // By using Sender, one method could handle multiple ComboBoxes
            ComboBox cbx = sender as ComboBox;
            if (cbx != null)
            {
                // Всегда рисуем задний фон
                e.DrawBackground();

                // Drawing one of the items?
                if (e.Index >= 0)
                {
                    // Установка положения строки (alignment). Допустимы значения Center, Near и Far
                    StringFormat sf = new StringFormat();
                    sf.LineAlignment = StringAlignment.Center;
                    sf.Alignment = StringAlignment.Center;

                    // Set the Brush to ComboBox ForeColor to maintain any ComboBox color settings
                    // Assumes Brush is solid
                    Brush brush = new SolidBrush(cbx.ForeColor);

                    // If drawing highlighted selection, change brush
                    if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                        brush = SystemBrushes.HighlightText;

                    // Отрисовка строки
                    e.Graphics.DrawString(cbx.Items[e.Index].ToString(), cbx.Font, brush, e.Bounds, sf);
                }
            }
        }
    }



    // ***********************************************************************************************************
    // ***********************************************************************************************************
    // ***********************************************************************************************************
    // *****                                                                                                 *****
    // *****    ПЕРЕПИСЫВАЕМ ОТРИСОВКУ DataGridViewComboBoxCell, ДЛЯ ВЕРТИКАЛЬНОГО ВЫРАВНИВАНИЯ ПО ЦЕНТРУ    *****
    // *****                                                                                                 *****
    // ***********************************************************************************************************
    // ***********************************************************************************************************
    // ***********************************************************************************************************

    public class DataGridViewComboBoxCellEx : DataGridViewComboBoxCell
    {
        private ComboBox _EditingComboBox; //Исходный ComboBox
        private Panel _Panel; //Панель оригинального ComboBox

        private enum COMBOBOXPARTS
        {
            CP_DROPDOWNBUTTON = 1,
            CP_BACKGROUND = 2,
            CP_TRANSPARENTBACKGROUND = 3,
            CP_BORDER = 4,
            CP_READONLY = 5,
            CP_DROPDOWNBUTTONRIGHT = 6,
            CP_DROPDOWNBUTTONLEFT = 7,
            CP_CUEBANNER = 8,
        }

        protected override void Paint
            (Graphics graphics
            , Rectangle clipBounds
            , Rectangle cellBounds
            , int rowIndex
            , DataGridViewElementStates elementState
            , object value
            , object formattedValue
            , string errorText
            , DataGridViewCellStyle cellStyle
            , DataGridViewAdvancedBorderStyle advancedBorderStyle
            , DataGridViewPaintParts paintParts)
        {

            // Есть ли курсор мыши над ячейкой?
            Rectangle rect = new Rectangle(this.DataGridView.PointToScreen(cellBounds.Location), cellBounds.Size);
            Point mousePos = Cursor.Position;
            bool isHot = rect.Contains(mousePos);

            // Ширина раскрывающейся кнопки и ее положение
            int buttonWidth = SystemInformation.VerticalScrollBarWidth + 3; // Этого должно хватить, чтобы скрыть исходную кнопку
            Rectangle border = this.BorderWidths(advancedBorderStyle);
            Rectangle buttonRect = new Rectangle(cellBounds.Right - buttonWidth, cellBounds.Top + border.Top, buttonWidth, cellBounds.Height - border.Top - border.Bottom);

            // Расстановка элементов
            StringFormat sf = new StringFormat();
            DataGridViewContentAlignment ali = DataGridViewContentAlignment.TopLeft;
            Brush background = Brushes.White;
            Brush textBrush = Brushes.Black;
            try
            {
                ali = this.InheritedStyle.Alignment;

                if (this.Selected)
                {
                    background = new SolidBrush(this.InheritedStyle.SelectionBackColor);
                    textBrush = new SolidBrush(this.InheritedStyle.SelectionForeColor);
                }
                else
                {
                    background = new SolidBrush(this.InheritedStyle.BackColor);
                    textBrush = new SolidBrush(this.InheritedStyle.ForeColor);
                }
            }
            catch
            {
            }
            switch (ali)
            {
                case DataGridViewContentAlignment.BottomCenter:
                    sf.LineAlignment = StringAlignment.Far;
                    sf.Alignment = StringAlignment.Center;
                    break;
                case DataGridViewContentAlignment.BottomLeft:
                    sf.LineAlignment = StringAlignment.Near;
                    sf.Alignment = StringAlignment.Near;
                    break;
                case DataGridViewContentAlignment.BottomRight:
                    sf.LineAlignment = StringAlignment.Far;
                    sf.Alignment = StringAlignment.Far;
                    break;
                case DataGridViewContentAlignment.MiddleCenter:
                    sf.LineAlignment = StringAlignment.Center;
                    sf.Alignment = StringAlignment.Center;
                    break;
                case DataGridViewContentAlignment.MiddleLeft:
                    sf.LineAlignment = StringAlignment.Center;
                    sf.Alignment = StringAlignment.Near;
                    break;
                case DataGridViewContentAlignment.MiddleRight:
                    sf.LineAlignment = StringAlignment.Center;
                    sf.Alignment = StringAlignment.Far;
                    break;
                case DataGridViewContentAlignment.NotSet:
                    sf.LineAlignment = StringAlignment.Near;
                    sf.Alignment = StringAlignment.Near;
                    break;
                case DataGridViewContentAlignment.TopCenter:
                    sf.LineAlignment = StringAlignment.Near;
                    sf.Alignment = StringAlignment.Center;
                    break;
                case DataGridViewContentAlignment.TopLeft:
                    sf.LineAlignment = StringAlignment.Near;
                    sf.Alignment = StringAlignment.Near;
                    break;
                case DataGridViewContentAlignment.TopRight:
                    sf.LineAlignment = StringAlignment.Near;
                    sf.Alignment = StringAlignment.Far;
                    break;
                default:
                    break;
            }

            graphics.FillRectangle(background, cellBounds);
            base.Paint(graphics, clipBounds, cellBounds, rowIndex, elementState, value, formattedValue, errorText, cellStyle, advancedBorderStyle, DataGridViewPaintParts.Border);

            //Различные методы рисования в зависимости от темы окна
            if (ComboBoxRenderer.IsSupported)
            {
                ComboBoxState state = isHot ? ComboBoxState.Hot : ComboBoxState.Normal;
                var render = new VisualStyleRenderer("COMBOBOX", (int)COMBOBOXPARTS.CP_READONLY, (int)state);
                //render.DrawBackground(graphics, cellBounds); //отвечает за отрисовку объемного заднего фона ячейки
                //ComboBoxRenderer.DrawDropDownButton(graphics, buttonRect, state); // в оригинале не закомментировано, но при плоской кнопке не нужно
                ControlPaint.DrawComboButton(graphics, buttonRect, ButtonState.Flat); // свойство Flat отвечает за вид кнопки
                textBrush = new SolidBrush(this.InheritedStyle.ForeColor);
            }
            else
            {
                //bool hasContent = ((paintParts & DataGridViewPaintParts.ContentForeground) != DataGridViewPaintParts.None);
                //paintParts = paintParts & ~DataGridViewPaintParts.ContentForeground;
                //base.Paint(graphics, clipBounds, cellBounds, rowIndex, elementState, value, formattedValue, errorText, cellStyle, advancedBorderStyle, paintParts);
                ControlPaint.DrawComboButton(graphics, buttonRect, ButtonState.Flat); // свойство Flat отвечает за вид кнопки
            }

            graphics.DrawString
                (formattedValue.ToString()
                , this.InheritedStyle.Font
                , textBrush
                , new RectangleF(cellBounds.Left, cellBounds.Top, cellBounds.Width - buttonWidth, cellBounds.Height)
                , sf);
        }

        protected override void OnMouseDown(DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex < 0 || e.RowIndex < 0 || this.OwningColumn.ReadOnly
                || this.DataGridView.CurrentCell != this)
            {
                return;
            }

            // Щелчок по ячейке
            var rect = this.DataGridView.GetCellDisplayRectangle(this.ColumnIndex, this.RowIndex, false);
            if (!this.IsInEditMode)
            {
                // Перевести в состояние редактирования, если еще не в нем
                this.DataGridView.BeginEdit(true);

                if (this.IsInEditMode)
                {// Находясь в состоянии редактирования, вынимаем элемент управления редактированием (Combobox) и регистрируем обработку событий и т.д.
                    SetEditingComboBox((ComboBox)this.DataGridView.EditingControl);

                }
            }
            else
            {
                base.OnMouseDown(e);
            }
        }

        private void SetEditingComboBox(ComboBox comboBox)
        {
            if (_EditingComboBox != null)
            {
                _EditingComboBox = null;
            }

            _EditingComboBox = comboBox;

            if (comboBox != null)
            {
                //Помещаем панель поверх оригинального ComboBox
                _Panel = new Panel();
                _Panel.Paint += _Panel_Paint;
                _Panel.Click += _Panel_Click;
                _Panel.MouseEnter += UpdatePanel;

                _EditingComboBox.Resize += UpdatePanel;
                _EditingComboBox.LocationChanged += UpdatePanel;
                _EditingComboBox.TextChanged += UpdatePanel;
                _EditingComboBox.VisibleChanged += UpdatePanel;
                _EditingComboBox.DrawMode = DrawMode.OwnerDrawVariable;
                _EditingComboBox.MeasureItem += _EditingComboBox_MeasureItem;
                _EditingComboBox.DrawItem += _EditingComboBox_DrawItem;
                this.DataGridView.Scroll += UpdatePanel;
                this.DataGridView.RowHeightChanged += UpdatePanel;
                this.DataGridView.ColumnWidthChanged += UpdatePanel;
                this.DataGridView.Controls.Add(_Panel);

                this.DataGridView.CellEndEdit += DataGridView_CellEndEdit;
                UpdatePanel();
            }
            _EditingComboBox.DroppedDown = true;
        }

        /// <summary>Размер элемента DropDown</summary>
        void _EditingComboBox_MeasureItem(object sender, MeasureItemEventArgs e)
        {
            //e.ItemHeight = this.OwningRow.Height;
            e.ItemWidth = this.OwningColumn.Width;
        }
        void _EditingComboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            Brush brush;
            //Font font; // в оригинале не закомментировано, но нигде не используется
            ComboBox combo = (ComboBox)sender;
            e.DrawBackground();

            if ((e.State & DrawItemState.Selected) != DrawItemState.None)
            {
                e.DrawFocusRectangle();
                brush = new SolidBrush(this.InheritedStyle.SelectionForeColor);
            }
            else
            {
                brush = new SolidBrush(this.InheritedStyle.ForeColor);
            }

            string text;
            if (e.Index >= 0)
            {
                text = combo.Items[e.Index].ToString();
            }
            else
            {
                text = combo.Text;
            }
            //int w = this.OwningColumn.Width; // в оригинале не закомментировано, но нигде не используется
            //int h = this.OwningRow.Height; // в оригинале не закомментировано, но нигде не используется
            StringFormat sf = new StringFormat();
            sf.LineAlignment = StringAlignment.Center;
            sf.Alignment = StringAlignment.Center; // задает положение текста в выпадающем списке в режиме редактирования ComboBox'а
                                                   //e.Graphics.DrawString(combo.Items[e.Index].ToString(), this.InheritedStyle.Font, Brushes.Black, new RectangleF(0, 0, w, h), sf);
            e.Graphics.DrawString
                (text
                , this.InheritedStyle.Font
                , brush
                , e.Bounds, sf);
        }

        void _Panel_Click(object sender, EventArgs e)
        {
            _EditingComboBox.DroppedDown = !_EditingComboBox.DroppedDown;
        }

        private void UpdatePanel(object sender, EventArgs e)
        {
            UpdatePanel();
        }

        // Перемещаем наложенную панель и выравниваем исходное поле со списком по ячейке
        private void UpdatePanel()
        {
            if (_Panel != null)
            {
                var rect = this.DataGridView.GetCellDisplayRectangle(this.ColumnIndex, this.RowIndex, false);

                _Panel.Location = rect.Location;
                _Panel.Size = rect.Size;
                _Panel.BringToFront();
                _EditingComboBox.Top = rect.Height - _EditingComboBox.Height;
                _EditingComboBox.DropDownHeight = rect.Height * 10; // по умолчанию было *5, но при этом в выпадающем списке была полоса прокрутки

                _Panel.Refresh();
            }
        }

        // Пусть закрытая панель нарисует изображение поля со списком
        void _Panel_Paint(object sender, PaintEventArgs e)
        {
            if (_EditingComboBox != null)
            {
                var gs = e.Graphics.Save();
                var rect = this.DataGridView.GetCellDisplayRectangle(this.ColumnIndex, this.RowIndex, false);

                Rectangle bounds = e.ClipRectangle;
                this.Paint
                    (e.Graphics
                    , bounds
                    , bounds
                    , this.RowIndex
                    , DataGridViewElementStates.Selected
                    , _EditingComboBox.Text
                    , _EditingComboBox.Text
                    , string.Empty
                    , this.InheritedStyle
                    , new DataGridViewAdvancedBorderStyle() { All = DataGridViewAdvancedCellBorderStyle.Single }
                    , DataGridViewPaintParts.All);

                e.Graphics.Restore(gs);
            }
        }

        // Отписываемся от событий
        void DataGridView_CellEndEdit(object sender, EventArgs e)
        {
            this.DataGridView.CellEndEdit -= DataGridView_CellEndEdit;

            _EditingComboBox.Resize -= UpdatePanel;
            _EditingComboBox.LocationChanged -= UpdatePanel;
            _EditingComboBox.TextChanged -= UpdatePanel;
            _EditingComboBox.VisibleChanged -= UpdatePanel;
            _EditingComboBox.MeasureItem -= _EditingComboBox_MeasureItem;
            _EditingComboBox.DrawItem -= _EditingComboBox_DrawItem;
            _EditingComboBox.DrawMode = DrawMode.Normal;
            _EditingComboBox = null;

            _Panel.Parent.Controls.Remove(_Panel);
            _Panel.Paint -= _Panel_Paint;
            _Panel.Click -= _Panel_Click;
            _Panel = null;

            this.DataGridView.CellEndEdit -= DataGridView_CellEndEdit;
            this.DataGridView.Scroll -= UpdatePanel;
            this.DataGridView.RowHeightChanged -= UpdatePanel;
            this.DataGridView.ColumnWidthChanged -= UpdatePanel;
            this.DataGridView.InvalidateCell(this);
        }
    }

    // ###############  ПЕРЕНАЗНАЧАЕМ ДЕЙСТВИЯ В СВОЁМ КАСТОМНОМ dataGridView  ###############
    public class MyCustomDataGrid : DataGridView
    {
        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // клавиша обработана
                e.Handled = true;

                // SendKeys.Send(" ");
            }
            base.OnKeyDown(e);
        }

        protected override bool ProcessDialogKey(Keys keyData)
        {
            // Extract the key code from the key value. 
            Keys key = (keyData & Keys.KeyCode);

            // Handle the ENTER key as if it were a RIGHT ARROW key. 
            if (key == Keys.Enter)
            {
                return this.ProcessRightKey(keyData);
            }
            return base.ProcessDialogKey(keyData);
        }
    }
    /*
    public class CenteredDateTimePicker : DateTimePicker
    {
        public CenteredDateTimePicker()
        {
            SetStyle(ControlStyles.UserPaint, true);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.DrawString(Text, Font, new SolidBrush(ForeColor), ClientRectangle, new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            });
        }
    }
    */
}
