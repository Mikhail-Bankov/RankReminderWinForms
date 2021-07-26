using System;
using System.Linq;
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

        // Индексы колонок listVisible (грид только с "видимыми" колонками)
        int IndexVisibleCnum; // Порядковый номер
        int IndexVisibleDateOfBirth; // Дата рождения
        int IndexVisibleRank; // Звание
        int IndexVisibleRankDate; // Дата присвоения звания
        int IndexVisibleRankLimit; // Потолок по званию
        int IndexVisibleNextRankDate; // Следующая дата присвоения звания
        int IndexVisibleKlassnost; // Квалификационное звание
        int IndexVisibleKlassnostDate; // Дата присвоения квалиф. звания
        int IndexVisibleNextKlassnostDate; // Следующая дата присвоения квалиф. звания


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
        DataGridViewTextBoxColumn nextattestaciyadate = new DataGridViewTextBoxColumn(); // Дата следующей аттестации
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

            dataSet1.ReadXml(XMLDB.Path); // считываем в dataSet1 нашу базу в формате XML

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

            // ###############  РАЗНЫЕ ОБРАБОТЧИКИ СОБЫТИЙ  ###############
            dataSet1.Tables["Kadry"].RowDeleting += new System.Data.DataRowChangeEventHandler(RowDeleting); // обработчик события попытки удаления строки
            dataSet1.Tables["Kadry"].RowDeleted += new System.Data.DataRowChangeEventHandler(RowDeleted); // обработчик события удаления строки                                                                                         
            dataSet1.Tables["Archive"].RowDeleting += new System.Data.DataRowChangeEventHandler(RowDeleting); // обработчик события попытки удаления строки
            dataSet1.Tables["Archive"].RowDeleted += new System.Data.DataRowChangeEventHandler(RowDeleted); // обработчик события удаления строки     
            dataGridView1.Sorted += new System.EventHandler(dataGridView1_Sorted); // обработчик события сортировки колонки
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
            dataGridView_Married.CellValueChanged += new DataGridViewCellEventHandler(SaveChangesToDataGridView_Married);
            dataGridView_Family.CellValueChanged += new DataGridViewCellEventHandler(SaveChangesToDataGridView_Family);
            dataGridView_Study.CellValueChanged += new DataGridViewCellEventHandler(SaveChangesToDataGridView_Study);
            dataGridView_UchStepen.CellValueChanged += new DataGridViewCellEventHandler(SaveChangesToDataGridView_UchStepen);
            dataGridView_PrisvZvaniy.CellValueChanged += new DataGridViewCellEventHandler(SaveChangesToDataGridView_PrisvZvaniy);
            dataGridView_TrudDeyat.CellValueChanged += new DataGridViewCellEventHandler(SaveChangesToDataGridView_TrudDeyat);
            dataGridView_StazhVysluga.CellValueChanged += new DataGridViewCellEventHandler(SaveChangesToDataGridView_StazhVysluga);
            dataGridView_RabotaGFS.CellValueChanged += new DataGridViewCellEventHandler(SaveChangesToDataGridView_RabotaGFS);
            dataGridView_Attestaciya.CellValueChanged += new DataGridViewCellEventHandler(SaveChangesToDataGridView_Attestaciya);
            dataGridView_ProfPodg.CellValueChanged += new DataGridViewCellEventHandler(SaveChangesToDataGridView_ProfPodg);
            dataGridView_KlassnostOld.CellValueChanged += new DataGridViewCellEventHandler(SaveChangesToDataGridView_KlassnostOld);
            dataGridView_Nagrady.CellValueChanged += new DataGridViewCellEventHandler(SaveChangesToDataGridView_Nagrady);
            dataGridView_Prodlenie.CellValueChanged += new DataGridViewCellEventHandler(SaveChangesToDataGridView_Prodlenie);
            dataGridView_Boevye.CellValueChanged += new DataGridViewCellEventHandler(SaveChangesToDataGridView_Boevye);
            dataGridView_Rezerv.CellValueChanged += new DataGridViewCellEventHandler(SaveChangesToDataGridView_Rezerv);
            dataGridView_Vzyskaniya.CellValueChanged += new DataGridViewCellEventHandler(SaveChangesToDataGridView_Vzyskaniya);
            dataGridView_Uvolnenie.CellValueChanged += new DataGridViewCellEventHandler(SaveChangesToDataGridView_Uvolnenie);

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

            CardsFIO_label.Text = dataGridView1[IndexSurname, IndexRowLichnayaKarta].Value.ToString() + " "
                + dataGridView1[IndexName, IndexRowLichnayaKarta].Value.ToString() + " "
                + dataGridView1[IndexMiddleName, IndexRowLichnayaKarta].Value.ToString(); // Прописываем ФИО над стрелками в карточке

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
                    dataGridView1[IndexCnum, i].Value = i + 1; //увеличиваем порядковый номер в каждой последующей строке на единицу
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
            this.ShowAllColumns();
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
            this.ShowAllRows();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            this.ShowVysluga();
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
            this.ShowAllRows();
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            this.ShowKlassnost();
        }

        private void ShowKlassnost() // показать только классность
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
            this.ShowAllRows();
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            this.ShowAttestaciya();
        }

        private void ShowAttestaciya() // показать только тех, у кого аттестация в следующем году (проводится каждые 4 года)
        {
            if (radioButton4.Checked == true) //без этой проверки, при radioButton1.Checked = true, метод срабатывает второй раз (позже разобраться почему)
            {
                int ProverkaNaPustotu = 0;
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
                            ProverkaNaPustotu++;
                        }
                    }
                    else
                    {
                        dataGridView1.CurrentCell = null;
                        row.Visible = false; // скрываем строку
                        ProverkaNaPustotu++;
                    }
                }
                if (ProverkaNaPustotu == dataGridView1.Rows.Count) // Если по итогу перебора строк, все строки оказались скрыты
                {
                    MessageBox.Show("Сотрудники, подпадающие под данный фильтр отсутствуют!");
                    radioButton1.Checked = true; // сбрасываем выбор фильтра
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
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // ####################################
        // ##  КНОПКА "ДОБАВИТЬ СОТРУДНИКА"  ##
        // ####################################
        private void button3_Click(object sender, EventArgs e)
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
        private void Archive_button_Click(object sender, EventArgs e)
        {
            if (CurrentDataTableName == "Kadry")
            {
                Archive_button.Text = "Действующие сотрудники";
                CurrentDataTableName = "Archive";
                OtherDataTableName = "Kadry";
                CurrentBase_label.Text = "Текущая БД: Архивные сотрудники";
                MessageBox.Show("Перешли в архивную базу данных");
            }
            else if (CurrentDataTableName == "Archive")
            {
                Archive_button.Text = "Архивные сотрудники";
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
        private void button4_Click(object sender, EventArgs e)
        {
            this.ExportDataGridToExcel();
        }

        // ###############  ВЫГРУЗКА dataGridView1 В EXCEL ФАЙЛ  ###############
        public void ExportDataGridToExcel()
        {
            //Формируем новую таблицу только из видимых столбцов
            List<DataGridViewColumn> listVisible = new List<DataGridViewColumn>();
            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                if (col.Visible)
                {
                    listVisible.Add(col);
                }
            }

            Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Application.Workbooks.Add(Type.Missing);
            ExcelApp.Columns.ColumnWidth = 15; // устанавливаем ширину столбцов
            ExcelApp.Cells.WrapText = "true"; // устанавливаем перенос по словам

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)ExcelApp.Worksheets.get_Item(1);

            var _with1 = xlWorkSheet.PageSetup; // блок параметров листа
            _with1.PaperSize = Excel.XlPaperSize.xlPaperA4; // размер А4
            _with1.Orientation = Excel.XlPageOrientation.xlLandscape; // ландшафтная ориентация
            _with1.Zoom = false;
            _with1.FitToPagesWide = 1;
            _with1.FitToPagesTall = 1;

            xlWorkSheet.Name = "Сведения о личном составе"; // именуем лист

            Excel.Range range1 = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, listVisible.Count]); // диапазон заголовка в файле Excel
            range1.Cells.Font.Bold = true; // жирный шрифт
            range1.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // горизонтальное выравнивание по центру
            range1.BorderAround(Type.Missing, Excel.XlBorderWeight.xlThick, Excel.XlColorIndex.xlColorIndexAutomatic); // увеличиваем толщину внешних границ
            range1.Borders.Color = Color.Black; // черный цвет границ

            for (int i = 0; i < listVisible.Count; i++) // Заполняем заголовок в Excel файле заголовками столбцов
            {
                ExcelApp.Cells[1, i + 1] = listVisible[i].HeaderText; // заполняется строго первая строка

                switch (listVisible[i].HeaderText) // просчитываем индексы в "видимом" гриде
                {
                    case Cnum_HeaderText: // Порядковый номер
                        IndexVisibleCnum = i;
                        break;
                    case DateOfBirth_HeaderText: // Дата рождения
                        IndexVisibleDateOfBirth = i;
                        break;
                    case Rank_HeaderText: // Звание
                        IndexVisibleRank = i;
                        break;
                    case RankDate_HeaderText: // Дата присвоения звания
                        IndexVisibleRankDate = i;
                        break;
                    case RankLimit_HeaderText: // Потолок по званию
                        IndexVisibleRankLimit = i;
                        break;
                    case NextRankDate_HeaderText: // Следующая дата присвоения звания
                        IndexVisibleNextRankDate = i;
                        break;
                    case Klassnost_HeaderText: // Квалификационное звание
                        IndexVisibleKlassnost = i;
                        break;
                    case KlassnostDate_HeaderText: // Дата присвоения квалиф. звания
                        IndexVisibleKlassnostDate = i;
                        break;
                    case NextKlassnostDate_HeaderText: // Следующая дата присвоения квалиф. звания
                        IndexVisibleNextKlassnostDate = i;
                        break;
                }
            }

            Excel.Range range2 = xlWorkSheet.get_Range(xlWorkSheet.Columns[1], xlWorkSheet.Columns[listVisible.Count]);// диапазон всех рабочих колонок
            range2.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter; // вертикальное выравнивание по центру

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < listVisible.Count; j++)
                {
                    ExcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[listVisible[j].Index].Value.ToString();
                }
            }

            Excel.Range range3 = xlWorkSheet.get_Range(xlWorkSheet.Cells[2, 1], xlWorkSheet.Cells[dataGridView1.RowCount + 1, listVisible.Count]);
            range3.Borders[Excel.XlBordersIndex.xlInsideVertical].Color = Color.LightGray; //внутренние вертикальные границы области с данными
            range3.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Color = Color.Black; //внутренние горизонтальные границы области с данными
            range3.Borders[Excel.XlBordersIndex.xlEdgeRight].Color = Color.Black; //крайняя правая граница области с данными
            range3.Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = Color.Black; //крайняя левая граница области с данными
            range3.Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = Color.Black; //крайняя нижняя граница области с данными

            // диапазон ячеек с порядковым номером
            Excel.Range rangeCnum = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, IndexVisibleCnum + 1], xlWorkSheet.Cells[dataGridView1.RowCount + 1, IndexVisibleCnum + 1]);
            rangeCnum.ColumnWidth = 5; // уменьшаем ширину
            rangeCnum.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // горизонтальное выравнивание по центру

            // ВЫВОД В EXCEL В ЗАВИСИМОСТИ ОТ ВЫБРАННОГО ФИЛЬТРА
            if (radioButton1.Checked == true) //если фильтрация выключена
            {
                // Диапазон "Дата рождения"
                Excel.Range rangeVisibleDateOfBirth = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, IndexVisibleDateOfBirth + 1], xlWorkSheet.Cells[dataGridView1.RowCount + 1, IndexVisibleDateOfBirth + 1]);
                rangeVisibleDateOfBirth.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // горизонтальное выравнивание по центру
                // Диапазон "Дата присвоения звания"
                Excel.Range rangeVisibleRankDate = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, IndexVisibleRankDate + 1], xlWorkSheet.Cells[dataGridView1.RowCount + 1, IndexVisibleRankDate + 1]);
                rangeVisibleRankDate.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // горизонтальное выравнивание по центру
                // Диапазон "Следующая дата присвоения звания"
                Excel.Range rangeVisibleNextRankDate = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, IndexVisibleNextRankDate + 1], xlWorkSheet.Cells[dataGridView1.RowCount + 1, IndexVisibleNextRankDate + 1]);
                rangeVisibleNextRankDate.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // горизонтальное выравнивание по центру
                // Диапазон "Квалификационное звание"
                Excel.Range rangeVisibleKlassnost = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, IndexVisibleKlassnost + 1], xlWorkSheet.Cells[dataGridView1.RowCount + 1, IndexVisibleKlassnost + 1]);
                rangeVisibleKlassnost.ColumnWidth = 20; // увеличиваем ширину
                // Диапазон "Дата присвоения квалиф. звания"
                Excel.Range rangeVisibleKlassnostDate = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, IndexVisibleKlassnostDate + 1], xlWorkSheet.Cells[dataGridView1.RowCount + 1, IndexVisibleKlassnostDate + 1]);
                rangeVisibleKlassnostDate.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // горизонтальное выравнивание по центру
                // Диапазон "Следующая дата присвоения квалиф. звания"
                Excel.Range rangeNextKlassnostDate = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, IndexVisibleNextKlassnostDate + 1], xlWorkSheet.Cells[dataGridView1.RowCount + 1, IndexVisibleNextKlassnostDate + 1]);
                rangeNextKlassnostDate.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // горизонтальное выравнивание по центру
            }

            if (radioButton2.Checked == true) //если фильтрация по званию
            {
                // Диапазон "Дата присвоения звания"
                Excel.Range rangeVisibleRankDate = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, IndexVisibleRankDate + 1], xlWorkSheet.Cells[dataGridView1.RowCount + 1, IndexVisibleRankDate + 1]);
                rangeVisibleRankDate.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // горизонтальное выравнивание по центру
                // Диапазон "Следующая дата присвоения звания"
                Excel.Range rangeVisibleNextRankDate = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, IndexVisibleNextRankDate + 1], xlWorkSheet.Cells[dataGridView1.RowCount + 1, IndexVisibleNextRankDate + 1]);
                rangeVisibleNextRankDate.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // горизонтальное выравнивание по центру
            }

            if (radioButton3.Checked == true) //если фильтрация по классности
            {
                // Диапазон "Квалификационное звание"
                Excel.Range rangeVisibleKlassnost = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, IndexVisibleKlassnost + 1], xlWorkSheet.Cells[dataGridView1.RowCount + 1, IndexVisibleKlassnost + 1]);
                rangeVisibleKlassnost.ColumnWidth = 20; // увеличиваем ширину
                // Диапазон "Дата присвоения квалиф. звания"
                Excel.Range rangeVisibleKlassnostDate = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, IndexVisibleKlassnostDate + 1], xlWorkSheet.Cells[dataGridView1.RowCount + 1, IndexVisibleKlassnostDate + 1]);
                rangeVisibleKlassnostDate.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // горизонтальное выравнивание по центру
                // Диапазон "Следующая дата присвоения квалиф. звания"
                Excel.Range rangeNextKlassnostDate = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, IndexVisibleNextKlassnostDate + 1], xlWorkSheet.Cells[dataGridView1.RowCount + 1, IndexVisibleNextKlassnostDate + 1]);
                rangeNextKlassnostDate.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // горизонтальное выравнивание по центру
            }
            if (radioButton4.Checked == true) //если фильтрация по аттестации
            {
                // здесь будет находится код для вывода в Excel списка с аттестацией сотрудников
            }
            ExcelApp.Visible = true;
        }


        // ###############  ДЕЙСТВИЯ ПРИ СРАБАТЫВАНИИ СОБЫТИЯ СОРТИРОВКИ  ###############
        private void dataGridView1_Sorted(object sender, EventArgs e) //отработка события изменения сортировки
        {
            this.PereschetCnum();
        }

        // ###############  ДЕЙСТВИЯ, ЕСЛИ БЫЛИ КАКИЕ-ЛИБО ИЗМЕНЕНИЯ В dataGridView1  ###############
        public void DataGridWasChanged()
        {
            MessageBox.Show("grid изменен");
            this.PereschetZvanie();
            this.PereschetKlassnost();
            this.AcceptAndWriteChanges(); // сохраняем изменения в XML
            this.RefreshDataGridView1(); // обновляем DataGridView1
        }

        // ###############  ОБНОВЛЕНИЕ dataGridView1  ###############
        public void RefreshDataGridView1()
        {
            dataSet1.Clear(); // очищаем dataSet1
            dataGridView1.DataSource = null; // очищаем DataSource
            dataSet1.ReadXml(XMLDB.Path); //  считываем XML
            dataGridView1.DataSource = dataSet1.Tables[CurrentDataTableName]; // присваиваем DataSource
        }

        // ###############  ПРИМЕНИТЬ ВСЕ ИЗМЕНЕНИЯ И СОХРАНИТЬ XML  ###############
        public void AcceptAndWriteChanges()
        {
            MessageBox.Show("Произошло сохранение базы данных"); // позже будет закомментировано 
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


        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В TextBox'ах НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########

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

            string StudyStringFromCurrentCell = dataGridView1[IndexStudy, IndexRowLichnayaKarta].Value.ToString();
            if (StudyStringFromCurrentCell != "") //проверка на существование данных в таблице
            {
                string[] study_array = StudyStringFromCurrentCell.Split('$');

                foreach (string s in study_array)
                {
                    string[] StudyRow = s.Split(',');
                    dataGridView_Study.Rows.Add(StudyRow);
                }
            }
            Study_FormaObucheniya.MinimumWidth = 120;

            //Чтобы не обрезался текст, расчитываем ширину выпадающего списка, когда ComboBox в режиме редактирования
            Study_FormaObucheniya.DropDownWidth = Study_FormaObucheniya.Items.Cast<Object>().Select(x => x.ToString())
    .Max(x => TextRenderer.MeasureText(x, Study_FormaObucheniya.InheritedStyle.Font, Size.Empty, TextFormatFlags.Default).Width);

            Study_Naimenovanie.MinimumWidth = 130;
            Study_DataPost.MinimumWidth = 120;
            Study_DataPost.Width = 120;
            Study_DataOkonch.MinimumWidth = 120;
            Study_DataOkonch.Width = 120;
            Study_Document.MinimumWidth = 140;

            /*==============================================================================================================*/
            /*==============================================================================================================*/

            dataGridView_UchStepen.Rows.Clear();
            dataGridView_UchStepen.AutoGenerateColumns = false;
            string UchStepenStringFromCurrentCell = dataGridView1[IndexUchStepen, IndexRowLichnayaKarta].Value.ToString();
            if (UchStepenStringFromCurrentCell != "") //проверка на существование данных в таблице
            {
                string[] uchstepen_array = UchStepenStringFromCurrentCell.Split(';');

                foreach (string s in uchstepen_array)
                {
                    string[] UchStepenRow = s.Split('^');
                    dataGridView_UchStepen.Rows.Add(UchStepenRow);
                }
            }
        }

        // ###################################################################
        // ##  КНОПКА "ДОБАВИТЬ УЧЕНУЮ СТЕПЕНЬ" НА ВКЛАДКЕ "КАРТОЧКА 10-11" ##
        // ###################################################################
        private void UchStepenAdd_button_Click(object sender, EventArgs e)
        {
            dataGridView_UchStepen.Rows.Add("---", DateTime.Now.ToString("dd.MM.yyyy")); // добавить ученую степень
            this.SaveChangesToDataGridView_UchStepen(sender, e);
        }

        // ################################################################
        // ##  КНОПКА "ДОБАВИТЬ ОБРАЗОВАНИЕ" НА ВКЛАДКЕ "КАРТОЧКА 10-11" ##
        // ################################################################
        private void StudyAdd_button_Click(object sender, EventArgs e)
        {
            dataGridView_Study.Rows.Add("Высшее (очное)", "---", DateTime.Now.ToString("dd.MM.yyyy"), DateTime.Now.ToString("dd.MM.yyyy"), "---", "---", "---"); // добавить образование
            this.SaveChangesToDataGridView_Study(sender, e);
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В DataGridView_UchStepen НА ВКЛАДКЕ "КАРТОЧКА 10-11" ##########
        private void SaveChangesToDataGridView_UchStepen(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder(); // создаем строку для построения
            foreach (DataGridViewRow row in dataGridView_UchStepen.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    // Обработка неверного формата даты. Возможно так и должно быть. Разобраться позже.
                    if (cell.OwningColumn.Name == UchStepenDataPrisuzhdeniya.Name)
                    {
                        DateTime UchStepen_wrongdatetoconvert = DateTime.Parse(cell.Value.ToString()); // парсим её в формат DateTime
                        cell.Value = UchStepen_wrongdatetoconvert.ToString("dd.MM.yyyy"); // и конвертируем в нужный формат
                    }
                    sb.Append(cell.Value); // добавляем значение ячейки
                    sb.Append("^"); // ставим разделитель ячеек 
                }
                sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель ячеек
                sb.Append(";"); // ставим разделитель строки
            }
            sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель строки

            dataGridView1[IndexUchStepen, IndexRowLichnayaKarta].Value = sb.ToString(); // Заполняем ячейку результирующей строкой
            this.AcceptAndWriteChanges(); // Применить изменения
        }

        // ###########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В DataGridView_Study НА ВКЛАДКЕ "КАРТОЧКА 10-11" ###########
        private void SaveChangesToDataGridView_Study(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder(); // создаем строку для построения
            foreach (DataGridViewRow row in dataGridView_Study.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    // Обработка неверного формата даты. Возможно так и должно быть. Разобраться позже.
                    if ((cell.OwningColumn.Name == Study_DataPost.Name) || (cell.OwningColumn.Name == Study_DataOkonch.Name))
                    {
                        DateTime Study_wrongdatetoconvert = DateTime.Parse(cell.Value.ToString()); // парсим её в формат DateTime
                        cell.Value = Study_wrongdatetoconvert.ToString("dd.MM.yyyy"); // и конвертируем в нужный формат
                    }
                    sb.Append(cell.Value); // добавляем значение ячейки
                    sb.Append(","); // ставим разделитель ячеек 
                }
                sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель ячеек
                sb.Append("$"); // ставим разделитель строки 
            }
            sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель строки

            dataGridView1[IndexStudy, IndexRowLichnayaKarta].Value = sb.ToString(); // Заполняем ячейку результирующей строкой
            this.AcceptAndWriteChanges(); // Применить изменения
        }



        //               //"""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 12"  ############################################################
        public void UpdateCard12()
        {
            dataGridView_PrisvZvaniy.Rows.Clear();
            dataGridView_PrisvZvaniy.AutoGenerateColumns = false;
            string PrisvZvaniyStringFromCurrentCell = dataGridView1[IndexPrisvZvaniy, IndexRowLichnayaKarta].Value.ToString();
            if (PrisvZvaniyStringFromCurrentCell != "") //проверка на существование данных в таблице
            {
                string[] prisvzvaniy_array = PrisvZvaniyStringFromCurrentCell.Split('$');

                foreach (string s in prisvzvaniy_array)
                {
                    string[] PrisvZvaniyRow = s.Split('^');
                    dataGridView_PrisvZvaniy.Rows.Add(PrisvZvaniyRow);
                }
            }
        }

        // ######################################################################
        // ##  КНОПКА "ДОБАВИТЬ ЗВАНИЕ, КЛАССНЫЙ ЧИН" НА ВКЛАДКЕ "КАРТОЧКА 12" ##
        // ######################################################################
        private void ZvanieAdd_button_Click(object sender, EventArgs e)
        {
            dataGridView_PrisvZvaniy.Rows.Add("---", DateTime.Now.ToString("dd.MM.yyyy"), "---", "---", DateTime.Now.ToString("dd.MM.yyyy")); // добавить звание, классный чин
            this.SaveChangesToDataGridView_PrisvZvaniy(sender, e);
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В DataGridView_PrisvZvaniy НА ВКЛАДКЕ "КАРТОЧКА 12" ##########
        private void SaveChangesToDataGridView_PrisvZvaniy(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder(); // создаем строку для построения
            foreach (DataGridViewRow row in dataGridView_PrisvZvaniy.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    // Обработка неверного формата даты. Возможно так и должно быть. Разобраться позже.
                    if ((cell.OwningColumn.Name == PrisvZvaniy_DataPrisv.Name) || (cell.OwningColumn.Name == PrisvZvaniy_DataPrikaza.Name))
                    {
                        DateTime PrisvZvaniy_wrongdatetoconvert = DateTime.Parse(cell.Value.ToString()); // парсим её в формат DateTime
                        cell.Value = PrisvZvaniy_wrongdatetoconvert.ToString("dd.MM.yyyy"); // и конвертируем в нужный формат
                    }
                    sb.Append(cell.Value); // добавляем значение ячейки
                    sb.Append("^"); // ставим разделитель ячеек
                }
                sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель ячеек
                sb.Append("$"); // ставим разделитель строки
            }
            sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель строки

            dataGridView1[IndexPrisvZvaniy, IndexRowLichnayaKarta].Value = sb.ToString(); // Заполняем ячейку результирующей строкой
            this.AcceptAndWriteChanges(); // Применить изменения
        }



        //               //""""""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 13-14"  ############################################################
        public void UpdateCard13and14()
        {
            dataGridView_Married.Rows.Clear();
            dataGridView_Married.AutoGenerateColumns = false;
            string MarriedStringFromCurrentCell = dataGridView1[IndexMarried, IndexRowLichnayaKarta].Value.ToString();

            if (MarriedStringFromCurrentCell != "") //проверка на существование данных в таблице
            {
                string[] married_array = MarriedStringFromCurrentCell.Split('$');

                foreach (string s in married_array)
                {
                    string[] MarriedRow = s.Split('^');
                    dataGridView_Married.Rows.Add(MarriedRow);
                }
            }

            /*==============================================================================================================*/
            /*==============================================================================================================*/

            dataGridView_Family.Rows.Clear();
            dataGridView_Family.AutoGenerateColumns = false;
            string FamilyStringFromCurrentCell = dataGridView1[IndexFamily, IndexRowLichnayaKarta].Value.ToString();

            if (FamilyStringFromCurrentCell != "") //проверка на существование данных в таблице
            {
                string[] family_array = FamilyStringFromCurrentCell.Split('$');

                foreach (string s in family_array)
                {
                    string[] FamilyRow = s.Split('^');
                    dataGridView_Family.Rows.Add(FamilyRow);
                }
            }
            Family_StepenRodstva.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // Выравнивание по центру в колонке "Степень родства"
            Family_DateOfBirth.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // Выравнивание по центру в колонке "Дата рождения"
            Family_DateOfBirth.MinimumWidth = 120;
        }

        // ############################################################
        // ##  КНОПКА "ДОБАВИТЬ СОБЫТИЕ" НА ВКЛАДКЕ "КАРТОЧКА 13-14" ##
        // ############################################################
        private void MarriedAdd_button_Click(object sender, EventArgs e)
        {
            dataGridView_Married.Rows.Add("Женат", DateTime.Now.ToString("yyyy")); // добавить событие (свадьба, развод)
            this.SaveChangesToDataGridView_Married(sender, e);
        }

        // ################################################################
        // ##  КНОПКА "ДОБАВИТЬ ЧЛЕНА СЕМЬИ" НА ВКЛАДКЕ "КАРТОЧКА 13-14" ##
        // ################################################################
        private void FamilyAddPerson_button_Click(object sender, EventArgs e)
        {
            dataGridView_Family.Rows.Add("Мать", DateTime.Now.ToString("dd.MM.yyyy"), "---"); // добавить члена семьи
            this.SaveChangesToDataGridView_Family(sender, e);
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В DataGridView_Married НА ВКЛАДКЕ "КАРТОЧКА 13-14" ##########
        private void SaveChangesToDataGridView_Married(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder(); // создаем строку для построения
            foreach (DataGridViewRow row in dataGridView_Married.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    sb.Append(cell.Value); // добавляем значение ячейки
                    sb.Append("^"); // ставим ставим разделитель ячеек
                }
                sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель ячеек
                sb.Append("$"); // ставим разделитель строки
            }
            sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель строки

            dataGridView1[IndexMarried, IndexRowLichnayaKarta].Value = sb.ToString(); // Заполняем ячейку результирующей строкой
            this.AcceptAndWriteChanges(); // Применить изменения
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В DataGridView_Family НА ВКЛАДКЕ "КАРТОЧКА 13-14" ##########
        private void SaveChangesToDataGridView_Family(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder(); // создаем строку для построения
            foreach (DataGridViewRow row in dataGridView_Family.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    // Обработка неверного формата даты. Возможно так и должно быть. Разобраться позже.
                    if (cell.OwningColumn.Name == Family_DateOfBirth.Name)
                    {
                        DateTime Family_wrongdatetoconvert = DateTime.Parse(cell.Value.ToString()); // парсим её в формат DateTime
                        cell.Value = Family_wrongdatetoconvert.ToString("dd.MM.yyyy"); // и конвертируем в нужный формат
                    }
                    sb.Append(cell.Value); // добавляем значение ячейки
                    sb.Append("^"); // ставим разделитель ячеек
                }
                sb.Remove(sb.Length - 1, 1); // Убираем последнюю запятую
                sb.Append("$"); // ставим разделитель строки
            }
            sb.Remove(sb.Length - 1, 1); // Убираем последнюю точку с запятой

            dataGridView1[IndexFamily, IndexRowLichnayaKarta].Value = sb.ToString(); // Заполняем ячейку результирующей строкой
            this.AcceptAndWriteChanges(); // Применить изменения
        }



        //               //"""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 15"  ############################################################
        public void UpdateCard15()
        {
            dataGridView_TrudDeyat.Rows.Clear();
            dataGridView_TrudDeyat.AutoGenerateColumns = false;
            string TrudDeyatStringFromCurrentCell = dataGridView1[IndexTrudDeyat, IndexRowLichnayaKarta].Value.ToString();
            if (TrudDeyatStringFromCurrentCell != "") //проверка на существование данных в таблице
            {
                string[] truddeyat_array = TrudDeyatStringFromCurrentCell.Split('$');

                foreach (string s in truddeyat_array)
                {
                    string[] TrudDeyatRow = s.Split('^');
                    dataGridView_TrudDeyat.Rows.Add(TrudDeyatRow);
                }
            }
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
        private void TrudDeyatAdd_button_Click(object sender, EventArgs e)
        {
            dataGridView_TrudDeyat.Rows.Add(DateTime.Now.ToString("dd.MM.yyyy"), DateTime.Now.ToString("dd.MM.yyyy"), "1", "---", "---"); // добавить место работы
            this.SaveChangesToDataGridView_TrudDeyat(sender, e);
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В DataGridView_Married НА ВКЛАДКЕ "КАРТОЧКА 13-14" ##########
        private void SaveChangesToDataGridView_TrudDeyat(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder(); // создаем строку для построения
            foreach (DataGridViewRow row in dataGridView_TrudDeyat.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    // Обработка неверного формата даты. Возможно так и должно быть. Разобраться позже.
                    if ((cell.OwningColumn.Name == TrudDeyat_DataNaznach.Name) || (cell.OwningColumn.Name == TrudDeyat_DataOsvobozhd.Name))
                    {
                        DateTime TrudDeyat_wrongdatetoconvert = DateTime.Parse(cell.Value.ToString()); // парсим её в формат DateTime
                        cell.Value = TrudDeyat_wrongdatetoconvert.ToString("dd.MM.yyyy"); // и конвертируем в нужный формат
                    }
                    sb.Append(cell.Value); // добавляем значение ячейки
                    sb.Append("^"); // ставим разделитель ячейки 
                }
                sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель ячейки
                sb.Append("$"); // ставим разделитель строки
            }
            sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель строки

            dataGridView1[IndexTrudDeyat, IndexRowLichnayaKarta].Value = sb.ToString(); // Заполняем ячейку результирующей строкой
            this.AcceptAndWriteChanges(); // Применить изменения
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
            string StazhVyslugaStringFromCurrentCell = dataGridView1[IndexStazhVysluga, IndexRowLichnayaKarta].Value.ToString();
            if (StazhVyslugaStringFromCurrentCell != "") //проверка на существование данных в таблице
            {
                string[] stazhvysluga_array = StazhVyslugaStringFromCurrentCell.Split('$');

                foreach (string s in stazhvysluga_array)
                {
                    string[] StazhVyslugaRow = s.Split('^');
                    dataGridView_StazhVysluga.Rows.Add(StazhVyslugaRow);
                }
            }

            dataGridView_StazhVysluga.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            StazhVysluga_Resize();

            /*==============================================================================================================*/
            /*==============================================================================================================*/

            DataPrisyagi_dateTimePicker.Text = dataGridView1[IndexDataPrisyagi, IndexRowLichnayaKarta].Value.ToString();

            /*==============================================================================================================*/
            /*==============================================================================================================*/

            dataGridView_RabotaGFS.Rows.Clear();
            dataGridView_RabotaGFS.AutoGenerateColumns = false;
            string RabotaGFSStringFromCurrentCell = dataGridView1[IndexRabotaGFS, IndexRowLichnayaKarta].Value.ToString();
            if (RabotaGFSStringFromCurrentCell != "") //проверка на существование данных в таблице
            {
                string[] rabotagfs_array = RabotaGFSStringFromCurrentCell.Split('$');

                foreach (string s in rabotagfs_array)
                {
                    string[] RabotaGFSRow = s.Split('^');
                    dataGridView_RabotaGFS.Rows.Add(RabotaGFSRow);
                }
            }
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
                        //MessageBox.Show("Строка: " + (row.Index + 1).ToString() + " Должно быть: " + ResultHeight.ToString() + " ( " + dataGridView_StazhVysluga.Height.ToString() + " - " + dataGridView_StazhVysluga[0, 3].OwningRow.Height.ToString() + ") / " + (dataGridView_StazhVysluga.Rows.Count).ToString() + " а по факту: " + row.Height.ToString());
                    }
                    else row.Height = Stroka4; // Восстанавливаем высоту строки, присвоенную в самом начале свойством AutoSizeRowsMode
                }
        }

        // #################################################################
        // ##  КНОПКА "ДОБАВИТЬ МЕСТО СЛУЖБЫ" НА ВКЛАДКЕ "КАРТОЧКА 16-18" ##
        // #################################################################
        private void RabotaGFSAdd_button_Click(object sender, EventArgs e)
        {
            dataGridView_RabotaGFS.Rows.Add(DateTime.Now.ToString("dd.MM.yyyy"), DateTime.Now.ToString("dd.MM.yyyy"), "---", "---", "---", DateTime.Now.ToString("dd.MM.yyyy"), "1", "0"); // добавить место службы
            this.SaveChangesToDataGridView_RabotaGFS(sender, e);
        }


        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В DataGridView_StazhVysluga НА ВКЛАДКЕ "КАРТОЧКА 16-18" ##########
        private void SaveChangesToDataGridView_StazhVysluga(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder(); // создаем строку для построения
            foreach (DataGridViewRow row in dataGridView_StazhVysluga.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    sb.Append(cell.Value); // добавляем значение ячейки
                    sb.Append("^"); // ставим разделитель ячейки 
                }
                sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель
                sb.Append("$"); // ставим разделитель строки
            }
            sb.Remove(sb.Length - 1, 1); // Убираем последнюю точку с запятой

            dataGridView1[IndexStazhVysluga, IndexRowLichnayaKarta].Value = sb.ToString(); // Заполняем ячейку результирующей строкой
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

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В DataGridView_RabotaGFS НА ВКЛАДКЕ "КАРТОЧКА 16-18" ##########
        private void SaveChangesToDataGridView_RabotaGFS(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder(); // создаем строку для построения
            foreach (DataGridViewRow row in dataGridView_RabotaGFS.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    // Обработка неверного формата даты. Возможно так и должно быть. Разобраться позже.
                    if ((cell.OwningColumn.Name == RabotaGFS_DataNaznach.Name) || (cell.OwningColumn.Name == RabotaGFS_DataOsvobozhd.Name) || (cell.OwningColumn.Name == RabotaGFS_DataPrikaza.Name))
                    {
                        DateTime RabotaGFS_wrongdatetoconvert = DateTime.Parse(cell.Value.ToString()); // парсим её в формат DateTime
                        cell.Value = RabotaGFS_wrongdatetoconvert.ToString("dd.MM.yyyy"); // и конвертируем в нужный формат
                    }
                    sb.Append(cell.Value); // добавляем значение ячейки
                    sb.Append("^"); // ставим разделитель ячейки 
                }
                sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель
                sb.Append("$"); // ставим разделитель строки
            }
            sb.Remove(sb.Length - 1, 1); // Убираем последнюю точку с запятой

            dataGridView1[IndexRabotaGFS, IndexRowLichnayaKarta].Value = sb.ToString(); // Заполняем ячейку результирующей строкой
            this.AcceptAndWriteChanges(); // Применить изменения
        }



        //               //""""""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 19-20"  ############################################################
        public void UpdateCard19and20()
        {
            dataGridView_Attestaciya.Rows.Clear();
            dataGridView_Attestaciya.AutoGenerateColumns = false;
            string AttestaciyaStringFromCurrentCell = dataGridView1[IndexAttestaciya, IndexRowLichnayaKarta].Value.ToString();

            if (AttestaciyaStringFromCurrentCell != "") //проверка на существование данных в таблице
            {

                string[] attestaciya_array = AttestaciyaStringFromCurrentCell.Split('$');

                foreach (string s in attestaciya_array)
                {
                    string[] AttestaciyaRow = s.Split('^');
                    dataGridView_Attestaciya.Rows.Add(AttestaciyaRow);
                }
            }
            Attestaciya_Data.Width = 140;
            Attestaciya_Data.MinimumWidth = 140;
            Attestaciya_Prichina.Width = 180;
            Attestaciya_Prichina.MinimumWidth = 180;

            /*==============================================================================================================*/
            /*==============================================================================================================*/

            dataGridView_ProfPodg.Rows.Clear();
            dataGridView_ProfPodg.AutoGenerateColumns = false;
            string ProfPodgStringFromCurrentCell = dataGridView1[IndexProfPodg, IndexRowLichnayaKarta].Value.ToString();

            if (ProfPodgStringFromCurrentCell != "") //проверка на существование данных в таблице
            {

                string[] profpodg_array = ProfPodgStringFromCurrentCell.Split('$');

                foreach (string s in profpodg_array)
                {
                    string[] ProfPodgRow = s.Split('^');
                    dataGridView_ProfPodg.Rows.Add(ProfPodgRow);
                }
            }
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
        private void AttestaciyaAdd_button_Click(object sender, EventArgs e)
        {
            dataGridView_Attestaciya.Rows.Add(DateTime.Now.ToString("dd.MM.yyyy"), "Плановая", "Cоответствует замещаемой должности"); // добавить аттестацию
            this.SaveChangesToDataGridView_Attestaciya(sender, e);
        }

        // ###############################################################
        // ##  КНОПКА "ДОБАВИТЬ ПОДГОТОВКУ" НА ВКЛАДКЕ "КАРТОЧКА 19-20" ##
        // ###############################################################
        private void ProfPodgAdd_button_Click(object sender, EventArgs e)
        {
            dataGridView_ProfPodg.Rows.Add("Первоначальное обучение", DateTime.Now.ToString("dd.MM.yyyy"), DateTime.Now.ToString("dd.MM.yyyy"), "---", "---"); // добавить аттестацию
            this.SaveChangesToDataGridView_ProfPodg(sender, e);
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В DataGridView_Attestaciya НА ВКЛАДКЕ "КАРТОЧКА 19-20" ##########
        private void SaveChangesToDataGridView_Attestaciya(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder(); // создаем строку для построения
            foreach (DataGridViewRow row in dataGridView_Attestaciya.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    // Обработка неверного формата даты. Возможно так и должно быть. Разобраться позже.
                    if (cell.OwningColumn.Name == Attestaciya_Data.Name)
                    {
                        DateTime Attestaciya_wrongdatetoconvert = DateTime.Parse(cell.Value.ToString()); // парсим её в формат DateTime
                        cell.Value = Attestaciya_wrongdatetoconvert.ToString("dd.MM.yyyy"); // и конвертируем в нужный формат

                        if (cell.RowIndex + 1 == dataGridView_Attestaciya.Rows.Count)
                        {
                            // Заполняем ячейку "Дата следующей аттестации", прибавив 4 года к последней аттестации
                            dataGridView1[IndexNextAttestaciyaDate, IndexRowLichnayaKarta].Value = Attestaciya_wrongdatetoconvert.AddYears(4).ToString("dd.MM.yyyy");
                        }
                    }
                    sb.Append(cell.Value); // добавляем значение ячейки
                    sb.Append("^"); // ставим разделитель ячейки 
                }
                sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель
                sb.Append("$"); // ставим разделитель строки
            }
            sb.Remove(sb.Length - 1, 1); // Убираем последнюю точку с запятой

            dataGridView1[IndexAttestaciya, IndexRowLichnayaKarta].Value = sb.ToString(); // Заполняем ячейку "Аттестация" результирующей строкой
            this.AcceptAndWriteChanges(); // Применить изменения
        }
        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В DataGridView_ProfPodg НА ВКЛАДКЕ "КАРТОЧКА 19-20" ##########
        private void SaveChangesToDataGridView_ProfPodg(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder(); // создаем строку для построения
            foreach (DataGridViewRow row in dataGridView_ProfPodg.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    // Обработка неверного формата даты. Возможно так и должно быть. Разобраться позже.
                    if ((cell.OwningColumn.Name == ProfPodg_DataNach.Name) || (cell.OwningColumn.Name == ProfPodg_DataOkonch.Name))
                    {
                        DateTime ProfPodg_wrongdatetoconvert = DateTime.Parse(cell.Value.ToString()); // парсим её в формат DateTime
                        cell.Value = ProfPodg_wrongdatetoconvert.ToString("dd.MM.yyyy"); // и конвертируем в нужный формат
                    }
                    sb.Append(cell.Value); // добавляем значение ячейки
                    sb.Append("^"); // ставим разделитель ячейки 
                }
                sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель
                sb.Append("$"); // ставим разделитель строки
            }
            sb.Remove(sb.Length - 1, 1); // Убираем последнюю точку с запятой

            dataGridView1[IndexProfPodg, IndexRowLichnayaKarta].Value = sb.ToString(); // Заполняем ячейку результирующей строкой
            this.AcceptAndWriteChanges(); // Применить изменения
        }



        //               //""""""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 21-22"  ############################################################
        public void UpdateCard21and22()
        {
            Card21and22WasLoaded = 0;
            //Klassnost_label.Text = Klassnost_HeaderText + ":"; // Квалификационное звание
            Klassnost_comboBox.Text = dataGridView1[IndexKlassnost, IndexRowLichnayaKarta].Value.ToString();
            KlassnostCheyPrikaz_textBox.Text = dataGridView1[IndexKlassnostCheyPrikaz, IndexRowLichnayaKarta].Value.ToString();
            KlassnostNomerPrikaza_textBox.Text = dataGridView1[IndexKlassnostNomerPrikaza, IndexRowLichnayaKarta].Value.ToString();
            KlassnostDate_textBox.Text = dataGridView1[IndexKlassnostDate, IndexRowLichnayaKarta].Value.ToString();

            /*==============================================================================================================*/
            /*==============================================================================================================*/

            dataGridView_KlassnostOld.Rows.Clear();
            dataGridView_KlassnostOld.AutoGenerateColumns = false;
            string KlassnostOldStringFromCurrentCell = dataGridView1[IndexKlassnostOld, IndexRowLichnayaKarta].Value.ToString();

            if (KlassnostOldStringFromCurrentCell != "") //проверка на существование данных в таблице
            {
                string[] klassnostold_array = KlassnostOldStringFromCurrentCell.Split('$');

                foreach (string s in klassnostold_array)
                {
                    string[] KlassnostOldRow = s.Split('^');
                    dataGridView_KlassnostOld.Rows.Add(KlassnostOldRow);
                }
            }

            /*==============================================================================================================*/
            /*==============================================================================================================*/

            dataGridView_Nagrady.Rows.Clear();
            dataGridView_Nagrady.AutoGenerateColumns = false;
            string NagradyStringFromCurrentCell = dataGridView1[IndexNagrady, IndexRowLichnayaKarta].Value.ToString();

            if (NagradyStringFromCurrentCell != "") //проверка на существование данных в таблице
            {
                string[] nagrady_array = NagradyStringFromCurrentCell.Split('$');

                foreach (string s in nagrady_array)
                {
                    string[] NagradyRow = s.Split('^');
                    dataGridView_Nagrady.Rows.Add(NagradyRow);
                }
            }
            Card21and22WasLoaded = 1; // карточка прогрузилась
        }

        // ##########################################################################
        // ##  КНОПКА "ДОБАВИТЬ ПРЕДЫДУЩУЮ КЛАССНОСТЬ" НА ВКЛАДКЕ "КАРТОЧКА 21-22" ##
        // ##########################################################################
        private void KlassnostAdd_button_Click(object sender, EventArgs e)
        {
            dataGridView_KlassnostOld.Rows.Add("Специалист 3 класса", "---", "---", DateTime.Now.ToString("dd.MM.yyyy")); // добавить предыдущую классность
            this.SaveChangesToDataGridView_KlassnostOld(sender, e);
        }

        // ########################################################################
        // ##  КНОПКА "ДОБАВИТЬ НАГРАДЫ / ПООЩРЕНИЯ" НА ВКЛАДКЕ "КАРТОЧКА 21-22" ##
        // ########################################################################
        private void NagradyAdd_button_Click(object sender, EventArgs e)
        {
            dataGridView_Nagrady.Rows.Add("---", "---", "---", DateTime.Now.ToString("dd.MM.yyyy")); // добавить награды / поощрения
            this.SaveChangesToDataGridView_Nagrady(sender, e);
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

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В DataGridView_KlassnostOld НА ВКЛАДКЕ "КАРТОЧКА 21-22" ##########
        private void SaveChangesToDataGridView_KlassnostOld(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder(); // создаем строку для построения
            foreach (DataGridViewRow row in dataGridView_KlassnostOld.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    // Обработка неверного формата даты. Возможно так и должно быть. Разобраться позже.
                    if (cell.OwningColumn.Name == KlassnostDataPrikaza_dGV.Name)
                    {
                        DateTime KlassnostOld_wrongdatetoconvert = DateTime.Parse(cell.Value.ToString()); // парсим её в формат DateTime
                        cell.Value = KlassnostOld_wrongdatetoconvert.ToString("dd.MM.yyyy"); // и конвертируем в нужный формат
                    }
                    sb.Append(cell.Value); // добавляем значение ячейки
                    sb.Append("^"); // ставим разделитель ячейки 
                }
                sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель
                sb.Append("$"); // ставим разделитель строки
            }
            sb.Remove(sb.Length - 1, 1); // Убираем последнюю точку с запятой

            dataGridView1[IndexKlassnostOld, IndexRowLichnayaKarta].Value = sb.ToString(); // Заполняем ячейку результирующей строкой
            this.AcceptAndWriteChanges(); // Применить изменения
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В DataGridView_Nagrady НА ВКЛАДКЕ "КАРТОЧКА 21-22" ##########
        private void SaveChangesToDataGridView_Nagrady(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder(); // создаем строку для построения
            foreach (DataGridViewRow row in dataGridView_Nagrady.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    // Обработка неверного формата даты. Возможно так и должно быть. Разобраться позже.
                    if (cell.OwningColumn.Name == Nagrady_DataPrikaza.Name)
                    {
                        DateTime Nagrady_wrongdatetoconvert = DateTime.Parse(cell.Value.ToString()); // парсим её в формат DateTime
                        cell.Value = Nagrady_wrongdatetoconvert.ToString("dd.MM.yyyy"); // и конвертируем в нужный формат
                    }
                    sb.Append(cell.Value); // добавляем значение ячейки
                    sb.Append("^"); // ставим разделитель ячейки 
                }
                sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель
                sb.Append("$"); // ставим разделитель строки
            }
            sb.Remove(sb.Length - 1, 1); // Убираем последнюю точку с запятой

            dataGridView1[IndexNagrady, IndexRowLichnayaKarta].Value = sb.ToString(); // Заполняем ячейку результирующей строкой
            this.AcceptAndWriteChanges(); // Применить изменения
        }


        //               //""""""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 23-25"  ############################################################
        public void UpdateCard23to25()
        {
            Card23to25WasLoaded = 0;
            dataGridView_Prodlenie.Rows.Clear();
            dataGridView_Prodlenie.AutoGenerateColumns = false;

            string ProdlenieStringFromCurrentCell = dataGridView1[IndexProdlenie, IndexRowLichnayaKarta].Value.ToString();

            if (ProdlenieStringFromCurrentCell != "") //проверка на существование данных в таблице
            {
                Prodlenie_checkBox.CheckState = CheckState.Checked;
                string[] prodlenie_array = ProdlenieStringFromCurrentCell.Split('$');

                foreach (string s in prodlenie_array)
                {
                    string[] ProdlenieRow = s.Split('^');
                    dataGridView_Prodlenie.Rows.Add(ProdlenieRow);
                }
            }
            else
            {
                Prodlenie_checkBox.CheckState = CheckState.Unchecked;
            }
            Card23to25WasLoaded = 1;

            /*==============================================================================================================*/
            /*==============================================================================================================*/

            dataGridView_Boevye.Rows.Clear();
            dataGridView_Boevye.AutoGenerateColumns = false;
            string BoevyeStringFromCurrentCell = dataGridView1[IndexBoevye, IndexRowLichnayaKarta].Value.ToString();

            if (BoevyeStringFromCurrentCell != "") //проверка на существование данных в таблице
            {
                string[] boevye_array = BoevyeStringFromCurrentCell.Split('$');

                foreach (string s in boevye_array)
                {
                    string[] BoevyeRow = s.Split('^');
                    dataGridView_Boevye.Rows.Add(BoevyeRow);
                }
            }

            /*==============================================================================================================*/
            /*==============================================================================================================*/

            dataGridView_Rezerv.Rows.Clear();
            dataGridView_Rezerv.AutoGenerateColumns = false;
            string RezervStringFromCurrentCell = dataGridView1[IndexRezerv, IndexRowLichnayaKarta].Value.ToString();

            if (RezervStringFromCurrentCell != "") //проверка на существование данных в таблице
            {
                string[] rezerv_array = RezervStringFromCurrentCell.Split('$');

                foreach (string s in rezerv_array)
                {
                    string[] RezervRow = s.Split('^');
                    dataGridView_Rezerv.Rows.Add(RezervRow);
                }
            }
        }

        // ###############################################################################
        // ##  КНОПКА "ДОБАВИТЬ УЧАСТИЕ В БОЕВЫХ ДЕЙСТВИЯХ" НА ВКЛАДКЕ "КАРТОЧКА 23-25" ##
        // ###############################################################################
        private void BoevyeAdd_button_Click(object sender, EventArgs e)
        {
            dataGridView_Boevye.Rows.Add("---", DateTime.Now.ToString("dd.MM.yyyy"), DateTime.Now.ToString("dd.MM.yyyy"), "1", "---"); // добавить участие в боевых действиях
            this.SaveChangesToDataGridView_Boevye(sender, e);
        }

        // ########################################################################
        // ##  КНОПКА "ДОБАВИТЬ СОСТОЯНИЕ В РЕЗЕРВЕ" НА ВКЛАДКЕ "КАРТОЧКА 23-25" ##
        // ########################################################################
        private void RezervAdd_button_Click(object sender, EventArgs e)
        {
            dataGridView_Rezerv.Rows.Add("---", DateTime.Now.ToString("yyyy"), "---", DateTime.Now.ToString("dd.MM.yyyy")); // добавить состояние в резерве
            this.SaveChangesToDataGridView_Rezerv(sender, e);
        }


        // ##########  ИЗМЕНЕНИЕ СОСТОЯНИЯ Prodlenie_checkBox НА ВКЛАДКЕ "КАРТОЧКА 23-25" ##########
        private void Prodlenie_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            if (Card23to25WasLoaded == 1)
            {
                if (Prodlenie_checkBox.CheckState == CheckState.Checked)
                {
                    dataGridView_Prodlenie.Rows.Add(DateTime.Now.ToString("dd.MM.yyyy"), "1"); // добавить продление службы
                    this.SaveChangesToDataGridView_Prodlenie(sender, e);
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

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В DataGridView_Prodlenie НА ВКЛАДКЕ "КАРТОЧКА 23-25" ##########
        private void SaveChangesToDataGridView_Prodlenie(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder(); // создаем строку для построения
            foreach (DataGridViewRow row in dataGridView_Prodlenie.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    // Обработка неверного формата даты. Возможно так и должно быть. Разобраться позже.
                    if (cell.OwningColumn.Name == Prodlenie_Data.Name)
                    {
                        DateTime Prodlenie_wrongdatetoconvert = DateTime.Parse(cell.Value.ToString()); // парсим её в формат DateTime
                        cell.Value = Prodlenie_wrongdatetoconvert.ToString("dd.MM.yyyy"); // и конвертируем в нужный формат
                    }
                    sb.Append(cell.Value); // добавляем значение ячейки
                    sb.Append("^"); // ставим разделитель ячейки 
                }
                sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель
                sb.Append("$"); // ставим разделитель строки
            }
            sb.Remove(sb.Length - 1, 1); // Убираем последнюю точку с запятой

            dataGridView1[IndexProdlenie, IndexRowLichnayaKarta].Value = sb.ToString(); // Заполняем ячейку результирующей строкой
            this.AcceptAndWriteChanges(); // Применить изменения
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В DataGridView_Boevye НА ВКЛАДКЕ "КАРТОЧКА 23-25" ##########
        private void SaveChangesToDataGridView_Boevye(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder(); // создаем строку для построения
            foreach (DataGridViewRow row in dataGridView_Boevye.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    // Обработка неверного формата даты. Возможно так и должно быть. Разобраться позже.
                    if ((cell.OwningColumn.Name == Boevye_DataNach.Name) || (cell.OwningColumn.Name == Boevye_DataOkonch.Name))
                    {
                        DateTime Boevye_wrongdatetoconvert = DateTime.Parse(cell.Value.ToString()); // парсим её в формат DateTime
                        cell.Value = Boevye_wrongdatetoconvert.ToString("dd.MM.yyyy"); // и конвертируем в нужный формат
                    }
                    sb.Append(cell.Value); // добавляем значение ячейки
                    sb.Append("^"); // ставим разделитель ячейки 
                }
                sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель
                sb.Append("$"); // ставим разделитель строки
            }
            sb.Remove(sb.Length - 1, 1); // Убираем последнюю точку с запятой

            dataGridView1[IndexBoevye, IndexRowLichnayaKarta].Value = sb.ToString(); // Заполняем ячейку результирующей строкой
            this.AcceptAndWriteChanges(); // Применить изменения
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В DataGridView_Rezerv НА ВКЛАДКЕ "КАРТОЧКА 23-25" ##########
        private void SaveChangesToDataGridView_Rezerv(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder(); // создаем строку для построения
            foreach (DataGridViewRow row in dataGridView_Rezerv.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    // Обработка неверного формата даты. Возможно так и должно быть. Разобраться позже.
                    if (cell.OwningColumn.Name == Rezerv_DataPrikaza.Name)
                    {
                        DateTime Rezerv_wrongdatetoconvert = DateTime.Parse(cell.Value.ToString()); // парсим её в формат DateTime
                        cell.Value = Rezerv_wrongdatetoconvert.ToString("dd.MM.yyyy"); // и конвертируем в нужный формат
                    }
                    sb.Append(cell.Value); // добавляем значение ячейки
                    sb.Append("^"); // ставим разделитель ячейки 
                }
                sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель
                sb.Append("$"); // ставим разделитель строки
            }
            sb.Remove(sb.Length - 1, 1); // Убираем последнюю точку с запятой

            dataGridView1[IndexRezerv, IndexRowLichnayaKarta].Value = sb.ToString(); // Заполняем ячейку результирующей строкой
            this.AcceptAndWriteChanges(); // Применить изменения
        }


        //               //""""""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 26-29"  ############################################################
        public void UpdateCard26to29()
        {
            Card26to29WasLoaded = 0;
            if (CurrentDataTableName == "Kadry")
            {
                SendToAnotherDataTable_button.Text = "Переместить в архив";
            }
            else SendToAnotherDataTable_button.Text = "Восстановить из архива";

            dataGridView_Vzyskaniya.Rows.Clear();
            dataGridView_Vzyskaniya.AutoGenerateColumns = false;
            string VzyskaniyaStringFromCurrentCell = dataGridView1[IndexVzyskaniya, IndexRowLichnayaKarta].Value.ToString();

            if (VzyskaniyaStringFromCurrentCell != "") //проверка на существование данных в таблице
            {
                string[] vzyskaniya_array = VzyskaniyaStringFromCurrentCell.Split('$');

                foreach (string s in vzyskaniya_array)
                {
                    string[] VzyskaniyaRow = s.Split('^');
                    dataGridView_Vzyskaniya.Rows.Add(VzyskaniyaRow);
                }
            }

            /*==============================================================================================================*/
            /*==============================================================================================================*/

            dataGridView_Uvolnenie.Rows.Clear();
            dataGridView_Uvolnenie.AutoGenerateColumns = false;
            string UvolnenieStringFromCurrentCell = dataGridView1[IndexUvolnenie, IndexRowLichnayaKarta].Value.ToString();

            if (UvolnenieStringFromCurrentCell != "") //проверка на существование данных в таблице
            {
                string[] uvolnenie_array = UvolnenieStringFromCurrentCell.Split('$');

                foreach (string s in uvolnenie_array)
                {
                    string[] UvolnenieRow = s.Split('^');
                    dataGridView_Uvolnenie.Rows.Add(UvolnenieRow);
                }
            }

            /*==============================================================================================================*/
            /*==============================================================================================================*/

            Zapolnil_textBox.Text = dataGridView1[IndexZapolnil, IndexRowLichnayaKarta].Value.ToString();

            /*==============================================================================================================*/
            /*==============================================================================================================*/

            DataZapolneniya_dateTimePicker.Text = dataGridView1[IndexDataZapolneniya, IndexRowLichnayaKarta].Value.ToString();

            Card26to29WasLoaded = 1; // карточка прогрузилась
        }


        // ##############################################################
        // ##  КНОПКА "ДОБАВИТЬ ВЗЫСКАНИЕ" НА ВКЛАДКЕ "КАРТОЧКА 26-29" ##
        // ##############################################################
        private void VzyskaniyaAdd_button_Click(object sender, EventArgs e)
        {
            dataGridView_Vzyskaniya.Rows.Add("---", "---", "---", "---", DateTime.Now.ToString("dd.MM.yyyy"), "---", "---", DateTime.Now.ToString("dd.MM.yyyy")); // добавить взыскание
            this.SaveChangesToDataGridView_Vzyskaniya(sender, e);
        }

        // ###############################################################
        // ##  КНОПКА "ДОБАВИТЬ УВОЛЬНЕНИЕ" НА ВКЛАДКЕ "КАРТОЧКА 26-29" ##
        // ###############################################################
        private void UvolnenieAdd_button_Click(object sender, EventArgs e)
        {
            string UvolnenieProverka = dataGridView1[IndexUvolnenie, IndexRowLichnayaKarta].Value.ToString();
            if (UvolnenieProverka == "")
            {
                dataGridView_Uvolnenie.Rows.Add(DateTime.Now.ToString("dd.MM.yyyy"), "---", "---", DateTime.Now.ToString("dd.MM.yyyy"), "---"); // добавить увольнение
                this.SaveChangesToDataGridView_Uvolnenie(sender, e);
            }
            else
            {
                MessageBox.Show("Информация об увольнении уже существует!");
            }
        }


        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В DataGridView_Vzyskaniya НА ВКЛАДКЕ "КАРТОЧКА 26-29" ##########
        private void SaveChangesToDataGridView_Vzyskaniya(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder(); // создаем строку для построения
            foreach (DataGridViewRow row in dataGridView_Vzyskaniya.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    // Обработка неверного формата даты. Возможно так и должно быть. Разобраться позже.
                    if ((cell.OwningColumn.Name == Vzyskaniya_DataPrikazaNakaz.Name) || (cell.OwningColumn.Name == Vzyskaniya_DataPrikazaSnyatie.Name))
                    {
                        DateTime Vzyskaniya_wrongdatetoconvert = DateTime.Parse(cell.Value.ToString()); // парсим её в формат DateTime
                        cell.Value = Vzyskaniya_wrongdatetoconvert.ToString("dd.MM.yyyy"); // и конвертируем в нужный формат
                    }
                    sb.Append(cell.Value); // добавляем значение ячейки
                    sb.Append("^"); // ставим разделитель ячейки 
                }
                sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель
                sb.Append("$"); // ставим разделитель строки
            }
            sb.Remove(sb.Length - 1, 1); // Убираем последнюю точку с запятой

            dataGridView1[IndexVzyskaniya, IndexRowLichnayaKarta].Value = sb.ToString(); // Заполняем ячейку результирующей строкой
            this.AcceptAndWriteChanges(); // Применить изменения
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В DataGridView_Uvolnenie НА ВКЛАДКЕ "КАРТОЧКА 26-29" ##########
        private void SaveChangesToDataGridView_Uvolnenie(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder(); // создаем строку для построения
            foreach (DataGridViewRow row in dataGridView_Uvolnenie.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    // Обработка неверного формата даты. Возможно так и должно быть. Разобраться позже.                   
                    if ((cell.OwningColumn.Name == Uvolnenie_Data.Name) || (cell.OwningColumn.Name == Uvolnenie_DataPrikaza.Name))
                    {
                        DateTime Uvolnenie_wrongdatetoconvert = DateTime.Parse(cell.Value.ToString()); // парсим её в формат DateTime
                        cell.Value = Uvolnenie_wrongdatetoconvert.ToString("dd.MM.yyyy"); // и конвертируем в нужный формат
                    }
                    sb.Append(cell.Value); // добавляем значение ячейки
                    sb.Append("^"); // ставим разделитель ячейки 
                }
                sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель
                sb.Append("$"); // ставим разделитель строки
            }
            sb.Remove(sb.Length - 1, 1); // Убираем последнюю точку с запятой

            dataGridView1[IndexUvolnenie, IndexRowLichnayaKarta].Value = sb.ToString(); // Заполняем ячейку результирующей строкой
            this.AcceptAndWriteChanges(); // Применить изменения
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

        private void VyslugaCalc_button_Click(object sender, EventArgs e)
        {

        }


        // ####################################################
        // ##  КНОПКА "ЗАКРЫТЬ" НА ВКЛАДКЕ "ЛИЧНАЯ КАРТОЧКА" ##
        // ####################################################
        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        // #######################################################################################
        // ##  КНОПКА "ПЕРЕМЕСТИТЬ В АРХИВ/ВОССТАНОВИТЬ ИЗ АРХИВА" НА ВКЛАДКЕ "ЛИЧНАЯ КАРТОЧКА" ##
        // #######################################################################################
        private void SendToAnotherDataTable_button_Click(object sender, EventArgs e)
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
                    //MessageBox.Show(row["Surname"].ToString() + " в строке номер: " + IndexOfRowToExport);
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
                else Archive_button_Click(sender, e); // если грид пустой - переходим в другую базу
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
                    //add these 2
                    cb.DrawMode = DrawMode.OwnerDrawFixed;
                    cb.DropDownStyle = ComboBoxStyle.DropDownList;

                    cb.DrawItem += new DrawItemEventHandler(ComboBox_DrawItem_Centered);
                }
            }
        }

        // Allow Combo Box to center aligned
        public static void ComboBox_DrawItem_Centered(object sender, DrawItemEventArgs e)
        {
            // By using Sender, one method could handle multiple ComboBoxes
            ComboBox cbx = sender as ComboBox;
            if (cbx != null)
            {
                // Always draw the background
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
