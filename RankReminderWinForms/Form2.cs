using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace RankReminderWinForms
{

    public partial class Form2 : Form
    {
        //string filePath = @"C:\C#_Projects\Rank_Reminder\BaseLichSost.xml";

        public DataSet dsMain = new DataSet(); // создаем DataSet с именем dsMain
        string Required_PersonalFileNum, Required_PersonalNum, Required_Surname, Required_Name, Required_MiddleName, Required_DateOfBirth; // переменные для однозначного поиска сотрудника в Datatable

        readonly List<string> ZvanieList = new List<string>() //словарь "Звания"
            {
                "Рядовой", "Мл. сержант", "Сержант", "Ст. сержант", "Старшина",
                "Прапорщик", "Ст. прапорщик", "Мл. лейтенант", "Лейтенант",
                "Ст. лейтенант", "Капитан", "Майор", "Подполковник", "Полковник",
                "Генерал-майор", "Генерал-лейтенант", "Генерал-полковник", "Генерал"
            };

        readonly List<string> KlassnostList = new List<string>() //словарь "Классность"
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


        // Индексы колонок dgvMainList
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
        readonly DataGridViewTextBoxColumn cnum = new DataGridViewTextBoxColumn(); // Порядковый номер
        readonly DataGridViewTextBoxColumn personalfilenum = new DataGridViewTextBoxColumn(); // Номер личного дела
        readonly DataGridViewTextBoxColumn personalnum = new DataGridViewTextBoxColumn(); // Личный номер
        readonly DataGridViewTextBoxColumn surname = new DataGridViewTextBoxColumn(); // Фамилия
        readonly DataGridViewTextBoxColumn name = new DataGridViewTextBoxColumn(); // Имя
        readonly DataGridViewTextBoxColumn middleName = new DataGridViewTextBoxColumn(); // Отчество
        readonly DataGridViewTextBoxColumn gender = new DataGridViewTextBoxColumn(); // Пол
        readonly CalendarColumn dateofbirth = new CalendarColumn(); // Дата рождения
        readonly DataGridViewTextBoxColumn placeofbirth = new DataGridViewTextBoxColumn(); // Место рождения
        readonly DataGridViewTextBoxColumn registration = new DataGridViewTextBoxColumn(); // Прописан
        readonly DataGridViewTextBoxColumn placeofliving = new DataGridViewTextBoxColumn(); // Место жительства
        readonly DataGridViewTextBoxColumn phoneregistration = new DataGridViewTextBoxColumn(); // Телефон по прописке
        readonly DataGridViewTextBoxColumn phoneplaceofliving = new DataGridViewTextBoxColumn(); // Телефон по месту жительства
        readonly DataGridViewTextBoxColumn post = new DataGridViewTextBoxColumn(); // Должность
        readonly DataGridViewComboBoxColumn rank = new DataGridViewComboBoxColumn(); // Звание
        readonly CalendarColumn rankdate = new CalendarColumn(); // Дата присвоения звания
        readonly DataGridViewComboBoxColumn ranklimit = new DataGridViewComboBoxColumn(); // Потолок по званию
        readonly CalendarColumn nextrankdate = new CalendarColumn(); // Следующая дата присвоения звания
        readonly DataGridViewComboBoxColumn klassnost = new DataGridViewComboBoxColumn(); // Квалификационное звание (Классность)
        readonly CalendarColumn klassnostdate = new CalendarColumn(); // Дата присвоения квалиф. звания
        readonly CalendarColumn nextklassnostdate = new CalendarColumn(); // Следующая дата присвоения квалиф. звания
        readonly DataGridViewTextBoxColumn study = new DataGridViewTextBoxColumn(); // Образование
        readonly DataGridViewTextBoxColumn uchstepen = new DataGridViewTextBoxColumn(); // Ученая степень
        readonly DataGridViewTextBoxColumn prisvzvaniy = new DataGridViewTextBoxColumn(); // Дата присвоения званий и чинов        
        readonly DataGridViewTextBoxColumn married = new DataGridViewTextBoxColumn(); // Семейное положение
        readonly DataGridViewTextBoxColumn family = new DataGridViewTextBoxColumn(); // Члены семьи
        readonly DataGridViewTextBoxColumn truddeyat = new DataGridViewTextBoxColumn(); // Трудовая деятельность до прихода
        readonly DataGridViewTextBoxColumn stazhvysluga = new DataGridViewTextBoxColumn(); // Стаж и выслуга до прихода
        readonly DataGridViewTextBoxColumn dataprisyagi = new DataGridViewTextBoxColumn(); // Дата принятия присяги
        readonly DataGridViewTextBoxColumn rabotagfs = new DataGridViewTextBoxColumn(); // Прохождение службы (работа) в ГФС России
        readonly DataGridViewTextBoxColumn attestaciya = new DataGridViewTextBoxColumn(); // Аттестация
        readonly CalendarColumn nextattestaciyadate = new CalendarColumn(); // Дата следующей аттестации
        readonly DataGridViewTextBoxColumn profpodg = new DataGridViewTextBoxColumn(); // Профессиональная подготовка
        readonly DataGridViewTextBoxColumn klassnostcheyprikaz = new DataGridViewTextBoxColumn(); // Чей приказ о присвоении квалиф. звания
        readonly DataGridViewTextBoxColumn klassnostnomerprikaza = new DataGridViewTextBoxColumn(); // Номер приказа о присвоении квалиф. звания
        readonly DataGridViewTextBoxColumn klassnostold = new DataGridViewTextBoxColumn(); // Предыдущие квалификационные звания
        readonly DataGridViewTextBoxColumn nagrady = new DataGridViewTextBoxColumn(); // Награды и поощрения
        readonly DataGridViewTextBoxColumn prodlenie = new DataGridViewTextBoxColumn(); // Продление службы
        readonly DataGridViewTextBoxColumn boevye = new DataGridViewTextBoxColumn(); // Участие в боевых действиях
        readonly DataGridViewTextBoxColumn rezerv = new DataGridViewTextBoxColumn(); // Состояние в резерве
        readonly DataGridViewTextBoxColumn vzyskaniya = new DataGridViewTextBoxColumn(); // Взыскания
        readonly DataGridViewTextBoxColumn uvolnenie = new DataGridViewTextBoxColumn(); // Увольнение
        readonly DataGridViewTextBoxColumn zapolnil = new DataGridViewTextBoxColumn(); // Увольнение
        readonly DataGridViewTextBoxColumn datazapolneniya = new DataGridViewTextBoxColumn(); // Дата заполнения карточки     
        readonly DataGridViewTextBoxColumn imagestring = new DataGridViewTextBoxColumn(); // Изображение в виде текста


        //Наименования столбцов dgvMainList
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
            dgvMainList.AutoGenerateColumns = false;
            dgvMainList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            if (!File.Exists(XMLDB.Path)) // Если база данных в формате XML не существует...
            {
                MessageBox.Show("Похоже, что Вы запустили программу впервые, либо переместили файл базы данных. База будет создана заново.", "Внимание!",
                                 MessageBoxButtons.OK,
                                 MessageBoxIcon.Information);
                File.WriteAllBytes(XMLDB.Path, Convert.FromBase64String(XMLDB.DefaultXMLDBBase64)); //Декодируем строку с шаблоном базы данных из Base64 и создаем файл

                if (File.Exists(XMLDB.Path)) // Еще раз проверяем, создалась ли база данных
                {
                    MessageBox.Show($"База данных успешно создана по пути:\n{XMLDB.Path}", "Внимание!",
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

            dsMain.ReadXml(XMLDB.Path); // считываем XML базу данных в dsMain

            if (dsMain.Tables["Kadry"] == null) // Если DataTable "Kadry" отсутствует
            {
                dsMain.Tables.Add("Kadry");

                foreach (DataColumn column in dsMain.Tables["BackUp"].Columns) // Заполняем новый DataTable необходимыми колонками
                {
                    dsMain.Tables["Kadry"].Columns.Add(column.ColumnName);
                }
            }

            if (dsMain.Tables["Archive"] == null) // Если DataTable "Archive" отсутствует
            {
                dsMain.Tables.Add("Archive");

                foreach (DataColumn column in dsMain.Tables["BackUp"].Columns) // Заполняем новый DataTable необходимыми колонками
                {
                    dsMain.Tables["Archive"].Columns.Add(column.ColumnName);
                }
            }

            dgvMainList.DataSource = dsMain.Tables[CurrentDataTableName]; // присваиваем источник данных для dgvMainList

            DrawDatagrid(); // формируем DataGrid
            CheckColumnsIndex(); // сверяем индексы просчитываемых столбцов
            PereschetZvanie(); // пересчитываем звания
            PereschetKlassnost(); // пересчитываем классность
            PereschetCnum(); // пересчитываем порядковые номера 
        }

        // ###############  ДЕЙСТВИЯ ПОСЛЕ ЗАГРУЗКИ ФОРМЫ  ###############
        private void Form2_Load(object sender, EventArgs e)
        {
            radioButton1.Checked = true; // изначально показывать все колонки
            Cards_groupBox.Visible = false; // изначально не показывать кнопки карточек

            // ###############  ОБРАБОТЧИКИ СОБЫТИЙ УДАЛЕНИЯ СТРОК В ТАБЛИЦАХ  ###############
            dsMain.Tables["Kadry"].RowDeleting += new DataRowChangeEventHandler(RowDeleting); // обработчик события попытки удаления строки
            dsMain.Tables["Kadry"].RowDeleted += new DataRowChangeEventHandler(RowDeleted); // обработчик события удаления строки                                                                                         
            dsMain.Tables["Archive"].RowDeleting += new DataRowChangeEventHandler(RowDeleting); // обработчик события попытки удаления строки
            dsMain.Tables["Archive"].RowDeleted += new DataRowChangeEventHandler(RowDeleted); // обработчик события удаления строки     
                                                                                              //dgvFamily.UserDeletedRow += new System.Windows.Forms.DataGridViewRowEventHandler(WhichRowDeleted); // обработчик события удаления строки

            // ###############  РАЗНЫЕ ОБРАБОТЧИКИ СОБЫТИЙ  ###############
            dgvMainList.Sorted += new System.EventHandler(DataGridView1_Sorted); // обработчик события сортировки колонки
            tabControl1.Selecting += TabControl1_Selecting; // обработчик события перед сменой активной вкладки
            tabControl1.SelectedIndexChanged += TabControl1_SelectedIndexChanged; // обработчик события смены активной вкладки


            // ###############  ОБРАБОТЧИК ComboBox'ов ДЛЯ ЦЕНТРОВКИ В РЕЖИМЕ РЕДАКТИРОВАНИЯ  ###############
            cbxGender.DrawItem += new DrawItemEventHandler(CenteredComboBox.ComboBox_DrawItem_Centered); // Пол
            cbxKlassnost.DrawItem += new DrawItemEventHandler(CenteredComboBox.ComboBox_DrawItem_Centered); // Текущее квалификационное звание
            cbxRank.DrawItem += new DrawItemEventHandler(CenteredComboBox.ComboBox_DrawItem_Centered); // Текущее звание
            cbxRankLimit.DrawItem += new DrawItemEventHandler(CenteredComboBox.ComboBox_DrawItem_Centered); // Потолок по званию
                                                                                                            // Test_comboBox.DrawItem += new DrawItemEventHandler(OnDrawItem); // Потолок по званию

            // ###############  ОБРАБОТЧИК ComboBox'ов В dataGridView ДЛЯ ЦЕНТРОВКИ В РЕЖИМЕ РЕДАКТИРОВАНИЯ  ###############
            dgvProdlenie.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(CenteredComboBox.MyDGV_EditingControlShowing); // Продление службы
            dgvKlassnostOld.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(CenteredComboBox.MyDGV_EditingControlShowing); // Классность
            dgvAttestaciya.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(CenteredComboBox.MyDGV_EditingControlShowing); // Аттестация
            dgvFamily.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(CenteredComboBox.MyDGV_EditingControlShowing); // Члены семьи
            dgvMarried.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(CenteredComboBox.MyDGV_EditingControlShowing); // Семейное положение
            dgvProfPodg.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(CenteredComboBox.MyDGV_EditingControlShowing); // Профессиональная подготовка

            // ###############  ОБРАБОТЧИКИ СОБЫТИЙ ИЗМЕНЕНИЯ ТАБЛИЦ  ###############
            dgvMarried.CellValueChanged += new DataGridViewCellEventHandler(MarriedAdd_Click);
            dgvFamily.CellValueChanged += new DataGridViewCellEventHandler(FamilyAddPerson_Click);
            dgvStudy.CellValueChanged += new DataGridViewCellEventHandler(StudyAdd_Click);
            dgvUchStepen.CellValueChanged += new DataGridViewCellEventHandler(UchStepenAdd_Click);
            dgvPrisvZvaniy.CellValueChanged += new DataGridViewCellEventHandler(ZvanieAdd_Click);
            dgvTrudDeyat.CellValueChanged += new DataGridViewCellEventHandler(TrudDeyatAdd_Click);
            dgvStazhVysluga.CellValueChanged += new DataGridViewCellEventHandler(SaveChangesToDataGridView_StazhVysluga);
            dgvRabotaGFS.CellValueChanged += new DataGridViewCellEventHandler(RabotaGFSAdd_Click);
            dgvAttestaciya.CellValueChanged += new DataGridViewCellEventHandler(AttestaciyaAdd_Click);
            dgvProfPodg.CellValueChanged += new DataGridViewCellEventHandler(ProfPodgAdd_Click);
            dgvKlassnostOld.CellValueChanged += new DataGridViewCellEventHandler(KlassnostAdd_Click);
            dgvNagrady.CellValueChanged += new DataGridViewCellEventHandler(NagradyAdd_Click);
            dgvProdlenie.CellValueChanged += new DataGridViewCellEventHandler(Prodlenie_checkBox_CheckedChanged);
            dgvBoevye.CellValueChanged += new DataGridViewCellEventHandler(BoevyeAdd_Click);
            dgvRezerv.CellValueChanged += new DataGridViewCellEventHandler(RezervAdd_Click);
            dgvVzyskaniya.CellValueChanged += new DataGridViewCellEventHandler(VzyskaniyaAdd_Click);
            dgvUvolnenie.CellValueChanged += new DataGridViewCellEventHandler(UvolnenieAdd_Click);

            // ###############  ОБРАБОТЧИКИ СОБЫТИЙ ИЗМЕНЕНИЯ РАЗМЕРА ТАБЛИЦ  ###############
            ResizeBegin += FormResizeBegin;
            Resize += FormResize;
            ResizeEnd += FormResizeEnd;

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

            lblCnum.Text = $"{IndexRowLichnayaKarta + 1} из {dgvMainList.RowCount}"; // Порядковый номер личной карточки

            lblCurrentDate.Text = $"Сегодня: {DateTime.Today.ToShortDateString()}"; //ставим текущую дату внизу формы
        }

        // ###############  ОТРИСОВКА dgvMainList  ###############
        private void DrawDatagrid()
        {
            DataGridViewCellStyle style = dgvMainList.ColumnHeadersDefaultCellStyle;
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
            rank.DataPropertyName = "rankLimit";
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
            dgvMainList.Columns.AddRange(cnum, personalfilenum, personalnum, surname, name, middleName, gender, dateofbirth, placeofbirth, registration, placeofliving, phoneregistration, phoneplaceofliving, post, rank,
            rankdate, ranklimit, nextrankdate, klassnost, klassnostdate, nextklassnostdate, study, uchstepen, prisvzvaniy, married, family, truddeyat, stazhvysluga, dataprisyagi, rabotagfs, attestaciya, nextattestaciyadate, profpodg, klassnostcheyprikaz,
            klassnostnomerprikaza, klassnostold, nagrady, prodlenie, boevye, rezerv, vzyskaniya, uvolnenie, zapolnil, datazapolneniya, imagestring);
        }


        // ###############  ОСНОВНОЙ МЕТОД СЧИТЫВАНИЯ ИНДЕКСОВ У НЕОБХОДИМЫХ КОЛОНОК  ###############
        private void CheckColumnsIndex()
        {
            foreach (DataGridViewColumn currColumn in dgvMainList.Columns) //пробегаем по всем колонкам в dgvMainList
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
            if (dgvMainList.Rows.Count != 0) // Проверка dgvMainList на пустоту
            {
                for (int i = 0; i < dgvMainList.RowCount; i++) // Заполняем колонку с порядковыми номерами строк
                {
                    dgvMainList[IndexCnum, i].Value = i + 1; // Увеличиваем порядковый номер в каждой последующей строке на единицу
                }
            }
        }


        // ###############  ОСНОВНОЙ МЕТОД ПРОВЕРКИ И ПЕРЕСЧЁТА СРОКОВ ВЫСЛУГИ  ###############
        private void PereschetZvanie()
        {
            if (dgvMainList.Rows.Count != 0) // Проверка dgvMainList на пустоту
            {
                int RankVariable = 0;
                int CurrentRankToCompare = 0; //"вес" текущего звания
                int RankLimitToCompare = 0; //"вес" потолка по званию
                int NumberOfYears = 1; //срок выслуги до следующего звания
                int pYear2, pMonth2, pDay2; //переменные для парсинга текстовой даты в год, месяц, день
                string peremennaya2; //переменная для хранения даты из ячеек

                foreach (DataGridViewRow currRow in dgvMainList.Rows) // проходим по каждой строке в таблице
                {
                    string rankCurrent = currRow.Cells[IndexRank].Value.ToString(); // Переменная с текущим званием
                    string rankLimit = currRow.Cells[IndexRankLimit].Value.ToString(); // Переменная с "потолком" по званию


                    if (rankCurrent == rankLimit) // Чтобы лишний раз не гонять циклы, первым делом проверяем,
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
                                elemVariable = rankCurrent; // ...то работаем с текущим званием
                            }
                            else if (i == 1) // если второй проход цикла...
                            {
                                elemVariable = rankLimit; // ...то работаем с "потолком" по званию
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
                        else if ((CurrentRankToCompare < RankLimitToCompare) &&
                            (CurrentRankToCompare == 5 || CurrentRankToCompare == 7 || CurrentRankToCompare == 14 || CurrentRankToCompare == 15 || CurrentRankToCompare == 16 || CurrentRankToCompare == 17 || CurrentRankToCompare == 18))
                        {
                            currRow.Cells[IndexNextRankDate].Value = "не установлена"; // для перечисленных званий срок выслуги не установлен
                        }
                        else if (CurrentRankToCompare < RankLimitToCompare) // Если есть куда расти, то определяем количество лет до следующего звания
                        {
                            switch (CurrentRankToCompare)
                            {
                                // Звания со сроком выслуги в один год:
                                case 1:
                                case 2:
                                case 8:
                                    NumberOfYears = 1;
                                    break;
                                // Звания со сроком выслуги в два года:
                                case 3:
                                case 9:
                                    NumberOfYears = 2;
                                    break;
                                // Звания со сроком выслуги в три года:
                                case 4:
                                case 10:
                                case 11:
                                    NumberOfYears = 3;
                                    break;
                                // Звания со сроком выслуги в четыре года:
                                case 12:
                                    NumberOfYears = 4;
                                    break;
                                // Звания со сроком выслуги в пять лет:
                                case 6:
                                case 13:
                                    NumberOfYears = 5;
                                    break;
                            }

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
            if (dgvMainList.Rows.Count != 0) // Проверка dgvMainList на пустоту
            {
                int NumberOfYearsKlassnost = 3; // срок выслуги до следующего квалификационного звания
                int pYear3, pMonth3, pDay3; // переменные для парсинга текстовой даты в год, месяц, день
                string peremennaya3; // переменная для хранения даты из ячеек

                foreach (DataGridViewRow currRow in dgvMainList.Rows) // проходим по каждой строке в таблице
                {
                    string elemStr = currRow.Cells[IndexKlassnost].Value.ToString(); // Переменная с текущим значением классности

                    // Если классность отсутствует
                    if (elemStr == "Отсутствует")
                    {
                        currRow.Cells[IndexKlassnostDate].Value = "--.--.----"; // убираем дату присвоения классности
                        currRow.Cells[IndexNextKlassnostDate].Value = "--.--.----"; // убираем следующую дату присвоения классности
                    }

                    // Если сотрудник является специалистом 3, 2 или 1 класса
                    if ((elemStr == "Специалист 3 класса")
                        || (elemStr == "Специалист 2 класса")
                        || (elemStr == "Специалист 1 класса")) //(KlassnostCompare == 3)
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

        private void ShowAllRows() // показать все строки dgvMainList
        {
            dgvMainList.CurrentCell = null;
            foreach (DataGridViewRow currentRow in dgvMainList.Rows)
            {
                currentRow.Visible = true; // показываем строку
            }
        }

        private void RadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (SomeRowsWasHidden == 1) // Если есть скрытые строки
            {
                ShowAllRows(); // Показываем все строки
                PereschetCnum(); // Пересчитываем порядковые номера строк
                SomeRowsWasHidden = 0; // Переводим маркер в состояние "Нет скрытых строк"
            }
            ShowAllColumns(); // Показываем "стандартные" колонки dgvMainList

        }

        private void ShowAllColumns() // показать все "стандартные" колонки dgvMainList
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

        private void RadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (SomeRowsWasHidden == 1) // Если есть скрытые строки
            {
                ShowAllRows(); // Показываем все строки
                PereschetCnum(); // Пересчитываем порядковые номера строк
                SomeRowsWasHidden = 0; // Переводим маркер в состояние "Нет скрытых строк"
            }
            ShowVysluga(); // Показываем колонки dgvMainList, связанные со специальным званием
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

        private void RadioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (SomeRowsWasHidden == 1) // Если есть скрытые строки
            {
                ShowAllRows(); // Показываем все строки
                PereschetCnum(); // Пересчитываем порядковые номера строк
                SomeRowsWasHidden = 0; // Переводим маркер в состояние "Нет скрытых строк"
            }
            ShowKlassnost(); // Показываем колонки dgvMainList, связанные с классным званием
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

        private void RadioButton4_CheckedChanged(object sender, EventArgs e)
        {
            ShowAttestaciya();
        }

        private void ShowAttestaciya() // показать только тех, у кого аттестация в следующем году (проводится каждые 4 года)
        {
            if (radioButton4.Checked == true) //без этой проверки, при radioButton1.Checked = true, метод срабатывает второй раз (позже разобраться почему)
            {
                bool SomeRowsMustBeHidden = false; // Переменная, обозначающая, нужно ли будет скрывать строки
                int VisibleCnum = 0; // Переменная для пересчета порядковых номеров в видимых строках

                DateTime CheckedYear = DateTime.Now.AddYears(1); // прибавляем к текущей дате один год в формат DateTime

                // Пробегаем по строкам первый раз и при нахождении нужных - проставляем им новые порядковые номера
                foreach (DataGridViewRow currentRow in dgvMainList.Rows)
                {
                    if (dgvMainList[IndexNextAttestaciyaDate, currentRow.Index].Value.ToString() != "") // если значение даты аттестации в строке не пустое
                    {
                        DateTime DateFromCurrentRow = DateTime.Parse(dgvMainList[IndexNextAttestaciyaDate, currentRow.Index].Value.ToString()); // парсим значение даты аттестации в формат DateTime

                        if (DateFromCurrentRow.Year == CheckedYear.Year) // Если год из ячейки с датой следующей аттестации является следующим годом
                        {
                            SomeRowsMustBeHidden = true;
                            VisibleCnum++;
                            dgvMainList[IndexCnum, currentRow.Index].Value = VisibleCnum; //вызывает ошибку
                        }
                    }
                }

                if (SomeRowsMustBeHidden == true) // Если были найдены сотрудники, у которых аттестация в следующем году
                {
                    // Пробегаем по строкам второй раз и прячем все строки, кроме найденных 
                    foreach (DataGridViewRow currentRow in dgvMainList.Rows)
                    {
                        if (dgvMainList[IndexNextAttestaciyaDate, currentRow.Index].Value.ToString() != "") // если значение даты аттестации в строке не пустое
                        {
                            DateTime DateFromCurrentRow = DateTime.Parse(dgvMainList[IndexNextAttestaciyaDate, currentRow.Index].Value.ToString()); // парсим значение даты аттестации в формат DateTime

                            if (DateFromCurrentRow.Year != CheckedYear.Year) // Если год из ячейки с датой следующей аттестации не является следующим годом
                            {
                                dgvMainList.CurrentCell = null;
                                currentRow.Visible = false; // скрываем строку
                            }
                        }
                        else // Если значение даты аттестации пустое
                        {
                            dgvMainList.CurrentCell = null;
                            currentRow.Visible = false; // скрываем строку
                        }
                    }
                    // Скрываем лишние столбцы
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

                    SomeRowsWasHidden = 1; // Переводим маркер в состояние "Есть скрытые строки"
                }
                else // Если по итогу первого перебора строк ничего не найдено
                {
                    MessageBox.Show("Сотрудники, подпадающие под данный фильтр отсутствуют!");
                    radioButton1.Checked = true; // сбрасываем выбор фильтра
                    ShowAllRows(); // Показываем все строки
                }
            }
        }


        private void DataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Tab && dgvMainList.CurrentCell.ColumnIndex == 1)
            {
                e.Handled = true;
                DataGridViewCell cell = dgvMainList.Rows[0].Cells[0];
                dgvMainList.CurrentCell = cell;
                dgvMainList.BeginEdit(true);
            }
        }



        // #################################################
        // ##  КНОПКА "ЗАКРЫТЬ" НА ВКЛАДКЕ "ОБЩИЙ СПИСОК" ##
        // #################################################
        private void Close1_Click(object sender, EventArgs e)
        {
            Close();
        }

        // ####################################
        // ##  КНОПКА "ДОБАВИТЬ СОТРУДНИКА"  ##
        // ####################################
        private void AddPerson_Click(object sender, EventArgs e)
        {
            int id;

            if (dgvMainList.Rows.Count == 0) // Проверка dgvMainList на пустоту
            {
                id = 0;
            }
            else
            {
                PereschetCnum();
                id = Convert.ToInt32(dgvMainList[IndexCnum, dgvMainList.RowCount - 1].Value); //присваиваем переменной ID последний порядковый номер
            }

            dsMain.Tables[CurrentDataTableName].Rows.Add(id + 1, "не указан", "не указан", "Фамилия", "Имя", "Отчество", "М",
                DateTime.Now.ToString("dd.MM.yyyy")/* Дата рождения */,
                "Город", "Прописка", "Адрес проживания",
                "не указан"/* Телефон 1 */,
                "не указан"/* Телефон 2 */, "Должность",
                "Рядовой"/* Спец. звание */,
                DateTime.Now.ToString("dd.MM.yyyy")/* Дата присвоения звания */,
                "Рядовой"/* Потолок по званию */,
                "роста нет"/* Дата след. звания */,
                "Отсутствует"/* Классность */,
                "--.--.----"/* Дата классности */,
                "--.--.----"/* След. дата классности */,
                ""/* 10.Образование */,
                ""/* 11. Ученая степень */,
                ""/* 12.Присвоение званий, чинов */,
                ""/* 13.Семейное положение */,
                ""/* 14.Члены семьи */,
                ""/* 15.Труд. деят. до прихода */,
                "Общий трудовой стаж^0^0^0$Льготная выслуга^0^0^0$Стаж для государственных служащих^0^0^0$Половина периода обучения в высш. и сред. спец. учебных заведениях (для лиц начальствующего состава)^0^0^0$Календарная выслуга^0^0^0"/* 16.Стаж и выслуга до прихода */,
                DateTime.Now.ToString("dd.MM.yyyy")/* 17.Дата принятия присяги */,
                ""/* 18.Прохождение службы (работа) в ГФС России */,
                ""/* 19.Аттестация */,
                ""/* 20.Дата следующей аттестации */,
                ""/* 21.Профессиональная подготовка */,
                "---"/* 22.Чей приказ о присвоении квалиф. звания */, "---"/* 23.Дата приказа о присвоении квалиф. звания */, ""/* 24.Сведения о присвоенных ранее квалиф. званиях  */,
                ""/* 25.Награды и поощрения */, ""/* 26.Продление службы */, ""/* 27.Участие в боевых действиях */, ""/* 28.Состояние в резерве */,
                ""/* 29.Взыскания */, ""/* 30.Увольнение */, ""/* 31.Карточку заполнил */,
                DateTime.Now.ToString("dd.MM.yyyy")/* 32.Дата заполнения карточки */,
                ""/* 33.Фото */);
            AcceptAndWriteChanges();
            lblCnum.Text = $"{IndexRowLichnayaKarta + 1} из {dgvMainList.RowCount}"; // Порядковый номер личной карточки         
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
                lblCurrentBase.Text = "Текущая БД: Архивные сотрудники";
                MessageBox.Show("Перешли в архивную базу данных");
            }
            else if (CurrentDataTableName == "Archive")
            {
                Archive.Text = "Архивные сотрудники";
                CurrentDataTableName = "Kadry";
                OtherDataTableName = "Archive";
                lblCurrentBase.Text = "Текущая БД: Действующие сотрудники";
                MessageBox.Show("Перешли в текущую базу данных");
            }
            RefreshDgvMainList(); // обновляем DataGridView1

            if (dgvMainList.Rows.Count != 0) // Проверка dgvMainList на пустоту
            {
                IndexRowLichnayaKarta = 0;
                PereschetZvanie();
                PereschetKlassnost();
                PereschetCnum(); // пересчитываем порядковые номера 
                lblCnum.Text = $"{IndexRowLichnayaKarta + 1} из {dgvMainList.RowCount}"; // Порядковый номер личной карточки
                lblCardsFIO.Text = $"{dgvMainList[IndexSurname, IndexRowLichnayaKarta].Value} {dgvMainList[IndexName, IndexRowLichnayaKarta].Value} {dgvMainList[IndexMiddleName, IndexRowLichnayaKarta].Value}"; // Прописываем ФИО над стрелками в карточке
            }
            // При необходимости, добавить сюда события при пустом гриде
        }


        // ##################################
        // ##  КНОПКА "ВЫГРУЗИТЬ В EXCEL"  ##
        // ##################################
        private void ExportToExcel_Click(object sender, EventArgs e)
        {
            ExportDataGridToExcel();
        }

        // ###############  ВЫГРУЗКА dgvMainList В EXCEL ФАЙЛ  ###############
        public void ExportDataGridToExcel()
        {
            //Формируем новый список listVisibleColumns, состоящий только из видимых столбцов
            List<DataGridViewColumn> listVisibleColumns = new List<DataGridViewColumn>();
            foreach (DataGridViewColumn currentCol in dgvMainList.Columns)
            {
                if (currentCol.Visible)
                {
                    listVisibleColumns.Add(currentCol);
                }
            }

            //Формируем новый список listVisibleRows, состоящий только из видимых строк
            List<DataGridViewRow> listVisibleRows = new List<DataGridViewRow>();
            foreach (DataGridViewRow currentRow in dgvMainList.Rows)
            {
                if (currentRow.Visible)
                {
                    listVisibleRows.Add(currentRow);
                }
            }

            /*==============================================================================================================*/

            // Подготавливаем Excel для экспорта dgvMainList
            Excel.Application ExcelApp = new Excel.Application();
            ExcelApp.Application.Workbooks.Add(Type.Missing); //Создаем новую книгу
            ExcelApp.Columns.ColumnWidth = 15; // устанавливаем ширину столбцов
            ExcelApp.Cells.WrapText = "true"; // устанавливаем перенос по словам

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)ExcelApp.Worksheets.get_Item(1); //Создаем новый лист
            xlWorkSheet.Name = "Сведения о личном составе"; // именуем лист

            Excel.PageSetup pageSetup = xlWorkSheet.PageSetup; // Блок параметров листа
            pageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4; // размер А4
            pageSetup.Orientation = Excel.XlPageOrientation.xlLandscape; // ландшафтная ориентация
            pageSetup.Zoom = false;
            // Ужимаем всё при выводе на печать
            pageSetup.FitToPagesWide = 1;
            pageSetup.FitToPagesTall = 1;

            /*==============================================================================================================*/

            // Заполняем заголовки
            for (int i = 0; i < listVisibleColumns.Count; i++) // Проходим только по видимым столбцам
            {
                ExcelApp.Cells[1, i + 1] = listVisibleColumns[i].HeaderText; // Заполняем первую строку Excel заголовками видимых столбцов
            }

            // Украшаем заголовки
            Excel.Range rngZagolovki = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, listVisibleColumns.Count]); // диапазон заголовка в файле Excel
            rngZagolovki.Cells.Font.Bold = true; // жирный шрифт
            rngZagolovki.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // горизонтальное выравнивание по центру
            rngZagolovki.BorderAround(Type.Missing, Excel.XlBorderWeight.xlThick, Excel.XlColorIndex.xlColorIndexAutomatic); // увеличиваем толщину внешних границ
            rngZagolovki.Borders.Color = Color.Black; // черный цвет границ
            rngZagolovki.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter; // вертикальное выравнивание по центру

            /*==============================================================================================================*/

            //Заполняем лист Excel видимыми строками, столбцами и центрируем столбцы с датами
            for (int col = 0; col < listVisibleColumns.Count; col++)
            {
                if (listVisibleColumns[col] is CalendarColumn) // Центрируем столбцы с датами в Excel
                {
                    Excel.Range rngColWithDate = xlWorkSheet.get_Range(xlWorkSheet.Cells[2, col + 1], xlWorkSheet.Cells[listVisibleRows.Count + 1, col + 1]); // диапазон столбца, где обнаружена дата (без заголовка)
                    rngColWithDate.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // горизонтальное выравнивание по центру
                }
                for (int row = 0; row < listVisibleRows.Count; row++)
                {
                    ExcelApp.Cells[row + 2, col + 1] = dgvMainList.Rows[listVisibleRows[row].Index].Cells[listVisibleColumns[col].Index].Value.ToString(); // Наполняем лист Excel видимыми ячейками, начиная с первой строки после заголовка
                }
            }

            //Украшаем все, кроме заголовка
            Excel.Range rngAllCellsWithoutHeaders = xlWorkSheet.get_Range(xlWorkSheet.Cells[2, 1], xlWorkSheet.Cells[listVisibleRows.Count + 1, listVisibleColumns.Count]); // диапазон всех ячеек, кроме заголовка
            rngAllCellsWithoutHeaders.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter; // вертикальное выравнивание по центру
            rngAllCellsWithoutHeaders.Borders[Excel.XlBordersIndex.xlInsideVertical].Color = Color.LightGray; //внутренние вертикальные границы области с данными
            rngAllCellsWithoutHeaders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Color = Color.Black; //внутренние горизонтальные границы области с данными
            rngAllCellsWithoutHeaders.Borders[Excel.XlBordersIndex.xlEdgeRight].Color = Color.Black; //крайняя правая граница области с данными
            rngAllCellsWithoutHeaders.Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = Color.Black; //крайняя левая граница области с данными
            rngAllCellsWithoutHeaders.Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = Color.Black; //крайняя нижняя граница области с данными

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
        private void DataGridView1_Sorted(object sender, EventArgs e) //отработка события изменения сортировки
        {
            PereschetCnum();
        }

        // ###############  ДЕЙСТВИЯ, ЕСЛИ БЫЛИ КАКИЕ-ЛИБО ИЗМЕНЕНИЯ В dgvMainList  ###############
        public void DataGridWasChanged()
        {
            // MessageBox.Show("grid изменен");  // позже будет закомментировано 
            PereschetZvanie(); // пересчитываем звание
            PereschetKlassnost(); // пересчитываем классность
            AcceptAndWriteChanges(); // сохраняем изменения в XML
            RefreshDgvMainList(); // обновляем DataGridView1
        }

        // ###############  ОБНОВЛЕНИЕ dgvMainList  ###############
        public void RefreshDgvMainList()
        {
            dsMain.Clear(); // очищаем dsMain
            dgvMainList.DataSource = null; // очищаем DataSource
            dsMain.ReadXml(XMLDB.Path); // считываем XML
            dgvMainList.DataSource = dsMain.Tables[CurrentDataTableName]; // присваиваем DataSource
        }

        // ###############  ПРИМЕНИТЬ ВСЕ ИЗМЕНЕНИЯ И СОХРАНИТЬ XML  ###############
        public void AcceptAndWriteChanges()
        {
            // MessageBox.Show("Произошло сохранение базы данных"); // позже будет закомментировано 
            dsMain.AcceptChanges(); // применяем изменения в dsMain
            dsMain.WriteXml(XMLDB.Path); // сохраняем изменения в XML          
        }


        // ###############  НАЧАЛО РЕДАКТИРОВАНИЯ ЯЧЕЙКИ  ###############
        private void DataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            CellValueToCompare = dgvMainList.CurrentCell.Value.ToString(); // присваиваем переменной CellValueToCompare текущее значение ячейки до редактирования
            LastEditedCellRow = dgvMainList.CurrentCell.RowIndex;
            LastEditedCellCol = dgvMainList.CurrentCell.ColumnIndex;
        }

        // ###############  ЗАВЕРШЕНИЕ РЕДАКТИРОВАНИЯ ЯЧЕЙКИ  ###############
        private void DataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (dgvMainList.CurrentCell.Value.ToString() != CellValueToCompare) // сравниваем CellValueToCompare со значением ячейки после редактирования
            {
                DataGridWasChanged();
                dgvMainList.CurrentCell = dgvMainList[LastEditedCellCol, LastEditedCellRow];
            }
        }

        // ###############  ДЕЙСТВИЯ ПРИ СРАБАТЫВАНИИ СОБЫТИЯ RowDeleting (ПЕРЕД УДАЛЕНИЕМ СТРОКИ)  ###############
        private void RowDeleting(object sender, DataRowChangeEventArgs e)
        {
            if (tabControl1.SelectedTab.Text == "Общий список")
            {
                DialogResult result = MessageBox.Show("Удалить данную запись?", "Вы уверены?",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (result == DialogResult.No) // если была нажать кнопка "Нет"
                {
                    WantToDeleteRow = 0; // сбрасываем маркер удаления строки в ноль
                    dsMain.Tables[CurrentDataTableName].RejectChanges(); // отменяем изменения
                    RefreshDgvMainList(); // обновляем DataGridView1
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
                AcceptAndWriteChanges(); // сохраняем изменения
                WantToDeleteRow = 0; // сбрасываем маркер удаления строки в ноль

                if (dgvMainList.Rows.Count != 0) // Проверка dgvMainList на пустоту
                {
                    IndexRowLichnayaKarta = 0;
                    PereschetCnum(); // пересчитываем порядковые номера 
                    lblCnum.Text = $"{IndexRowLichnayaKarta + 1} из {dgvMainList.RowCount}"; // Порядковый номер личной карточки
                    lblCardsFIO.Text = $"{dgvMainList[IndexSurname, IndexRowLichnayaKarta].Value} {dgvMainList[IndexName, IndexRowLichnayaKarta].Value} {dgvMainList[IndexMiddleName, IndexRowLichnayaKarta].Value}"; // Прописываем ФИО над стрелками в карточке
                }
                // Позже, при необходимости, описать события для ситуации, когда таблица остается пустой

            }
        }

        // ###############  СОБЫТИЕ, ПРИ СМЕНЕ АКТИВНОЙ ВКЛАДКИ ###############
        private void TabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (dgvMainList.Rows.Count == 0) // Проверка dgvMainList на пустоту. Если грид пустой - не даем уйти с вкладки "Общий список"
            {
                e.Cancel = true;
                MessageBox.Show("Сначала добавьте хотя бы одного сотрудника!");
            }
        }
        // ###############  СОБЫТИЕ, ПОСЛЕ СМЕНЫ АКТИВНОЙ ВКЛАДКИ ###############
        private void TabControl1_SelectedIndexChanged(Object sender, EventArgs e)
        {
            if (IndexRowLichnayaKarta > dgvMainList.RowCount - 1) // Проверка на выход за пределы диапазона личных карточек.
                                                                  // Такое может произойти, если была активной последняя карточка,
                                                                  // после чего её удалили и снова "вышли" из "Общего списка"
            {
                IndexRowLichnayaKarta = 0;
                lblCnum.Text = $"{IndexRowLichnayaKarta + 1} из {dgvMainList.RowCount}"; // Порядковый номер личной карточки
                lblCardsFIO.Text = $"{dgvMainList[IndexSurname, IndexRowLichnayaKarta].Value} {dgvMainList[IndexName, IndexRowLichnayaKarta].Value} {dgvMainList[IndexMiddleName, IndexRowLichnayaKarta].Value}"; // Прописываем ФИО над стрелками в карточке
            }
            NeedToUpdateCard();
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
                    UpdateCard1to9();
                    break;
                case "Карточка 10-11": // 
                    UpdateCard10and11();
                    break;
                case "Карточка 12": //
                    UpdateCard12();
                    break;
                case "Карточка 13-14": // 
                    UpdateCard13and14();
                    break;
                case "Карточка 15": // 
                    UpdateCard15();
                    break;
                case "Карточка 16-18": // 
                    UpdateCard16to18();
                    break;
                case "Карточка 19-20": // 
                    UpdateCard19and20();
                    break;
                case "Карточка 21-22": // 
                    UpdateCard21and22();
                    break;
                case "Карточка 23-25": // 
                    UpdateCard23to25();
                    break;
                case "Карточка 26-29": // 
                    UpdateCard26to29();
                    break;
            }
        }


        //               //""""""""""""""""""""""""\\ 
        // ###############  ВКЛАДКА "КАРТОЧКА 1-9"  ############################################################
        public void UpdateCard1to9()
        {
            Card1to9WasLoaded = 0;

            if (dgvMainList[IndexImageString, IndexRowLichnayaKarta].Value.ToString() == "") // Если картинка отсутствует
            {
                dgvMainList[IndexImageString, IndexRowLichnayaKarta].Value = XMLDB.DefaultImageBase64; // присваиваем pbMainPhoto стандартную картинку с крестиком
                Bitmap bmp = new Bitmap(new MemoryStream(Convert.FromBase64String(dgvMainList[IndexImageString, IndexRowLichnayaKarta].Value.ToString()))); // собираем изображение
                pbMainPhoto.Image = bmp;
            }
            else
            {
                Bitmap bmp = new Bitmap(new MemoryStream(Convert.FromBase64String(dgvMainList[IndexImageString, IndexRowLichnayaKarta].Value.ToString()))); // собираем изображение
                pbMainPhoto.Image = bmp; //присваиваем pbMainPhoto собранную ячейку
            }

            // ЗАПОЛНЯЕМ textBox'ы:
            tbxPersonalFileNum.Text = dgvMainList[IndexPersonalFileNum, IndexRowLichnayaKarta].Value.ToString();
            tbxPersonalNum.Text = dgvMainList[IndexPersonalNum, IndexRowLichnayaKarta].Value.ToString();
            tbxSurname.Text = dgvMainList[IndexSurname, IndexRowLichnayaKarta].Value.ToString();
            tbxName.Text = dgvMainList[IndexName, IndexRowLichnayaKarta].Value.ToString();
            tbxMiddleName.Text = dgvMainList[IndexMiddleName, IndexRowLichnayaKarta].Value.ToString();
            cbxGender.Text = dgvMainList[IndexGender, IndexRowLichnayaKarta].Value.ToString();
            dtpDateOfBirth.Text = dgvMainList[IndexDateOfBirth, IndexRowLichnayaKarta].Value.ToString();
            dtpRankDate.Text = dgvMainList[IndexRankDate, IndexRowLichnayaKarta].Value.ToString();
            tbxPlaceOfBirth.Text = dgvMainList[IndexPlaceOfBirth, IndexRowLichnayaKarta].Value.ToString();
            tbxRegistration.Text = dgvMainList[IndexRegistration, IndexRowLichnayaKarta].Value.ToString();
            tbxPlaceOfLiving.Text = dgvMainList[IndexPlaceOfLiving, IndexRowLichnayaKarta].Value.ToString();
            tbxPhoneRegistration.Text = dgvMainList[IndexPhoneRegistration, IndexRowLichnayaKarta].Value.ToString();
            tbxPhonePlaceOfLiving.Text = dgvMainList[IndexPhonePlaceOfLiving, IndexRowLichnayaKarta].Value.ToString();
            tbxPost.Text = dgvMainList[IndexPost, IndexRowLichnayaKarta].Value.ToString();
            tbxNextRankDate.Text = dgvMainList[IndexNextRankDate, IndexRowLichnayaKarta].Value.ToString();

            cbxRank.BindingContext = new BindingContext();   //создаем новый контекст, иначе в определенный момент получаем null в одном из comboBox'ов
            cbxRank.DataSource = ZvanieList;
            cbxRank.Text = dgvMainList[IndexRank, IndexRowLichnayaKarta].Value.ToString();
            cbxRankLimit.BindingContext = new BindingContext();   //создаем новый контекст, иначе в определенный момент получаем null в одном из comboBox'ов
            cbxRankLimit.DataSource = ZvanieList;
            cbxRankLimit.Text = dgvMainList[IndexRankLimit, IndexRowLichnayaKarta].Value.ToString();

            Card1to9WasLoaded = 1; // карточка прогрузилась
        }


        // ##################################################################################
        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В TextBox'ах НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        // ##################################################################################

        // Общий метод для проверки наличия изменений, после выхода из textbox'ов
        private void Tbx_CheckForChanges(TextBox tbxName, int indexOfTBX)
        {
            if (tbxName.Text != dgvMainList[indexOfTBX, IndexRowLichnayaKarta].Value.ToString())
            {
                dgvMainList[indexOfTBX, IndexRowLichnayaKarta].Value = tbxName.Text;
                AcceptAndWriteChanges(); // применить изменения
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В tbxSurname НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void TbxSurname_Leave(object sender, EventArgs e)
        {
            Tbx_CheckForChanges(tbxSurname, IndexSurname);
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В tbxName НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void TbxName_Leave(object sender, EventArgs e)
        {
            Tbx_CheckForChanges(tbxName, IndexName);
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В tbxMiddleName НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void TbxMiddleName_Leave(object sender, EventArgs e)
        {
            Tbx_CheckForChanges(tbxMiddleName, IndexMiddleName);
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В tbxPlaceOfBirth НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void TbxPlaceOfBirth_Leave(object sender, EventArgs e)
        {
            Tbx_CheckForChanges(tbxPlaceOfBirth, IndexPlaceOfBirth);
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В tbxPersonalFileNum НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void TbxPersonalFileNum_Leave(object sender, EventArgs e)
        {
            Tbx_CheckForChanges(tbxPersonalFileNum, IndexPersonalFileNum);
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В tbxPersonalNum НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void TbxPersonalNum_Leave(object sender, EventArgs e)
        {
            Tbx_CheckForChanges(tbxPersonalNum, IndexPersonalNum);
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В tbxRegistration НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void TbxRegistration_Leave(object sender, EventArgs e)
        {
            Tbx_CheckForChanges(tbxRegistration, IndexRegistration);
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В tbxPlaceOfLiving НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void TbxPlaceOfLiving_Leave(object sender, EventArgs e)
        {
            Tbx_CheckForChanges(tbxPlaceOfLiving, IndexPlaceOfLiving);
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В tbxPhoneRegistration НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void TbxPhoneRegistration_Leave(object sender, EventArgs e)
        {
            Tbx_CheckForChanges(tbxPhoneRegistration, IndexPhoneRegistration);
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В tbxPhonePlaceOfLiving НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void TbxPhonePlaceOfLiving_Leave(object sender, EventArgs e)
        {
            Tbx_CheckForChanges(tbxPhonePlaceOfLiving, IndexPhonePlaceOfLiving);
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В tbxPost НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void TbxPost_Leave(object sender, EventArgs e)
        {
            Tbx_CheckForChanges(tbxPost, IndexPost);
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В dtpDateOfBirth НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void DtpDateOfBirth_ValueChanged(object sender, EventArgs e) // dateTimePicker "Дата рождения"
        {
            if (Card1to9WasLoaded == 1)
            {
                dgvMainList[IndexDateOfBirth, IndexRowLichnayaKarta].Value = dtpDateOfBirth.Value.ToString("dd.MM.yyyy");
                AcceptAndWriteChanges(); // применить изменения
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В dtpRankDate НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void DtpRankDate_ValueChanged(object sender, EventArgs e)
        {
            if (Card1to9WasLoaded == 1)
            {
                dgvMainList[IndexRankDate, IndexRowLichnayaKarta].Value = dtpRankDate.Value.ToString("dd.MM.yyyy");
                PereschetZvanie();
                AcceptAndWriteChanges(); // применить изменения
                tbxNextRankDate.Text = dgvMainList[IndexNextRankDate, IndexRowLichnayaKarta].Value.ToString();
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В cbxRank НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void CbxRank_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Card1to9WasLoaded == 1)
            {
                dgvMainList[IndexRank, IndexRowLichnayaKarta].Value = cbxRank.Text;
                PereschetZvanie();
                AcceptAndWriteChanges(); // применить изменения
                tbxNextRankDate.Text = dgvMainList[IndexNextRankDate, IndexRowLichnayaKarta].Value.ToString();
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В cbxRankLimit НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void CbxRankLimit_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Card1to9WasLoaded == 1)
            {
                dgvMainList[IndexRankLimit, IndexRowLichnayaKarta].Value = cbxRankLimit.Text;
                PereschetZvanie();
                AcceptAndWriteChanges(); // применить изменения
                tbxNextRankDate.Text = dgvMainList[IndexNextRankDate, IndexRowLichnayaKarta].Value.ToString();
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В cbxGender НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##########
        private void CbxGender_SelectedIndexChanged(object sender, EventArgs e) // ComboBox "Пол"
        {
            if (Card1to9WasLoaded == 1)
            {
                dgvMainList[IndexGender, IndexRowLichnayaKarta].Value = cbxGender.Text;
                AcceptAndWriteChanges(); // применить изменения
            }
        }


        // ######################################################
        // ##  КНОПКА "ВЫБРАТЬ ФОТО" НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##
        // ######################################################
        private void ChooseImage_Click(object sender, EventArgs e)
        {
            OpenFileDialog photoSelector = new OpenFileDialog
            {
                Title = "Выберите новую фотографию сотрудника",
                InitialDirectory = "c:\\",
                Filter = "Все изображения|*.bmp; *.jpg; *.jpeg; *.png; *.gif"
            };
            if (photoSelector.ShowDialog() == DialogResult.OK) // если пользователь выбрал файл изображения
            {
                Bitmap bmp = new Bitmap(photoSelector.FileName); // присваиваем переменной bmp выбранный файл
                TypeConverter converter = TypeDescriptor.GetConverter(typeof(Bitmap));
                string ImageBase64 = Convert.ToBase64String((byte[])converter.ConvertTo(bmp, typeof(byte[]))); // конвертируем изображение в текст
                dgvMainList[IndexImageString, IndexRowLichnayaKarta].Value = ImageBase64; // записываем результат в соответствующую ячейку
                pbMainPhoto.Image = bmp; //присваиваем pbMainPhoto собранную ячейку
                AcceptAndWriteChanges(); // применить изменения
            }
        }


        // ######################################################
        // ##  КНОПКА "УДАЛИТЬ ФОТО" НА ВКЛАДКЕ "КАРТОЧКА 1-9" ##
        // ######################################################
        private void RemoveImage_Click(object sender, EventArgs e)
        {
            dgvMainList[IndexImageString, IndexRowLichnayaKarta].Value = XMLDB.DefaultImageBase64;
            Bitmap bmp = new Bitmap(new MemoryStream(Convert.FromBase64String(dgvMainList[IndexImageString, IndexRowLichnayaKarta].Value.ToString()))); // собираем изображение
            pbMainPhoto.Image = bmp;
            AcceptAndWriteChanges(); // применить изменения
        }



        //               //""""""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 10-11"  ############################################################
        public void UpdateCard10and11()
        {
            dgvStudy.Rows.Clear();
            dgvStudy.AutoGenerateColumns = false;

            Dgv_Draw(IndexStudy, dgvStudy); // Отрисовываем таблицу dgvStudy

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

            dgvUchStepen.Rows.Clear();
            dgvUchStepen.AutoGenerateColumns = false;

            Dgv_Draw(IndexUchStepen, dgvUchStepen); // Отрисовываем таблицу dgvUchStepen
        }


        // ###################################################################
        // ##  КНОПКА "ДОБАВИТЬ УЧЕНУЮ СТЕПЕНЬ" НА ВКЛАДКЕ "КАРТОЧКА 10-11" ##
        // ###################################################################
        private void UchStepenAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                Dgv_SaveChanges(IndexUchStepen, dgvUchStepen);
            }
            else //Если метод вызван нажатием кнопки
            {
                dgvUchStepen.Rows.Add("---", DateTime.Now.ToString("dd.MM.yyyy")); // добавить ученую степень
                Dgv_SaveChanges(IndexUchStepen, dgvUchStepen);
            }
        }


        // ################################################################
        // ##  КНОПКА "ДОБАВИТЬ ОБРАЗОВАНИЕ" НА ВКЛАДКЕ "КАРТОЧКА 10-11" ##
        // ################################################################
        private void StudyAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                Dgv_SaveChanges(IndexStudy, dgvStudy);
            }
            else //Если метод вызван нажатием кнопки
            {
                dgvStudy.Rows.Add("Высшее (очное)", "---", DateTime.Now.ToString("dd.MM.yyyy"), DateTime.Now.ToString("dd.MM.yyyy"), "---", "---", "---"); // добавить образование
                Dgv_SaveChanges(IndexStudy, dgvStudy);
            }
        }



        //               //"""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 12"  ############################################################
        public void UpdateCard12()
        {
            dgvPrisvZvaniy.Rows.Clear();
            dgvPrisvZvaniy.AutoGenerateColumns = false;

            Dgv_Draw(IndexPrisvZvaniy, dgvPrisvZvaniy); // Отрисовываем таблицу dgvPrisvZvaniy
        }


        // ######################################################################
        // ##  КНОПКА "ДОБАВИТЬ ЗВАНИЕ, КЛАССНЫЙ ЧИН" НА ВКЛАДКЕ "КАРТОЧКА 12" ##
        // ######################################################################
        private void ZvanieAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                Dgv_SaveChanges(IndexPrisvZvaniy, dgvPrisvZvaniy);
            }
            else //Если метод вызван нажатием кнопки
            {
                dgvPrisvZvaniy.Rows.Add("---", DateTime.Now.ToString("dd.MM.yyyy"), "---", "---", DateTime.Now.ToString("dd.MM.yyyy")); // добавить звание, классный чин
                Dgv_SaveChanges(IndexPrisvZvaniy, dgvPrisvZvaniy);
            }
        }



        //               //""""""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 13-14"  ############################################################
        public void UpdateCard13and14()
        {
            dgvMarried.Rows.Clear();
            dgvMarried.AutoGenerateColumns = false;

            Dgv_Draw(IndexMarried, dgvMarried); // Отрисовываем таблицу dgvMarried

            /*==============================================================================================================*/

            dgvFamily.Rows.Clear();
            dgvFamily.AutoGenerateColumns = false;

            Dgv_Draw(IndexFamily, dgvFamily); // Отрисовываем таблицу dgvFamily

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
                Dgv_SaveChanges(IndexMarried, dgvMarried);
            }
            else //Если метод вызван нажатием кнопки
            {
                dgvMarried.Rows.Add("Женат", DateTime.Now.ToString("yyyy")); // добавить событие (свадьба, развод)
                Dgv_SaveChanges(IndexMarried, dgvMarried);
            }
        }


        // ################################################################
        // ##  КНОПКА "ДОБАВИТЬ ЧЛЕНА СЕМЬИ" НА ВКЛАДКЕ "КАРТОЧКА 13-14" ##
        // ################################################################
        private void FamilyAddPerson_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                Dgv_SaveChanges(IndexFamily, dgvFamily);
            }
            else //Если метод вызван нажатием кнопки
            {
                dgvFamily.Rows.Add("Мать", DateTime.Now.ToString("dd.MM.yyyy"), "---"); // добавить члена семьи
                Dgv_SaveChanges(IndexFamily, dgvFamily);
            }
        }



        //               //"""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 15"  ############################################################
        public void UpdateCard15()
        {
            dgvTrudDeyat.Rows.Clear();
            dgvTrudDeyat.AutoGenerateColumns = false;

            Dgv_Draw(IndexTrudDeyat, dgvTrudDeyat); // Отрисовываем таблицу dgvTrudDeyat

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
                Dgv_SaveChanges(IndexTrudDeyat, dgvTrudDeyat);
            }
            else //Если метод вызван нажатием кнопки
            {
                dgvTrudDeyat.Rows.Add(DateTime.Now.ToString("dd.MM.yyyy"), DateTime.Now.ToString("dd.MM.yyyy"), "1", "---", "У"); // добавить место работы
                Dgv_SaveChanges(IndexTrudDeyat, dgvTrudDeyat);
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

            dgvStazhVysluga.Rows.Clear();
            dgvStazhVysluga.AutoGenerateColumns = false;

            Dgv_Draw(IndexStazhVysluga, dgvStazhVysluga); // Отрисовываем таблицу dgvStazhVysluga

            dgvStazhVysluga.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            StazhVysluga_Resize();

            /*==============================================================================================================*/

            dtpDataPrisyagi.Text = dgvMainList[IndexDataPrisyagi, IndexRowLichnayaKarta].Value.ToString();

            /*==============================================================================================================*/

            dgvRabotaGFS.Rows.Clear();
            dgvRabotaGFS.AutoGenerateColumns = false;

            Dgv_Draw(IndexRabotaGFS, dgvRabotaGFS); // Отрисовываем таблицу dgvRabotaGFS

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
            dgvStazhVysluga.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells; // Включаем свойство AutoSizeRowsMode, чтобы оно автоматически подстроило высоту строк в таблице
            StazhVysluga_Poyasnenie.DefaultCellStyle.WrapMode = DataGridViewTriState.True; // Перенос слов в колонке "Пояснение" 
            StazhVysluga_Poyasnenie.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            StazhVysluga_Let.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // Выравнивание в колонке "Лет"
            StazhVysluga_Mesyacev.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // Выравнивание в колонке "Месяцев"
            StazhVysluga_Dney.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // Выравнивание в колонке "Дней"

            int Stroka4 = dgvStazhVysluga[0, 3].OwningRow.Height; // Записываем в переменную высоту строки, присвоенную AutoSizeRowsMode
            int Stroka_proverka = dgvStazhVysluga[0, 2].OwningRow.Height; // Записываем в переменную высоту "стандартной" строки для дальнейшего сравнения
            int Zagolovok = dgvStazhVysluga.ColumnHeadersHeight; // Записываем в переменную высоту заголовка, присвоенную AutoSizeRowsMode
            dgvStazhVysluga.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None; // Отключаем свойство AutoSizeRowsMode, чтобы далее можно было программно присвоить высоту строк в таблице
            dgvStazhVysluga.ColumnHeadersHeight = Zagolovok;

            if (Stroka4 == Stroka_proverka) // Если размер окна программы поволяет уместить текст в четвертой строке без переносов, значит ставим всем одну высоту  
                foreach (DataGridViewRow row in dgvStazhVysluga.Rows)
                {
                    row.Height = (dgvStazhVysluga.Height - Zagolovok) / (dgvStazhVysluga.Rows.Count);// Вычисляем высоту строк для заполнения всего свободного пространства
                }

            else // Если размер окна программы НЕ поволяет уместить текст в четвертой строке без переносов, значит высота этой строки должна быть больше, чем у других
                foreach (DataGridViewRow row in dgvStazhVysluga.Rows)
                {
                    if (row != dgvStazhVysluga[0, 3].OwningRow) // Для всех строк, кроме четвертой
                    {
                        row.Height = (dgvStazhVysluga.Height - Stroka4 - Zagolovok) / (dgvStazhVysluga.Rows.Count - 1); // Вычисляем высоту строк для заполнения всего свободного пространства
                    }
                    else row.Height = Stroka4; // Восстанавливаем высоту строки, присвоенную в самом начале свойством AutoSizeRowsMode
                }
        }


        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В DataGridView_StazhVysluga НА ВКЛАДКЕ "КАРТОЧКА 16-18" ##########
        private void SaveChangesToDataGridView_StazhVysluga(object sender, EventArgs e)
        {
            MessageBox.Show("Stazh");
            Dgv_SaveChanges(IndexStazhVysluga, dgvStazhVysluga);
        }


        // #################################################################
        // ##  КНОПКА "ДОБАВИТЬ МЕСТО СЛУЖБЫ" НА ВКЛАДКЕ "КАРТОЧКА 16-18" ##
        // #################################################################
        private void RabotaGFSAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                Dgv_SaveChanges(IndexRabotaGFS, dgvRabotaGFS);
            }
            else //Если метод вызван нажатием кнопки
            {
                dgvRabotaGFS.Rows.Add(DateTime.Now.ToString("dd.MM.yyyy"), DateTime.Now.ToString("dd.MM.yyyy"), "---", "---", "---", DateTime.Now.ToString("dd.MM.yyyy"), "1", "0"); // добавить место службы
                Dgv_SaveChanges(IndexRabotaGFS, dgvRabotaGFS);
            }
        }



        //               //""""""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 19-20"  ############################################################
        public void UpdateCard19and20()
        {
            dgvAttestaciya.Rows.Clear();
            dgvAttestaciya.AutoGenerateColumns = false;

            Dgv_Draw(IndexAttestaciya, dgvAttestaciya); // Отрисовываем таблицу dgvAttestaciya

            Attestaciya_Data.Width = 140;
            Attestaciya_Data.MinimumWidth = 140;
            Attestaciya_Prichina.Width = 180;
            Attestaciya_Prichina.MinimumWidth = 180;

            /*==============================================================================================================*/

            dgvProfPodg.Rows.Clear();
            dgvProfPodg.AutoGenerateColumns = false;

            Dgv_Draw(IndexProfPodg, dgvProfPodg); // Отрисовываем таблицу dgvProfPodg

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
                Dgv_SaveChanges(IndexAttestaciya, dgvAttestaciya);
            }
            else //Если метод вызван нажатием кнопки
            {
                dgvAttestaciya.Rows.Add(DateTime.Now.ToString("dd.MM.yyyy"), "Плановая", "Cоответствует замещаемой должности"); // добавить аттестацию
                Dgv_SaveChanges(IndexAttestaciya, dgvAttestaciya);
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
                Dgv_SaveChanges(IndexProfPodg, dgvProfPodg);
            }
            else //Если метод вызван нажатием кнопки
            {
                dgvProfPodg.Rows.Add("Первоначальное обучение", DateTime.Now.ToString("dd.MM.yyyy"), DateTime.Now.ToString("dd.MM.yyyy"), "---", "---"); // добавить проф. подготовку
                Dgv_SaveChanges(IndexProfPodg, dgvProfPodg);
            }
        }



        //               //""""""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 21-22"  ############################################################
        public void UpdateCard21and22()
        {
            Card21and22WasLoaded = 0;
            cbxKlassnost.Text = dgvMainList[IndexKlassnost, IndexRowLichnayaKarta].Value.ToString();
            tbxKlassnostCheyPrikaz.Text = dgvMainList[IndexKlassnostCheyPrikaz, IndexRowLichnayaKarta].Value.ToString();
            tbxKlassnostNomerPrikaza.Text = dgvMainList[IndexKlassnostNomerPrikaza, IndexRowLichnayaKarta].Value.ToString();
            tbxKlassnostDate.Text = dgvMainList[IndexKlassnostDate, IndexRowLichnayaKarta].Value.ToString();

            /*==============================================================================================================*/

            dgvKlassnostOld.Rows.Clear();
            dgvKlassnostOld.AutoGenerateColumns = false;

            Dgv_Draw(IndexKlassnostOld, dgvKlassnostOld); // Отрисовываем таблицу dgvKlassnostOld

            /*==============================================================================================================*/

            dgvNagrady.Rows.Clear();
            dgvNagrady.AutoGenerateColumns = false;

            Dgv_Draw(IndexNagrady, dgvNagrady); // Отрисовываем таблицу dgvNagrady

            Card21and22WasLoaded = 1; // карточка прогрузилась
        }


        // ##########################################################################
        // ##  КНОПКА "ДОБАВИТЬ ПРЕДЫДУЩУЮ КЛАССНОСТЬ" НА ВКЛАДКЕ "КАРТОЧКА 21-22" ##
        // ##########################################################################
        private void KlassnostAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                Dgv_SaveChanges(IndexKlassnostOld, dgvKlassnostOld);
            }
            else //Если метод вызван нажатием кнопки
            {
                dgvKlassnostOld.Rows.Add("Специалист 3 класса", "---", "---", DateTime.Now.ToString("dd.MM.yyyy")); // добавить предыдущую классность
                Dgv_SaveChanges(IndexKlassnostOld, dgvKlassnostOld);
            }
        }


        // ########################################################################
        // ##  КНОПКА "ДОБАВИТЬ НАГРАДЫ / ПООЩРЕНИЯ" НА ВКЛАДКЕ "КАРТОЧКА 21-22" ##
        // ########################################################################
        private void NagradyAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                Dgv_SaveChanges(IndexNagrady, dgvNagrady);
            }
            else //Если метод вызван нажатием кнопки
            {
                dgvNagrady.Rows.Add("---", "---", "---", DateTime.Now.ToString("dd.MM.yyyy")); // добавить награды / поощрения
                Dgv_SaveChanges(IndexNagrady, dgvNagrady);
            }
        }


        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В cbxKlassnost НА ВКЛАДКЕ "КАРТОЧКА 21-22" ##########
        private void Klassnost_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Card21and22WasLoaded == 1)
            {
                dgvMainList[IndexKlassnost, IndexRowLichnayaKarta].Value = cbxKlassnost.Text; // заполняем combobox значением текущей классности
                switch (cbxKlassnost.Text) // проверяем, какую классность выбрал пользователь
                {
                    case "Отсутствует":
                        dgvMainList[IndexKlassnostDate, IndexRowLichnayaKarta].Value = "--.--.----"; // "Обнуляем" дату текущей классности
                        dgvMainList[IndexNextKlassnostDate, IndexRowLichnayaKarta].Value = "--.--.----"; // "Обнуляем" дату следующей классности
                        dgvMainList[IndexKlassnostCheyPrikaz, IndexRowLichnayaKarta].Value = "---"; // "Обнуляем" чей приказ о присвоении классности
                        dgvMainList[IndexKlassnostNomerPrikaza, IndexRowLichnayaKarta].Value = "---"; // "Обнуляем" номер приказа о присвоении классности
                        tbxKlassnostCheyPrikaz.Text = dgvMainList[IndexKlassnostCheyPrikaz, IndexRowLichnayaKarta].Value.ToString(); // Обновляем textbox "Чей приказ" 
                        tbxKlassnostNomerPrikaza.Text = dgvMainList[IndexKlassnostNomerPrikaza, IndexRowLichnayaKarta].Value.ToString(); // Обновляем textbox "Номер приказа" 
                        tbxKlassnostCheyPrikaz.ReadOnly = true; // Если классность отсутствует, окно для ввода должно быть неактивным 
                        tbxKlassnostNomerPrikaza.ReadOnly = true; // Если классность отсутствует, окно для ввода должно быть неактивным
                        break;

                    case "Специалист 3 класса":
                    case "Специалист 2 класса":
                    case "Специалист 1 класса":
                        tbxKlassnostCheyPrikaz.ReadOnly = false;
                        tbxKlassnostNomerPrikaza.ReadOnly = false;
                        dgvMainList[IndexKlassnostDate, IndexRowLichnayaKarta].Value = DateTime.Now.ToString("dd.MM.yyyy"); // выводим дату присвоения классности 
                        dgvMainList[IndexNextKlassnostDate, IndexRowLichnayaKarta].Value = DateTime.Now.AddYears(3).ToString("dd.MM.yyyy"); // дата присвоения, плюс 3 года
                        break;

                    case "Мастер":
                        tbxKlassnostCheyPrikaz.ReadOnly = false;
                        tbxKlassnostNomerPrikaza.ReadOnly = false;
                        dgvMainList[IndexKlassnostDate, IndexRowLichnayaKarta].Value = DateTime.Now.ToString("dd.MM.yyyy"); // выводим дату присвоения классности 
                        dgvMainList[IndexNextKlassnostDate, IndexRowLichnayaKarta].Value = "высшее звание"; // высшая классность
                        break;
                }
                AcceptAndWriteChanges(); // применить изменения
                tbxKlassnostDate.Text = dgvMainList[IndexKlassnostDate, IndexRowLichnayaKarta].Value.ToString(); //обновить окошко даты присвоения классности
            }
        }



        //               //""""""""""""""""""""""""""\\
        // ###############  ВКЛАДКА "КАРТОЧКА 23-25"  ############################################################
        public void UpdateCard23to25()
        {
            Card23to25WasLoaded = 0;
            dgvProdlenie.Rows.Clear();
            dgvProdlenie.AutoGenerateColumns = false;

            Dgv_Draw(IndexProdlenie, dgvProdlenie); // Отрисовываем таблицу dgvProdlenie

            if (dgvProdlenie.Rows.Count != 0) //проверка на существование данных в таблице
            {
                chbxProdlenie.CheckState = CheckState.Checked;
            }
            else
            {
                chbxProdlenie.CheckState = CheckState.Unchecked;
            }

            /*==============================================================================================================*/

            dgvBoevye.Rows.Clear();
            dgvBoevye.AutoGenerateColumns = false;

            Dgv_Draw(IndexBoevye, dgvBoevye); // Отрисовываем таблицу dgvBoevye

            /*==============================================================================================================*/

            dgvRezerv.Rows.Clear();
            dgvRezerv.AutoGenerateColumns = false;

            Dgv_Draw(IndexRezerv, dgvRezerv); // Отрисовываем таблицу dgvRezerv

            Card23to25WasLoaded = 1; // карточка прогрузилась
        }


        // ###############################################################################
        // ##  КНОПКА "ДОБАВИТЬ УЧАСТИЕ В БОЕВЫХ ДЕЙСТВИЯХ" НА ВКЛАДКЕ "КАРТОЧКА 23-25" ##
        // ###############################################################################
        private void BoevyeAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                Dgv_SaveChanges(IndexBoevye, dgvBoevye);
            }
            else //Если метод вызван нажатием кнопки
            {
                dgvBoevye.Rows.Add("---", DateTime.Now.ToString("dd.MM.yyyy"), DateTime.Now.ToString("dd.MM.yyyy"), "1", "---"); // добавить участие в боевых действиях
                Dgv_SaveChanges(IndexBoevye, dgvBoevye);
            }
        }

        // ########################################################################
        // ##  КНОПКА "ДОБАВИТЬ СОСТОЯНИЕ В РЕЗЕРВЕ" НА ВКЛАДКЕ "КАРТОЧКА 23-25" ##
        // ########################################################################
        private void RezervAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                Dgv_SaveChanges(IndexRezerv, dgvRezerv);
            }
            else //Если метод вызван нажатием кнопки
            {
                dgvRezerv.Rows.Add("---", DateTime.Now.ToString("yyyy"), "---", DateTime.Now.ToString("dd.MM.yyyy")); // добавить состояние в резерве
                Dgv_SaveChanges(IndexRezerv, dgvRezerv);
            }
        }


        // ##########  ИЗМЕНЕНИЕ СОСТОЯНИЯ chbxProdlenie НА ВКЛАДКЕ "КАРТОЧКА 23-25" ##########
        private void Prodlenie_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            if (Card23to25WasLoaded == 1)
            {
                if (chbxProdlenie.CheckState == CheckState.Checked)
                {
                    if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
                    {
                        Dgv_SaveChanges(IndexProdlenie, dgvProdlenie);
                    }
                    else //Если метод вызван нажатием кнопки
                    {
                        dgvProdlenie.Rows.Add(DateTime.Now.ToString("dd.MM.yyyy"), "1"); // добавить продление службы
                        Dgv_SaveChanges(IndexProdlenie, dgvProdlenie);
                    }
                }
                else if (chbxProdlenie.CheckState == CheckState.Unchecked)
                {
                    dgvProdlenie.Rows.Clear();
                    dgvMainList[IndexProdlenie, IndexRowLichnayaKarta].Value = "";
                    AcceptAndWriteChanges(); // Применить изменения
                }
                else if (chbxProdlenie.CheckState == CheckState.Indeterminate)
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

            dgvVzyskaniya.Rows.Clear();
            dgvVzyskaniya.AutoGenerateColumns = false;

            Dgv_Draw(IndexVzyskaniya, dgvVzyskaniya); // Отрисовываем таблицу dgvVzyskaniya

            /*==============================================================================================================*/

            dgvUvolnenie.Rows.Clear();
            dgvUvolnenie.AutoGenerateColumns = false;

            Dgv_Draw(IndexUvolnenie, dgvUvolnenie); // Отрисовываем таблицу dgvUvolnenie

            /*==============================================================================================================*/

            tbxZapolnil.Text = dgvMainList[IndexZapolnil, IndexRowLichnayaKarta].Value.ToString();

            /*==============================================================================================================*/

            dtpDataZapolneniya.Text = dgvMainList[IndexDataZapolneniya, IndexRowLichnayaKarta].Value.ToString();

            Card26to29WasLoaded = 1; // карточка прогрузилась
        }



        // ##############################################################
        // ##  КНОПКА "ДОБАВИТЬ ВЗЫСКАНИЕ" НА ВКЛАДКЕ "КАРТОЧКА 26-29" ##
        // ##############################################################
        private void VzyskaniyaAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                Dgv_SaveChanges(IndexVzyskaniya, dgvVzyskaniya);
            }
            else //Если метод вызван нажатием кнопки
            {
                dgvVzyskaniya.Rows.Add("---", "---", "---", "---", DateTime.Now.ToString("dd.MM.yyyy"), "---", "---", DateTime.Now.ToString("dd.MM.yyyy")); // добавить взыскание
                Dgv_SaveChanges(IndexVzyskaniya, dgvVzyskaniya);
            }
        }

        // ###############################################################
        // ##  КНОПКА "ДОБАВИТЬ УВОЛЬНЕНИЕ" НА ВКЛАДКЕ "КАРТОЧКА 26-29" ##
        // ###############################################################
        private void UvolnenieAdd_Click(object sender, EventArgs e)
        {
            if (e is DataGridViewCellEventArgs) //Если метод вызван событием редактирования ячейки таблицы
            {
                Dgv_SaveChanges(IndexUvolnenie, dgvUvolnenie);
            }
            else //Если метод вызван нажатием кнопки
            {
                string UvolnenieProverka = dgvMainList[IndexUvolnenie, IndexRowLichnayaKarta].Value.ToString();
                if (UvolnenieProverka != "") //Если информация об увольнении отсутствует
                {
                    MessageBox.Show("Информация об увольнении уже существует!");
                    return;
                }

                dgvUvolnenie.Rows.Add(DateTime.Now.ToString("dd.MM.yyyy"), "---", "---", DateTime.Now.ToString("dd.MM.yyyy"), "---"); // добавить увольнение
                Dgv_SaveChanges(IndexUvolnenie, dgvUvolnenie);
            }
        }


        // ###########################################################################
        // ##########  ОБЩИЙ МЕТОД ОТРИСОВКИ dataGridView НА ВСЕХ ВКЛАДКАХ  ##########
        // ###########################################################################
        private void Dgv_Draw(int index, DataGridView dgvName)
        {
            string StringDataGrid = dgvMainList[index, IndexRowLichnayaKarta].Value.ToString();
            if (StringDataGrid != "") //проверка на существование данных в таблице
            {
                string[] string_array = StringDataGrid.Split('$');

                foreach (string s in string_array)
                {
                    string[] Row = s.Split('^');
                    dgvName.Rows.Add(Row);
                }
            }
        }

        // #######################################################################################
        // ##########  ОБЩИЙ МЕТОД СОХРАНЕНИЯ ИЗМЕНЕНИЙ В DataGridView НА ВСЕХ ВКЛАДКАХ ##########
        // #######################################################################################
        private void Dgv_SaveChanges(int index, DataGridView dgvName)
        {
            StringBuilder sb = new StringBuilder(); // создаем строку для построения
            foreach (DataGridViewRow row in dgvName.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    // Если обрабатываемая ячейка из колонки типа CalendarColumn, то обрабатываем неверный формат даты.
                    if (cell.OwningColumn is CalendarColumn)
                    {
                        DateTime wrongDateToConvert = DateTime.Parse(cell.Value.ToString()); // парсим её в формат DateTime
                        cell.Value = wrongDateToConvert.ToString("dd.MM.yyyy"); // и конвертируем в нужный формат
                    }
                    sb.Append(cell.Value); // добавляем значение ячейки
                    sb.Append("^"); // ставим разделитель ячеек 
                }
                sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель ячеек
                sb.Append("$"); // ставим разделитель строки
            }
            sb.Remove(sb.Length - 1, 1); // Убираем последний разделитель строки

            dgvMainList[index, IndexRowLichnayaKarta].Value = sb.ToString(); // Заполняем ячейку результирующей строкой
            AcceptAndWriteChanges(); // Применить изменения
        }


        // ##########  ВЫСЧИТЫВАНИЕ ДАТЫ СЛЕДУЮЩЕЙ АТТЕСТАЦИИ ##########
        private void Calculate_NextAttestaciyaDate()
        {
            foreach (DataGridViewRow row in dgvAttestaciya.Rows)
            {
                if (row.Index + 1 == dgvAttestaciya.Rows.Count) //Находим последнюю строку в dgvAttestaciya
                {
                    foreach (DataGridViewCell cell in row.Cells) //Пробегаем по ячейкам в найденной строке
                    {
                        // Обработка неверного формата даты и высчитывание даты следующей аттестации
                        if (cell.OwningColumn is CalendarColumn)
                        {
                            DateTime wrongDateToConvert = DateTime.Parse(cell.Value.ToString()); // парсим её в формат DateTime
                            // Заполняем ячейку "Дата следующей аттестации", прибавив 4 года к последней аттестации
                            dgvMainList[IndexNextAttestaciyaDate, IndexRowLichnayaKarta].Value = wrongDateToConvert.AddYears(4).ToString("dd.MM.yyyy");
                        }
                    }
                }
            }
            AcceptAndWriteChanges(); // Применить изменения
        }


        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В dtpDataPrisyagi НА ВКЛАДКЕ "КАРТОЧКА 16-18" ##########
        private void DataPrisyagi_dateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            if (Card16to18WasLoaded == 1)
            {
                dgvMainList[IndexDataPrisyagi, IndexRowLichnayaKarta].Value = dtpDataPrisyagi.Value.ToString("dd.MM.yyyy");
                AcceptAndWriteChanges(); // применить изменения
            }
        }


        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В tbxKlassnostCheyPrikaz НА ВКЛАДКЕ "КАРТОЧКА 21-22" ##########
        private void KlassnostCheyPrikaz_textBox_TextChanged(object sender, EventArgs e)
        {
            if (Card21and22WasLoaded == 1)
            {
                dgvMainList[IndexKlassnostCheyPrikaz, IndexRowLichnayaKarta].Value = tbxKlassnostCheyPrikaz.Text;
                AcceptAndWriteChanges(); // применить изменения
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В tbxKlassnostNomerPrikaza НА ВКЛАДКЕ "КАРТОЧКА 21-22" ##########
        private void KlassnostNomerPrikaza_textBox_TextChanged(object sender, EventArgs e)
        {
            if (Card21and22WasLoaded == 1)
            {
                dgvMainList[IndexKlassnostNomerPrikaza, IndexRowLichnayaKarta].Value = tbxKlassnostNomerPrikaza.Text;
                AcceptAndWriteChanges(); // применить изменения
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В tbxZapolnil НА ВКЛАДКЕ "КАРТОЧКА 26-29" ##########
        private void Zapolnil_textBox_TextChanged(object sender, EventArgs e)
        {
            if (Card26to29WasLoaded == 1)
            {
                dgvMainList[IndexZapolnil, IndexRowLichnayaKarta].Value = tbxZapolnil.Text;
                AcceptAndWriteChanges(); // применить изменения
            }
        }

        // ##########  СОХРАНЕНИЕ ИЗМЕНЕНИЙ В dtpDataZapolneniya НА ВКЛАДКЕ "КАРТОЧКА 26-29" ##########
        private void DataZapolneniya_dateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            if (Card26to29WasLoaded == 1)
            {
                dgvMainList[IndexDataZapolneniya, IndexRowLichnayaKarta].Value = dtpDataZapolneniya.Value.ToString("dd.MM.yyyy");
                AcceptAndWriteChanges(); // применить изменения
            }
        }


        // ###################################################
        // ##  КНОПКА "ПРЕДЫДУЩАЯ КАРТОЧКА" (СТРЕЛКА ВЛЕВО) ##
        // ###################################################
        private void PrevCard_Click(object sender, EventArgs e)
        {
            if (dgvMainList.Rows.Count == 1)
            {
                MessageBox.Show("Это единственная личная карточка");
            }
            else if (IndexRowLichnayaKarta == 0)
            {
                MessageBox.Show("Это первая личная карточка");
            }
            else
            {
                IndexRowLichnayaKarta--;
                NeedToUpdateCard(); // обновляем все поля личной карточки
            }
            lblCnum.Text = $"{IndexRowLichnayaKarta + 1} из {dgvMainList.RowCount}"; // Порядковый номер личной карточки
            lblCardsFIO.Text = $"{dgvMainList[IndexSurname, IndexRowLichnayaKarta].Value} {dgvMainList[IndexName, IndexRowLichnayaKarta].Value} {dgvMainList[IndexMiddleName, IndexRowLichnayaKarta].Value}"; // Прописываем ФИО над стрелками в карточке
        }



        // ###################################################
        // ##  КНОПКА "СЛЕДУЮЩАЯ КАРТОЧКА" (СТРЕЛКА ВПРАВО) ##
        // ###################################################
        private void NextCard_Click(object sender, EventArgs e)
        {
            if (dgvMainList.Rows.Count == 1)
            {
                MessageBox.Show("Это единственная личная карточка");
            }
            else if (IndexRowLichnayaKarta == dgvMainList.RowCount - 1)
            {
                MessageBox.Show("Это последняя личная карточка");
            }
            else
            {
                IndexRowLichnayaKarta++;
                NeedToUpdateCard(); // обновляем все поля личной карточки
            }
            lblCnum.Text = $"{IndexRowLichnayaKarta + 1} из {dgvMainList.RowCount}"; // Порядковый номер личной карточки
            lblCardsFIO.Text = $"{dgvMainList[IndexSurname, IndexRowLichnayaKarta].Value} {dgvMainList[IndexName, IndexRowLichnayaKarta].Value} {dgvMainList[IndexMiddleName, IndexRowLichnayaKarta].Value}"; // Прописываем ФИО над стрелками в карточке
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
            Required_PersonalFileNum = dgvMainList[IndexPersonalFileNum, IndexRowLichnayaKarta].Value.ToString(); // Искомый № личного дела
            Required_PersonalNum = dgvMainList[IndexPersonalNum, IndexRowLichnayaKarta].Value.ToString(); // Искомый личный номер
            Required_Surname = dgvMainList[IndexSurname, IndexRowLichnayaKarta].Value.ToString(); // Искомая фамилия
            Required_Name = dgvMainList[IndexName, IndexRowLichnayaKarta].Value.ToString(); // Искомое имя
            Required_MiddleName = dgvMainList[IndexMiddleName, IndexRowLichnayaKarta].Value.ToString(); // Искомое отчество
            Required_DateOfBirth = dgvMainList[IndexDateOfBirth, IndexRowLichnayaKarta].Value.ToString(); // Искомая дата рождения
            int NumOfFinds = 0; // Сбрасываем количество найденных записей в ноль

            foreach (DataRow row in dsMain.Tables[CurrentDataTableName].Rows) // Проходим по всем строкам активной DataTable
            {
                if ((row["PersonalFileNum"].ToString() == Required_PersonalFileNum) &&
                    (row["PersonalNum"].ToString() == Required_PersonalNum) &&
                    (row["Surname"].ToString() == Required_Surname) &&
                    (row["Name"].ToString() == Required_Name) &&
                    (row["MiddleName"].ToString() == Required_MiddleName) &&
                    (row["DateOfBirth"].ToString() == Required_DateOfBirth))
                {
                    IndexOfRowToExport = dsMain.Tables[CurrentDataTableName].Rows.IndexOf(row); // Записываем индекс искомой в Datatable строки
                    NumOfFinds++; // Увеличиваем количество найденных записей на единицу
                }
            }
            if (NumOfFinds == 0) MessageBox.Show("Похоже на ошибку в БД. Карточка не была перемещена в архив.");
            if (NumOfFinds == 1) // Если строка найдена и отсутствуют дубли
            {
                dsMain.Tables[OtherDataTableName].ImportRow(dsMain.Tables[CurrentDataTableName].Rows[IndexOfRowToExport]); // Импортируем найденную запись в другой DataTable
                dsMain.Tables[CurrentDataTableName].Rows.RemoveAt(IndexOfRowToExport); // Удаляем найденную запись из текущего DataTable                
                AcceptAndWriteChanges(); // применяем изменения после перемещения строки из одного DataTable в другой
                if (CurrentDataTableName == "Kadry")
                {
                    MessageBox.Show("Карточка успешно перемещена в архив");
                }
                else MessageBox.Show("Карточка успешно восстановлена из архива");

                RefreshDgvMainList(); // обновляем DataGridView1

                if (dgvMainList.Rows.Count != 0) // Проверка dgvMainList на пустоту
                {
                    IndexRowLichnayaKarta = 0; // Делаем активной первую запись, дабы избежать проблемы с несуществующими индексами
                    PereschetCnum(); // пересчитываем порядковые номера
                    AcceptAndWriteChanges(); // сохраняем изменения
                    lblCnum.Text = $"{IndexRowLichnayaKarta + 1} из {dgvMainList.RowCount}"; // Порядковый номер личной карточки
                    lblCardsFIO.Text = $"{dgvMainList[IndexSurname, IndexRowLichnayaKarta].Value} {dgvMainList[IndexName, IndexRowLichnayaKarta].Value} {dgvMainList[IndexMiddleName, IndexRowLichnayaKarta].Value}"; // Прописываем ФИО над стрелками в карточке
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
                if (e.Control is ComboBox cb)
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
            if (sender is ComboBox cbx)
            {
                // Всегда рисуем задний фон
                e.DrawBackground();

                // Drawing one of the items?
                if (e.Index >= 0)
                {
                    // Установка положения строки (alignment). Допустимы значения Center, Near и Far
                    StringFormat sf = new StringFormat
                    {
                        LineAlignment = StringAlignment.Center,
                        Alignment = StringAlignment.Center
                    };

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
            //Rectangle rect = new Rectangle(DataGridView.PointToScreen(cellBounds.Location), cellBounds.Size);
            //Point mousePos = Cursor.Position;
            //bool isHot = rect.Contains(mousePos);

            // Ширина раскрывающейся кнопки и ее положение
            int buttonWidth = SystemInformation.VerticalScrollBarWidth + 3; // Этого должно хватить, чтобы скрыть исходную кнопку
            Rectangle border = BorderWidths(advancedBorderStyle);
            Rectangle buttonRect = new Rectangle(cellBounds.Right - buttonWidth, cellBounds.Top + border.Top, buttonWidth, cellBounds.Height - border.Top - border.Bottom);

            // Расстановка элементов
            StringFormat sf = new StringFormat();
            DataGridViewContentAlignment ali = DataGridViewContentAlignment.TopLeft;
            Brush background = Brushes.White;
            Brush textBrush = Brushes.Black;
            try
            {
                ali = InheritedStyle.Alignment;

                if (Selected)
                {
                    background = new SolidBrush(InheritedStyle.SelectionBackColor);
                    textBrush = new SolidBrush(InheritedStyle.SelectionForeColor);
                }
                else
                {
                    background = new SolidBrush(InheritedStyle.BackColor);
                    textBrush = new SolidBrush(InheritedStyle.ForeColor);
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
                //ComboBoxState state = isHot ? ComboBoxState.Hot : ComboBoxState.Normal;
                //VisualStyleRenderer render = new VisualStyleRenderer("COMBOBOX", (int)COMBOBOXPARTS.CP_READONLY, (int)state);
                //render.DrawBackground(graphics, cellBounds); //отвечает за отрисовку объемного заднего фона ячейки
                //ComboBoxRenderer.DrawDropDownButton(graphics, buttonRect, state); // в оригинале не закомментировано, но при плоской кнопке не нужно
                ControlPaint.DrawComboButton(graphics, buttonRect, ButtonState.Flat); // свойство Flat отвечает за вид кнопки
                textBrush = new SolidBrush(InheritedStyle.ForeColor);
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
                , InheritedStyle.Font
                , textBrush
                , new RectangleF(cellBounds.Left, cellBounds.Top, cellBounds.Width - buttonWidth, cellBounds.Height)
                , sf);
        }

        protected override void OnMouseDown(DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex < 0 || e.RowIndex < 0 || OwningColumn.ReadOnly
                || DataGridView.CurrentCell != this)
            {
                return;
            }

            // Щелчок по ячейке
            //Rectangle rect = DataGridView.GetCellDisplayRectangle(ColumnIndex, RowIndex, false);
            if (!IsInEditMode)
            {
                // Перевести в состояние редактирования, если еще не в нем
                DataGridView.BeginEdit(true);

                if (IsInEditMode)
                {// Находясь в состоянии редактирования, вынимаем элемент управления редактированием (Combobox) и регистрируем обработку событий и т.д.
                    SetEditingComboBox((ComboBox)DataGridView.EditingControl);

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
                DataGridView.Scroll += UpdatePanel;
                DataGridView.RowHeightChanged += UpdatePanel;
                DataGridView.ColumnWidthChanged += UpdatePanel;
                DataGridView.Controls.Add(_Panel);

                DataGridView.CellEndEdit += DataGridView_CellEndEdit;
                UpdatePanel();
            }
            _EditingComboBox.DroppedDown = true;
        }

        /// <summary>Размер элемента DropDown</summary>
        void _EditingComboBox_MeasureItem(object sender, MeasureItemEventArgs e)
        {
            //e.ItemHeight = OwningRow.Height;
            e.ItemWidth = OwningColumn.Width;
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
                brush = new SolidBrush(InheritedStyle.SelectionForeColor);
            }
            else
            {
                brush = new SolidBrush(InheritedStyle.ForeColor);
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
            //int w = OwningColumn.Width; // в оригинале не закомментировано, но нигде не используется
            //int h = OwningRow.Height; // в оригинале не закомментировано, но нигде не используется
            StringFormat sf = new StringFormat
            {
                LineAlignment = StringAlignment.Center,
                Alignment = StringAlignment.Center // задает положение текста в выпадающем списке в режиме редактирования ComboBox'а
            };
            //e.Graphics.DrawString(combo.Items[e.index].ToString(), InheritedStyle.Font, Brushes.Black, new RectangleF(0, 0, w, h), sf);
            e.Graphics.DrawString
                (text
                , InheritedStyle.Font
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
                Rectangle rect = DataGridView.GetCellDisplayRectangle(ColumnIndex, RowIndex, false);

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
                System.Drawing.Drawing2D.GraphicsState gs = e.Graphics.Save();
                //Rectangle rect = DataGridView.GetCellDisplayRectangle(ColumnIndex, RowIndex, false);

                Rectangle bounds = e.ClipRectangle;
                Paint
                    (e.Graphics
                    , bounds
                    , bounds
                    , RowIndex
                    , DataGridViewElementStates.Selected
                    , _EditingComboBox.Text
                    , _EditingComboBox.Text
                    , string.Empty
                    , InheritedStyle
                    , new DataGridViewAdvancedBorderStyle() { All = DataGridViewAdvancedCellBorderStyle.Single }
                    , DataGridViewPaintParts.All);

                e.Graphics.Restore(gs);
            }
        }

        // Отписываемся от событий
        void DataGridView_CellEndEdit(object sender, EventArgs e)
        {
            DataGridView.CellEndEdit -= DataGridView_CellEndEdit;

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

            DataGridView.CellEndEdit -= DataGridView_CellEndEdit;
            DataGridView.Scroll -= UpdatePanel;
            DataGridView.RowHeightChanged -= UpdatePanel;
            DataGridView.ColumnWidthChanged -= UpdatePanel;
            DataGridView.InvalidateCell(this);
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
                return ProcessRightKey(keyData);
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
