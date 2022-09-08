using System;
using System.Windows.Forms;

namespace RankReminderWinForms
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());

        }
    }


    public class CalendarColumn : DataGridViewColumn
    {
        public CalendarColumn() : base(new CalendarCell())
        {
        }

        public override DataGridViewCell CellTemplate
        {
            get
            {
                return base.CellTemplate;
            }
            set
            {
                // Ensure that the cell used for the template is a CalendarCell.
                if (value != null &&
                    !value.GetType().IsAssignableFrom(typeof(CalendarCell)))
                {
                    throw new InvalidCastException("Must be a CalendarCell");
                }
                base.CellTemplate = value;
            }
        }
    }

    public class CalendarCell : DataGridViewTextBoxCell
    {
        //    ​private DateTimePicker cellDateTimePicker;
        public CalendarCell()
            : base()
        {
            Style.Format = "dd.MM.yyyy"; // После выбора даты выводим в ячейку текст в нужном формате
        }

        public override void InitializeEditingControl(int rowIndex, object
            initialFormattedValue, DataGridViewCellStyle dataGridViewCellStyle)
        {
            int pYear, pMonth, pDay;
            string peremennaya;


            // Set the value of the editing control to the current cell value.
            base.InitializeEditingControl(rowIndex, initialFormattedValue,
                dataGridViewCellStyle);
            CalendarEditingControl ctl =
                DataGridView.EditingControl as CalendarEditingControl;
            // Use the default row value when Value property is null.
            //ctl.CustomFormat = "dd/MM/yyyy";
            // ctl.Format = DateTimePickerFormat.Custom;

            if ((Value == null) || (Value.ToString() == "---"))
            {
                ctl.Value = DateTime.Now;
            }
            else if (Value.ToString() == "не установлена")
            {
                MessageBox.Show("Срок выслуги в текущем звании не установлен");
                SendKeys.Send("{ESC}");
            }
            else if (Value.ToString() == "роста нет")
            {
                MessageBox.Show("Текущее звание равно, либо превышает звание по должности");
                SendKeys.Send("{ESC}");
            }
            else if (Value.ToString() == "высшее звание")
            {
                MessageBox.Show("Сотрудник уже имеет высшее квалификационное звание");
                SendKeys.Send("{ESC}");
            }
            else if (Value.ToString() == "--.--.----")
            {
                MessageBox.Show("Сотрудник не имеет квалификационного звания");
                SendKeys.Send("{ESC}");
            }
            else
            {
                peremennaya = Value.ToString(); // считываем значение даты из ячейки в peremennaya типа string
                //MessageBox.Show(Value.ToString());
                bool isLetter = !String.IsNullOrEmpty(peremennaya) && char.IsLetter(peremennaya[0]) || char.IsLetter(peremennaya[1]); // проверяем, если ячейка начинается с сокращенного дня недели

                if (isLetter == true) // если текстовая дата записана в неправильном формате и начинается с сокращенного дня недели (ПН, ВТ, СР и т.д.)
                {
                    DateTime wrongdatetoconvert = DateTime.Parse(peremennaya); // парсим её в формат DateTime
                    peremennaya = wrongdatetoconvert.ToString("dd.MM.yyyy"); // и конвертируем в нужный формат
                }

                pYear = Convert.ToInt32(peremennaya.Substring(6, 4)); // парсим peremennaya с 7го символа, длина - 4 символа
                //MessageBox.Show(pYear.ToString());

                pMonth = Convert.ToInt32(peremennaya.Substring(3, 2)); // парсим peremennaya с 4го символа, длина - 2 символа
                //MessageBox.Show(pMonth.ToString());

                pDay = Convert.ToInt32(peremennaya.Substring(0, 2)); // парсим peremennaya с 1го символа, длина - 2 символа                                                                    

                DateTime proverka = new DateTime(pYear, pMonth, pDay);
                ctl.Value = (DateTime)proverka;
            }
        }


        public override Type EditType
        {
            get
            {
                // Return the type of the editing control that CalendarCell uses.
                return typeof(CalendarEditingControl);
            }
        }

        public override Type ValueType
        {
            get
            {
                // Return the type of the value that CalendarCell contains.

                return typeof(DateTime);
            }
        }

        public override object DefaultNewRowValue
        {
            get
            {
                // Use the current date and time as the default value.
                return DateTime.Now.ToString("dd.MM.yyyy");
            }
        }
    }

    class CalendarEditingControl : DateTimePicker, IDataGridViewEditingControl
    {
        DataGridView dataGridView;
        private bool valueChanged = false;
        int rowIndex;

        public CalendarEditingControl() // отвечает за отображение даты в ячейке во время редактирования (выбора даты)
        {
            Format = DateTimePickerFormat.Short;  // день недели сокращенно и дата

            // Format = DateTimePickerFormat.Custom;
            // CustomFormat = "dd.MM.yyyy";
        }

        // Implements the IDataGridViewEditingControl.EditingControlFormattedValue
        // property.
        public object EditingControlFormattedValue
        {
            get
            {
                //return Value.ToShortDateString();
                return Value.ToString("dd.MM.yyyy"); //Этот параметр отвечает за вывод даты в datagrid в нужном формате
            }
            set
            {
                if (value is String @string)
                {
                    try
                    {
                        // This will throw an exception of the string is
                        // null, empty, or not in the format of a date.
                        Value = DateTime.Parse(@string);
                    }
                    catch
                    {
                        // In the case of an exception, just use the
                        // default value so we're not left with a null
                        // value.
                        Value = DateTime.Now;
                    }
                }
            }
        }

        // Implements the
        // IDataGridViewEditingControl.GetEditingControlFormattedValue method.
        public object GetEditingControlFormattedValue(
            DataGridViewDataErrorContexts context)
        {
            return EditingControlFormattedValue;
        }

        // Implements the
        // IDataGridViewEditingControl.ApplyCellStyleToEditingControl method.
        public void ApplyCellStyleToEditingControl(
            DataGridViewCellStyle dataGridViewCellStyle)
        {
            Font = dataGridViewCellStyle.Font;
            //CalendarForeColor = dataGridViewCellStyle.ForeColor;
            //CalendarMonthBackground = dataGridViewCellStyle.BackColor;
        }

        // Implements the IDataGridViewEditingControl.EditingControlRowIndex
        // property.
        public int EditingControlRowIndex
        {
            get
            {
                return rowIndex;
            }
            set
            {
                rowIndex = value;
            }
        }

        // Implements the IDataGridViewEditingControl.EditingControlWantsInputKey
        // method.
        public bool EditingControlWantsInputKey(
            Keys key, bool dataGridViewWantsInputKey)
        {
            // Let the DateTimePicker handle the keys listed.
            switch (key & Keys.KeyCode)
            {
                case Keys.Left:
                case Keys.Up:
                case Keys.Down:
                case Keys.Right:
                case Keys.Home:
                case Keys.End:
                case Keys.PageDown:
                case Keys.PageUp:
                    return true;
                default:
                    return !dataGridViewWantsInputKey;
            }
        }

        // Implements the IDataGridViewEditingControl.PrepareEditingControlForEdit
        // method.
        public void PrepareEditingControlForEdit(bool selectAll)
        {
            // No preparation needs to be done.
        }

        // Implements the IDataGridViewEditingControl
        // .RepositionEditingControlOnValueChange property.
        public bool RepositionEditingControlOnValueChange
        {
            get
            {
                return false;
            }
        }

        // Implements the IDataGridViewEditingControl
        // .EditingControlDataGridView property.
        public DataGridView EditingControlDataGridView
        {
            get
            {
                return dataGridView;
            }
            set
            {
                dataGridView = value;
            }
        }

        // Implements the IDataGridViewEditingControl
        // .EditingControlValueChanged property.
        public bool EditingControlValueChanged
        {
            get
            {
                return valueChanged;
            }
            set
            {
                valueChanged = value;
            }
        }

        // Implements the IDataGridViewEditingControl
        // .EditingPanelCursor property.
        public Cursor EditingPanelCursor
        {
            get
            {
                return base.Cursor;
            }
        }

        protected override void OnValueChanged(EventArgs eventargs)
        {
            // Notify the DataGridView that the contents of the cell
            // have changed.
            valueChanged = true;
            EditingControlDataGridView.NotifyCurrentCellDirty(true);
            base.OnValueChanged(eventargs);
        }
    }

}
