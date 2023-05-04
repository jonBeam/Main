using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;

namespace AG.Windows.Controls
{
    public class DateTimePickerExtended : System.Windows.Forms.DateTimePicker
    {
        #region Member variables

        public enum SizeWidths : int { Normal=200, Small=100 }
        public enum ToolTipAges : byte { None, YearsMonthsDays, YearsDays }
        public enum DateRegulations:byte { None, PastOnly, FutureOnly }

        private ToolTip Tooltip = new System.Windows.Forms.ToolTip();
        private bool _IsNull = true;  // true, when no date shall be displayed (empty DateTimePicker)
        public delegate void ValidationDateoutofRange(object sender, EventArgs e, string Message);
        public event ValidationDateoutofRange ValidatingDateoutofRange;
        #endregion

        #region Constructor
        public DateTimePickerExtended() : base()
        {
            base.Format = DateTimePickerFormat.Custom;  //must set custom format type
            this.Value = null;  //since this is a null-able control, set it to null initially
            base.Format = DateTimePickerFormat.Custom;  //must set custom format type
        }
        #endregion

        #region Properties

        [DefaultValue(null)]
        public new Object Value
        {
            get
            {
                if (_IsNull)
                    return null;
                else
                    return base.Value;
            }
            set
            {
                if (value == null || value == DBNull.Value)
                {
                    SetToNullValue();
                }
                else
                {
                    SetToDateTimeValue();
                    base.Value = (DateTime)value;
                }
            }
        }

        private DateTimePickerFormat _Format = DateTimePickerFormat.Long; 
        [Browsable(true),
         TypeConverter(typeof(Enum)),
         DefaultValue(DateTimePickerFormat.Long)]
        public new DateTimePickerFormat Format
        {
            get { return _Format; }
            set
            {
                _Format = value;
                SetFormat();
                OnFormatChanged(EventArgs.Empty);
            }
        }

        private string _customFormat;  // The custom format of the DateTimePicker control
        private new String CustomFormat
        {
            get { return _customFormat; }
            set { _customFormat = value; }
        }

        /// <summary>
        /// Returns a Parsed Date Time or null value.
        /// </summary>
        [DebuggerStepThrough]
        public Object ParsedDateTime(Object value, bool InDatePickerMinMaxRange)
        {
            DateTime dateTime;
            DateTime.TryParse(this.Value.ToString(), out dateTime);

            if (dateTime != null && InDatePickerMinMaxRange && IsDateInDatePickerMinMaxRange(dateTime)) 
            {
            return (dateTime);
            }
            else if (dateTime != null && !InDatePickerMinMaxRange)
            {
                return (dateTime);
            }
            else
            {
                return (null);
            }
        }

        /// <summary>
        /// Returns a Boolean value indicating whether an expressions represents a valid Date value within the Min and Max Range.
        /// </summary>
        [DebuggerStepThrough]
        public bool IsDateInDatePickerMinMaxRange(Object value)
        {
            return (this.ParsedDateTime(value, true) != null);
        }

        /// <summary>
        /// Returns a Boolean value indicating whether this value represents a valid Date.
        /// </summary>
        [DebuggerStepThrough]
        public bool IsDate()
        {
            return (this.Value != null && this.ParsedDateTime(this.Value,false)  != null);
        }

        /// <summary>
        /// Returns a Boolean value indicating whether an expressions represents a valid Date value.
        /// </summary>
        [DebuggerStepThrough]
        public bool IsDate(Object value)
        {
            return (this.Value != null && this.ParsedDateTime(value, false) != null);
        }

        /// <summary>
        /// Returns a Boolean value indicating whether an expressions represents a valid Date value within the Min and Max Range.
        /// </summary>
        [DebuggerStepThrough]
        public bool IsDateInDatePickerMinMaxRange(DateTime dateTime)
        {
            return (dateTime >= this.MinDate && dateTime <= this.MaxDate);
        }

        private bool _NullOnBackspaceKeyHit = true;
        [Browsable(true),
         Category("Behavior"),
         Description("Determines if the date is set to null when the backspace key is hit"),
         DefaultValue(true)]
        public bool NullOnBackspaceKeyHit
        {
            get { return _NullOnBackspaceKeyHit; }
            set { _NullOnBackspaceKeyHit = value; }
        }

        private SizeWidths _SizeWidth = SizeWidths.Normal;
        [Browsable(true),
         Category("Behavior"),
         Description("Size the width to the typical sizes"),
         DefaultValue(true)]
        public SizeWidths SizeWidth
        {
            get { return _SizeWidth; }
            set 
            {
                _SizeWidth = value;
                if (_SizeWidth == SizeWidths.Small)
                {
                    this.Width = Convert.ToInt16(SizeWidths.Small);
                }
                else
                {
                    this.Width = Convert.ToInt16(SizeWidths.Normal);
                }
                this.Invalidate();
            }
        }


        private bool _NullOnDeleteKeyHit = true;
        [Browsable(true),
         Category("Behavior"),
         Description("Determines if the date is set to null when the delete key is hit"),
         DefaultValue(true)]

        public bool NullOnDeleteKeyHit
        {
            get { return _NullOnDeleteKeyHit; }
            set { _NullOnDeleteKeyHit = value; }
        }

        private bool _IsNullable = true; 
        [Browsable(true),
         Category("Behavior"),
         Description("Determines if the date can be a null value"),
         DefaultValue(true)]
        public bool IsNullable
        {
            get { return _IsNullable; }
            set { _IsNullable = value; }
        }
        
        private string _NullValue = " "; // If _IsNull = true, this value is shown in the DTP
        [Browsable(false),
         Category("Behavior"),
         Description("The string used to display null values in the control"),
         DefaultValue(" ")]
        public String NullValue
        {
            get { return _NullValue; }
            set { _NullValue = value; }
        }

        private DateRegulations _DateRegulation = DateRegulations.None;
        [Browsable(true),
         CategoryAttribute("Behavior"),
         DescriptionAttribute("Force dates past or future"),
         TypeConverter(typeof(DateRegulations)),
         DefaultValue(DateRegulations.None)]
        public DateRegulations DateRegulation
        {
            get { return _DateRegulation; }
            set { _DateRegulation = value; }
        }

        
        private ToolTipAges _ToolTipAge = ToolTipAges.YearsMonthsDays;
        [CategoryAttribute("Appearance"),
         DescriptionAttribute("Displayed Tooltip format."),
         TypeConverter(typeof(ToolTipAges)),
         DefaultValue(ToolTipAges.YearsMonthsDays)]
        public ToolTipAges ToolTipAge
        {
            get { return _ToolTipAge; }
            set { _ToolTipAge = value; }
        }

        private bool _TooltipActive = true;
        [CategoryAttribute("Behavior"),
         DescriptionAttribute("Determines if the tool tip is active."),
         DefaultValue(true)]
        public bool TooltipActive
        {
            get { return _TooltipActive; }
            set
            {
                _TooltipActive = value;
                Tooltip.Active = value;
            }
        }

        private bool _TooltipShowAlways = true;
        [CategoryAttribute("Behavior"),
         DescriptionAttribute("Determines if the tool tip will display always, "
                           + "even if the parent window is not active."),
         DefaultValue(true)]
        public bool TooltipShowAlways
        {
            get { return _TooltipShowAlways; }
            set
            {
                _TooltipShowAlways = value;
                Tooltip.ShowAlways = value;
            }
        }

        private int _TooltipAutomaticDelay = 100;  //default for tooltip was 500
        [CategoryAttribute("Behavior"),
         DescriptionAttribute("Sets the values of the InitalDelay, AutoPopDelay "
                            + ", and ReshowDelay to their relative values"),
         DefaultValue(100)]
        public int TooltipAutomaticDelay
        {
            get { return _TooltipAutomaticDelay; }
            set
            {
                _TooltipAutomaticDelay = value;
                Tooltip.AutomaticDelay = value;
            }
        }

        private int _TooltipReshowDelay = 100;
        [CategoryAttribute("Behavior"),
         DescriptionAttribute("Determines the length of time it takes for "
                            + "subsequent ToolTip windows to appears as the "
                            + "Pointer moves from one ToolTip region to another."),
         DefaultValue(100)]
        public int TooltipReshowDelay
        {
            get { return _TooltipReshowDelay; }
            set
            {
                _TooltipReshowDelay = value;
                Tooltip.ReshowDelay = value;
            }
        }

        private int _TooltipAutoPopDelay = 10000;  //default for tooltip is 5000
        [CategoryAttribute("Behavior"),
         DescriptionAttribute("Determines the length of time the ToolTip "
                            + "window remains visible if the pointer is "
                            + "stationary inside the ToolTip region."),
         DefaultValue(10000)]
        public int TooltipAutoPopDelay
        {
            get { return _TooltipAutoPopDelay; }
            set
            {
                _TooltipAutoPopDelay = value;
                Tooltip.AutoPopDelay = value;
            }
        }

        private string _FormatAsString;
        private string FormatAsString
        {
            get { return _FormatAsString; }
            set
            {
                _FormatAsString = value;
                base.CustomFormat = value;
            }
        }

        private void SetToNullValue()
        {
            _IsNull = true;
            if (_NullValue == null || _NullValue == String.Empty)
            {
                base.CustomFormat = " ";
            }
            else
            {
                base.CustomFormat = "'" + _NullValue + "'";
            }
        }

        private void SetToDateTimeValue()
        {
            if (_IsNull)
            {
                SetFormat();
                _IsNull = false;
                base.OnValueChanged(new EventArgs());
            }
        }

        private void SetFormat()
        {
            CultureInfo ci = Thread.CurrentThread.CurrentCulture;
            DateTimeFormatInfo dtf = ci.DateTimeFormat;
            switch (_Format)
            {
                case DateTimePickerFormat.Long:
                    FormatAsString = dtf.LongDatePattern;
                    break;
                case DateTimePickerFormat.Short:
                    FormatAsString = dtf.ShortDatePattern;
                    break;
                //case DateTimePickerFormat.Time:
                //    FormatAsString = dtf.ShortTimePattern;
                //    break;
                case DateTimePickerFormat.Custom:
                    FormatAsString = this.CustomFormat;
                    break;
            }
        }


        public string Age(object value)
        {

            if (value != null && ToolTipAge != ToolTipAges.None)
            {
                DateTime dateTime;
                DateTime.TryParse(value.ToString(), out dateTime);
                dateTime = dateTime.Date;

                string formatted = string.Empty;

                EllapsedYearsMonthsDays ellapsedYearsMonthsDays = new EllapsedYearsMonthsDays(DateTime.Today.Date, dateTime);

                if (ellapsedYearsMonthsDays.elapsedYears != 0)
                {
                    if (ellapsedYearsMonthsDays.elapsedYears > 1)
                        formatted = String.Format(" {0} years", ellapsedYearsMonthsDays.elapsedYears);
                    else
                        formatted = String.Format(" {0} year", ellapsedYearsMonthsDays.elapsedYears);
                }

                if (ToolTipAge == ToolTipAges.YearsMonthsDays && ellapsedYearsMonthsDays.elapsedMonths != 0)
                {
                    if (ellapsedYearsMonthsDays.elapsedMonths > 1)
                        formatted += String.Format(" {0} months", ellapsedYearsMonthsDays.elapsedMonths);
                    else
                        formatted += String.Format(" {0} month", ellapsedYearsMonthsDays.elapsedMonths);
                }

                int days = 0;
                if (ToolTipAge == ToolTipAges.YearsMonthsDays)
                {
                    days = ellapsedYearsMonthsDays.elapsedDays;
                }
                else if (ToolTipAge == ToolTipAges.YearsDays)
                {
                    days = ellapsedYearsMonthsDays.elapsedDaysInYear;
                }

                if (days > 1)
                  formatted += String.Format(" {0} days", days);
                else if (days == 1)
                  formatted += String.Format(" {0} day", days);
                

                if (DateTime.Today.Date.Subtract(dateTime).Days == 0)
                    formatted += " today";
                else if (dateTime > DateTime.Today)
                    formatted += " from today";
                else
                    formatted += " ago";

                return (formatted.Trim());
            }
            else return ("");
        }

        #endregion



        #region OnOverride()

        protected override void OnCloseUp(EventArgs e)
        {
            if (Control.MouseButtons == MouseButtons.None)
            {
                if (_IsNull)
                {
                    SetToDateTimeValue();
                }
            }
            base.OnCloseUp(e);
        }

        protected override void OnKeyUp(KeyEventArgs e)
        {
            if (IsNullable && ( (e.KeyCode == Keys.Delete && NullOnDeleteKeyHit )
                             || (e.KeyCode == Keys.Back && NullOnBackspaceKeyHit)))
            {
                this.Value = null;
                OnValueChanged(EventArgs.Empty);
            }
            base.OnKeyUp(e);
        }

        protected override void OnMouseHover(EventArgs e)
        {
            Tooltip.SetToolTip(this, Age(this.Value));
            base.OnMouseHover(e);
        }


        protected override void OnValidating(CancelEventArgs e)
        {
            if (this.Value != null)
            {
                if (this.Value.ToString() != String.Empty)
                {
                    DateTime dateTime;
                    DateTime.TryParse(Value.ToString(), out dateTime);

                    if (dateTime != null)
                    {
                        if (DateRegulation == DateRegulations.FutureOnly && dateTime <= DateTime.Today)
                        {
                            e.Cancel = true;
                            ValidatingDateoutofRange(this, e, "Past date not allowed");
                        }
                        else if (DateRegulation == DateRegulations.PastOnly && dateTime >= DateTime.Today)
                        {
                            e.Cancel = true;
                            ValidatingDateoutofRange(this, e, "Future date not allowed");
                        }
                        else
                        {
                            base.OnValidating(e);
                        }
                    }
                }
            }
        }

        #endregion
        

        public sealed class EllapsedYearsMonthsDays
        {

            private int _elapsedYears;
            public int elapsedYears
            {
                get { return _elapsedYears; }
                set { _elapsedYears = value; }
            }

            private int _elapsedMonths;
            public int elapsedMonths
            {
                get { return _elapsedMonths; }
                set { _elapsedMonths = value; }
            }

            private int _elapsedDays;
            public int elapsedDays
            {
                get { return _elapsedDays; }
                set { _elapsedDays = value; }
            }

            private int _elapsedDaysInYear;
            public int elapsedDaysInYear
            {
                get { return _elapsedDaysInYear; }
                set { _elapsedDaysInYear = value; }
            }


            public EllapsedYearsMonthsDays(DateTime date1, DateTime date2)
            {
                
                int years, months, days;
                if (date1 < date2)
                {
                    DateTime _date = date1;
                    date1 = date2;
                    date2 = _date;
                }

                months = 12 * (date1.Year - date2.Year) + (date1.Month - date2.Month);

                if (date1.Day < date2.Day)
                {
                    months--;
                    days = DateTime.DaysInMonth(date2.Year, date2.Month) - date2.Day + date1.Day;
                }
                else
                {
                    days = date1.Day - date2.Day;
                }

                years = months / 12;
                months -= years * 12;

                elapsedYears = years;
                elapsedMonths = months;
                elapsedDays = days;


                //days of year section
                int daysofYear;
                if (date1.DayOfYear - date2.DayOfYear >= 0)
                {
                    daysofYear = date1.DayOfYear - date2.DayOfYear;
                }
                else
                {
                    int daysInLastYear = new DateTime(date2.Year, 1, 1).AddDays(-1).DayOfYear;
                    daysofYear = (daysInLastYear-date2.DayOfYear) + date1.DayOfYear;
                }
                elapsedDaysInYear = daysofYear;
            }
        }

        
    }
}
