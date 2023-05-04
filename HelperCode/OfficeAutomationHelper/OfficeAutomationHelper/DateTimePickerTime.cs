using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.ComponentModel;
using System.Windows.Forms;

public class DateTimePickerIncTime : DateTimePicker
{

    //This is based on 5 minute increments
    public enum MinuteIncrements {None = 0,Five = 1,Ten = 2,Fifteen = 3,Thirty = 6}

    public DateTimePickerIncTime()
    {
        ValueChanged += DateTimePickerIncTime_ValueChanged;
        this.Format = DateTimePickerFormat.Custom;
        this.CustomFormat = "hh:mm tt";
        this.ShowUpDown = true;
        this.Value = new DateTime(this.Value.Year, this.Value.Month, this.Value.Day, 12, 0, 0);
        this.Width = 70;
    }

    private MinuteIncrements _MinuteIncrement = MinuteIncrements.Fifteen;
    [Description("Minutes to increment with Arrows and Buttons.")]
    public MinuteIncrements MinuteIncrement
    {
        get { return _MinuteIncrement; }
        set { _MinuteIncrement = value; }
    }

    private void DateTimePickerIncTime_ValueChanged(object sender, System.EventArgs e)
    {
        DateTimePickerIncrementChange((DateTimePicker)sender);
    }


    private void DateTimePickerIncrementChange(DateTimePicker myDateTimePicker)
    {
        const int FiveMinutes = 5;

        int myNewMinute = -1;
        int myMinuteInc = FiveMinutes * (int)_MinuteIncrement;
        var _with1 = myDateTimePicker.Value;
        if (_MinuteIncrement > 0)
        {
            if (_with1.Minute % myMinuteInc == 1)
            {
                myNewMinute = _with1.Minute - 1 + myMinuteInc;
                if (myNewMinute > 59)
                    myNewMinute = 0;
            }
            else if (_with1.Minute % myMinuteInc != 0)
            {
                myNewMinute = (_with1.Minute / myMinuteInc) * myMinuteInc;
            }

            if (myNewMinute >= 0)
            {
                myDateTimePicker.Value = new DateTime(_with1.Year, _with1.Month, _with1.Day, _with1.Hour, myNewMinute, 0);
            }

        }
    }
}
