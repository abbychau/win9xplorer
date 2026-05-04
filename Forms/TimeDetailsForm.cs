using System.Globalization;

namespace win9xplorer
{
    internal sealed class TimeDetailsForm : Form
    {
        private readonly Label dateLabel;
        private readonly Label timeLabel;
        private readonly Label lunarDateLabel;
        private readonly Label lunarTimerLabel;
        private readonly MonthCalendar monthCalendar;
        private readonly System.Windows.Forms.Timer tickTimer;

        private static readonly ChineseLunisolarCalendar LunarCalendar = new();
        private static readonly string[] LunarMonthNames =
        {
            "",
            "正月", "二月", "三月", "四月", "五月", "六月",
            "七月", "八月", "九月", "十月", "冬月", "臘月"
        };

        private static readonly string[] LunarDayNames =
        {
            "",
            "初一", "初二", "初三", "初四", "初五", "初六", "初七", "初八", "初九", "初十",
            "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十",
            "廿一", "廿二", "廿三", "廿四", "廿五", "廿六", "廿七", "廿八", "廿九", "三十"
        };

        private static readonly string[] EarthlyBranches =
        {
            "子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥"
        };

        private static readonly string[] KeNames =
        {
            "", "一", "二", "三", "四", "五", "六", "七", "八"
        };

        public TimeDetailsForm()
        {
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            TopMost = true;
            MaximizeBox = false;
            MinimizeBox = false;
            Text = "Date and Time";
            ClientSize = new Size(280, 316);
            Font = new Font("MS Sans Serif", 8.25f, FontStyle.Regular, GraphicsUnit.Point);

            dateLabel = new Label
            {
                Left = 12,
                Top = 14,
                Width = 130,
                Height = 28,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("MS Sans Serif", 10f, FontStyle.Bold, GraphicsUnit.Point)
            };

            timeLabel = new Label
            {
                Left = 146,
                Top = 14,
                Width = 122,
                Height = 28,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleRight,
                Font = new Font("MS Sans Serif", 10f, FontStyle.Bold, GraphicsUnit.Point)
            };

            lunarDateLabel = new Label
            {
                Left = 12,
                Top = 48,
                Width = 170,
                Height = 28,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("MS Sans Serif", 10f, FontStyle.Bold, GraphicsUnit.Point)
            };

            lunarTimerLabel = new Label
            {
                Left = 186,
                Top = 48,
                Width = 82,
                Height = 28,
                TextAlign = ContentAlignment.MiddleRight,
                Font = new Font("MS Sans Serif", 10f, FontStyle.Bold, GraphicsUnit.Point)
            };

            monthCalendar = new MonthCalendar
            {
                Left = 30,
                Top = 120,
                MaxSelectionCount = 1,
                ShowTodayCircle = true,
                ShowToday = true
            };

            Controls.Add(dateLabel);
            Controls.Add(timeLabel);
            Controls.Add(lunarDateLabel);
            Controls.Add(lunarTimerLabel);
            Controls.Add(monthCalendar);

            tickTimer = new System.Windows.Forms.Timer { Interval = 1000 };
            tickTimer.Tick += (_, _) => UpdateDateTime();

            VisibleChanged += (_, _) =>
            {
                if (Visible)
                {
                    UpdateDateTime();
                    tickTimer.Start();
                }
                else
                {
                    tickTimer.Stop();
                }
            };

            Deactivate += (_, _) => Hide();

            FormClosed += (_, _) =>
            {
                tickTimer.Stop();
                tickTimer.Dispose();
                dateLabel.Font.Dispose();
                timeLabel.Font.Dispose();
                lunarDateLabel.Font.Dispose();
                lunarTimerLabel.Font.Dispose();
            };
        }

        public void ShowAt(Point location)
        {
            Location = location;
            UpdateDateTime();
            monthCalendar.SetDate(DateTime.Today);
            Show();
            Activate();
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                Hide();
                return true;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void UpdateDateTime()
        {
            var now = DateTime.Now;
            dateLabel.Text = now.ToString("yyyy/MM/dd");
            timeLabel.Text = now.ToString("tt hh:mm:ss");
            lunarDateLabel.Text = GetLunarDateText(now);
            lunarTimerLabel.Text = GetLunarHourText(now);
        }

        private static string GetLunarDateText(DateTime date)
        {
            var year = LunarCalendar.GetYear(date);
            var month = LunarCalendar.GetMonth(date);
            var day = LunarCalendar.GetDayOfMonth(date);
            var leapMonth = LunarCalendar.GetLeapMonth(year);
            var isLeapMonth = leapMonth > 0 && month == leapMonth;

            if (leapMonth > 0 && month > leapMonth)
            {
                month--;
            }

            var monthText = month >= 1 && month < LunarMonthNames.Length ? LunarMonthNames[month] : $"{month}月";
            var dayText = day >= 1 && day < LunarDayNames.Length ? LunarDayNames[day] : $"{day}日";
            return isLeapMonth ? $"農曆 閏{monthText}{dayText}" : $"農曆 {monthText}{dayText}";
        }

        private static string GetLunarHourText(DateTime date)
        {
            var branchIndex = ((date.Hour + 1) / 2) % 12;
            var branch = EarthlyBranches[branchIndex];

            var startHour = branchIndex == 0 ? 23 : (branchIndex * 2) - 1;
            var startTotalMinutes = startHour * 60;
            var currentTotalMinutes = date.Hour * 60 + date.Minute;
            if (branchIndex == 0 && date.Hour < 12)
            {
                currentTotalMinutes += 24 * 60;
            }

            var elapsedMinutes = Math.Clamp(currentTotalMinutes - startTotalMinutes, 0, 119);
            var ke = Math.Clamp((elapsedMinutes / 15) + 1, 1, 8);
            return $"{branch}時{KeNames[ke]}刻";
        }
    }
}
