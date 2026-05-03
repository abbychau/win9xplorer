using System.Diagnostics;

namespace win9xplorer
{
    internal sealed partial class RetroTaskbarForm
    {
        private void AddressInputHostPanel_Paint(object? sender, PaintEventArgs e)
        {
            var rect = new Rectangle(0, 0, addressInputHostPanel.Width - 1, addressInputHostPanel.Height - 1);
            if (rect.Width <= 0 || rect.Height <= 0)
            {
                return;
            }

            using var fillBrush = new SolidBrush(Color.White);
            e.Graphics.FillRectangle(fillBrush, rect);
            using var darkPen = new Pen(taskbarDarkColor);
            using var lightPen = new Pen(taskbarLightColor);
            e.Graphics.DrawLine(darkPen, rect.Left, rect.Top, rect.Right, rect.Top);
            e.Graphics.DrawLine(darkPen, rect.Left, rect.Top, rect.Left, rect.Bottom);
            e.Graphics.DrawLine(lightPen, rect.Right, rect.Top, rect.Right, rect.Bottom);
            e.Graphics.DrawLine(lightPen, rect.Left, rect.Bottom, rect.Right, rect.Bottom);
        }

        private void CenterAddressTextBoxVertically()
        {
            addressToolbarComboBox.Left = 2;
            addressToolbarComboBox.Width = Math.Max(8, addressInputHostPanel.ClientSize.Width - 4);
            var desiredTop = Math.Max(0, (addressInputHostPanel.ClientSize.Height - addressToolbarComboBox.Height) / 2);
            addressToolbarComboBox.Top = desiredTop;
        }

        private void ApplyClassicAddressComboStyle()
        {
            if (!addressToolbarComboBox.IsHandleCreated)
            {
                return;
            }

            try
            {
                _ = WinApi.SetWindowTheme(addressToolbarComboBox.Handle, string.Empty, string.Empty);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to apply classic address combo style: {ex.Message}");
            }
        }

        private void AddressToolbarComboBox_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter)
            {
                return;
            }

            e.Handled = true;
            e.SuppressKeyPress = true;
            OpenAddressFromToolbar();
        }

        private void OpenAddressFromToolbar()
        {
            var text = addressToolbarComboBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                return;
            }

            if (Directory.Exists(text) || File.Exists(text))
            {
                LaunchProcess("explorer.exe", $"\"{text}\"");
                AddAddressToolbarHistoryEntry(text);
                return;
            }

            if (text.Contains(":") || text.StartsWith("\\", StringComparison.Ordinal))
            {
                LaunchProcess("explorer.exe", text);
                AddAddressToolbarHistoryEntry(text);
                return;
            }

            if (!text.Contains('.'))
            {
                var candidatePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), text);
                if (Directory.Exists(candidatePath))
                {
                    LaunchProcess("explorer.exe", $"\"{candidatePath}\"");
                    AddAddressToolbarHistoryEntry(candidatePath);
                    return;
                }
            }

            if (!Uri.TryCreate(text, UriKind.Absolute, out _))
            {
                text = "https://" + text;
            }

            LaunchProcess(text);
            AddAddressToolbarHistoryEntry(text);
        }

        private void AddAddressToolbarHistoryEntry(string entry)
        {
            var value = entry.Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            addressToolbarHistory.RemoveAll(existing => string.Equals(existing, value, StringComparison.OrdinalIgnoreCase));
            addressToolbarHistory.Insert(0, value);
            if (addressToolbarHistory.Count > AddressHistoryMaxEntries)
            {
                addressToolbarHistory.RemoveRange(AddressHistoryMaxEntries, addressToolbarHistory.Count - AddressHistoryMaxEntries);
            }

            RefreshAddressToolbarHistoryItems();
            addressToolbarComboBox.Text = value;
            addressToolbarComboBox.SelectionStart = addressToolbarComboBox.Text.Length;
            SaveTaskbarSettings();
        }

        private void RefreshAddressToolbarHistoryItems()
        {
            if (addressToolbarComboBox.IsDisposed)
            {
                return;
            }

            var currentText = addressToolbarComboBox.Text;
            addressToolbarComboBox.BeginUpdate();
            addressToolbarComboBox.Items.Clear();
            foreach (var item in addressToolbarHistory)
            {
                addressToolbarComboBox.Items.Add(item);
            }
            addressToolbarComboBox.EndUpdate();
            addressToolbarComboBox.Text = currentText;
            addressToolbarComboBox.SelectionStart = addressToolbarComboBox.Text.Length;
        }
    }
}
