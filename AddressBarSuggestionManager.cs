using System.Diagnostics;

namespace win9xplorer
{
    /// <summary>
    /// Manages address bar autocompletion suggestions
    /// </summary>
    internal class AddressBarSuggestionManager
    {
        private ListBox? suggestionListBox;
        private Form? parentForm;
        private ToolStripTextBox? addressTextBox;
        private bool isShowing = false;
        private string lastSuggestionPath = "";
        private bool isNavigatingProgrammatically = false;
        private string originalAddressText = ""; // Track original text before suggestion navigation
        private bool isNavigatingSuggestions = false; // Track if we're navigating through suggestions

        public void SetupSuggestions(Form parentForm, ToolStripTextBox addressTextBox)
        {
            this.parentForm = parentForm;
            this.addressTextBox = addressTextBox;
            
            // Create the suggestion listbox with improved styling
            suggestionListBox = new ListBox
            {
                Visible = false,
                Font = addressTextBox.Font,
                BorderStyle = BorderStyle.FixedSingle,
                IntegralHeight = false,
                MaximumSize = new Size(400, 150),
                BackColor = SystemColors.Window,
                ForeColor = SystemColors.WindowText,
                SelectionMode = SelectionMode.One,
                TabStop = false
            };
            
            // Add the listbox to the parent form
            parentForm.Controls.Add(suggestionListBox);
            
            // Ensure it's on top of other controls
            suggestionListBox.BringToFront();
            
            // Set up event handlers
            suggestionListBox.Click += SuggestionListBox_Click;
            suggestionListBox.KeyDown += SuggestionListBox_KeyDown;
            suggestionListBox.SelectedIndexChanged += SuggestionListBox_SelectedIndexChanged;
            suggestionListBox.MouseEnter += (s, e) => 
            {
                // Ensure the suggestion box stays focused when mouse enters
                if (isShowing && suggestionListBox != null)
                {
                    suggestionListBox.BringToFront();
                }
            };
            
            // Set up text change handler for the address textbox
            addressTextBox.TextChanged += AddressTextBox_TextChanged;
            addressTextBox.KeyDown += AddressTextBox_KeyDown;
            addressTextBox.LostFocus += AddressTextBox_LostFocus;
        }

        private void AddressTextBox_TextChanged(object? sender, EventArgs e)
        {
            if (addressTextBox == null || parentForm == null)
                return;
                
            // Don't show suggestions if we're navigating programmatically (mouse clicks, etc.)
            if (isNavigatingProgrammatically)
            {
                Debug.WriteLine("Ignoring TextChanged - navigating programmatically");
                return;
            }
            
            // Don't interfere if we're currently navigating through suggestions
            if (isNavigatingSuggestions)
            {
                Debug.WriteLine("Ignoring TextChanged - navigating through suggestions");
                return;
            }
                
            string currentText = addressTextBox.Text;
            Debug.WriteLine($"AddressTextBox_TextChanged: '{currentText}'");
            
            // Check if we should show suggestions
            if (ShouldShowSuggestions(currentText))
            {
                ShowSuggestions(currentText);
            }
            else
            {
                HideSuggestions();
            }
        }

        private bool ShouldShowSuggestions(string text)
        {
            // Show suggestions when:
            // 1. Text ends with backslash (indicating path continuation)
            // 2. Text looks like a partial path (contains backslash)
            // 3. Text is empty (show drives)
            
            if (string.IsNullOrEmpty(text))
                return true;
                
            if (text.EndsWith("\\"))
                return true;
                
            if (text.Contains("\\") && !text.Equals("My Computer", StringComparison.OrdinalIgnoreCase) 
                && !text.Equals("Favorites", StringComparison.OrdinalIgnoreCase))
                return true;
                
            return false;
        }

        private async void ShowSuggestions(string currentText)
        {
            if (suggestionListBox == null || addressTextBox == null || parentForm == null)
                return;

            // If we're currently navigating suggestions, don't interfere with the existing list
            if (isNavigatingSuggestions && isShowing)
            {
                Debug.WriteLine($"Ignoring ShowSuggestions call during navigation: '{currentText}'");
                return;
            }

            // Avoid showing suggestions for the same path repeatedly
            if (currentText == lastSuggestionPath && isShowing)
                return;

            try
            {
                var suggestions = await GetSuggestionsAsync(currentText);
                
                if (suggestions.Count > 0)
                {
                    // Store the original text when first showing suggestions
                    if (!isShowing)
                    {
                        originalAddressText = currentText;
                        lastSuggestionPath = currentText; // Only update this for new suggestion sessions
                        isNavigatingSuggestions = false;
                    }
                    
                    // Clear and populate suggestions
                    suggestionListBox.Items.Clear();
                    foreach (var suggestion in suggestions)
                    {
                        suggestionListBox.Items.Add(suggestion);
                    }
                    
                    // Position the suggestion box
                    PositionSuggestionBox();
                    
                    // Make sure the suggestion box is on top
                    suggestionListBox.BringToFront();
                    
                    // Show the suggestions
                    suggestionListBox.Visible = true;
                    isShowing = true;
                    
                    Debug.WriteLine($"Showing {suggestions.Count} suggestions for: '{currentText}' (Original: '{originalAddressText}')");
                }
                else
                {
                    // Only hide suggestions if we're not currently navigating through them
                    if (!isNavigatingSuggestions)
                    {
                        HideSuggestions();
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting suggestions: {ex.Message}");
                if (!isNavigatingSuggestions)
                {
                    HideSuggestions();
                }
            }
        }

        private async Task<List<string>> GetSuggestionsAsync(string currentText)
        {
            return await Task.Run(() => GetSuggestions(currentText));
        }

        private List<string> GetSuggestions(string currentText)
        {
            var suggestions = new List<string>();
            
            try
            {
                if (string.IsNullOrEmpty(currentText))
                {
                    // Show drives when empty
                    foreach (var drive in DriveInfo.GetDrives())
                    {
                        if (drive.IsReady)
                        {
                            var driveInfo = ExplorerUtils.GetDriveInfo(drive);
                            suggestions.Add($"{drive.Name} ({driveInfo.label})");
                        }
                    }
                }
                else if (currentText.EndsWith("\\"))
                {
                    // Show subdirectories when path ends with backslash
                    string basePath = currentText.TrimEnd('\\');
                    if (Directory.Exists(basePath))
                    {
                        var directories = Directory.GetDirectories(basePath)
                            .Take(10) // Limit to 10 suggestions for performance
                            .Select(dir => Path.GetFileName(dir))
                            .Where(name => !string.IsNullOrEmpty(name))
                            .OrderBy(name => name);
                            
                        foreach (var dir in directories)
                        {
                            suggestions.Add(Path.Combine(basePath, dir) + "\\");
                        }
                    }
                }
                else if (currentText.Contains("\\"))
                {
                    // Partial path - show matching directories
                    int lastBackslashIndex = currentText.LastIndexOf('\\');
                    if (lastBackslashIndex > 0)
                    {
                        string basePath = currentText.Substring(0, lastBackslashIndex);
                        string partialName = currentText.Substring(lastBackslashIndex + 1);
                        
                        if (Directory.Exists(basePath))
                        {
                            var matchingDirs = Directory.GetDirectories(basePath)
                                .Where(dir => Path.GetFileName(dir).StartsWith(partialName, StringComparison.OrdinalIgnoreCase))
                                .Take(10)
                                .Select(dir => dir + "\\")
                                .OrderBy(path => path);
                                
                            suggestions.AddRange(matchingDirs);
                        }
                    }
                }
                else
                {
                    // Root level - show drives that match
                    foreach (var drive in DriveInfo.GetDrives())
                    {
                        if (drive.IsReady && drive.Name.StartsWith(currentText, StringComparison.OrdinalIgnoreCase))
                        {
                            var driveInfo = ExplorerUtils.GetDriveInfo(drive);
                            suggestions.Add($"{drive.Name} ({driveInfo.label})");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting path suggestions: {ex.Message}");
            }
            
            return suggestions;
        }

        private void PositionSuggestionBox()
        {
            if (suggestionListBox == null || addressTextBox == null || parentForm == null)
                return;
                
            try
            {
                // Get more accurate positioning by walking up the control hierarchy
                Point addressBarLocation = GetAbsoluteLocation(addressTextBox);
                var textBoxBounds = addressTextBox.Bounds;
                
                // Calculate position below the address bar with proper spacing
                int x = addressBarLocation.X;
                int y = addressBarLocation.Y + textBoxBounds.Height + 3; // 3px gap below address bar
                
                // Adjust width to match address textbox width
                int width = Math.Max(textBoxBounds.Width, 200);
                
                // Calculate height based on items (max 8 items visible)
                int itemCount = Math.Min(suggestionListBox.Items.Count, 8);
                int height = Math.Max(itemCount * suggestionListBox.ItemHeight + 4, 50);
                
                // Make sure the suggestion box doesn't go off screen horizontally
                if (x + width > parentForm.ClientSize.Width)
                {
                    x = parentForm.ClientSize.Width - width - 10;
                }
                
                // Make sure we don't go off screen horizontally on the left
                if (x < 0)
                {
                    x = 5;
                }
                
                // Check if there's enough space below the address bar
                if (y + height > parentForm.ClientSize.Height)
                {
                    // Not enough space below, try to show above the address bar
                    int yAbove = addressBarLocation.Y - height - 3;
                    if (yAbove >= 0)
                    {
                        y = yAbove;
                    }
                    else
                    {
                        // Not enough space above either, limit height and show below
                        int availableHeight = parentForm.ClientSize.Height - y - 10;
                        if (availableHeight > 50) // Minimum useful height
                        {
                            height = availableHeight;
                        }
                        else
                        {
                            // As last resort, position at bottom of form with limited height
                            height = 100;
                            y = parentForm.ClientSize.Height - height - 5;
                        }
                    }
                }
                
                suggestionListBox.Location = new Point(x, y);
                suggestionListBox.Size = new Size(width, height);
                
                Debug.WriteLine($"Suggestion box positioned at: {x}, {y} with size: {width}x{height}");
                Debug.WriteLine($"Address bar location: {addressBarLocation.X}, {addressBarLocation.Y}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error positioning suggestion box: {ex.Message}");
            }
        }

        private Point GetAbsoluteLocation(ToolStripTextBox textBox)
        {
            // Get the absolute location of the textbox within the form
            Point location = Point.Empty;
            
            try
            {
                // Start with the textbox bounds within its owner (ToolStrip)
                location = textBox.Bounds.Location;
                
                // Add the ToolStrip's location
                Control? owner = textBox.Owner;
                while (owner != null && owner != parentForm)
                {
                    location.X += owner.Location.X;
                    location.Y += owner.Location.Y;
                    
                    // Move up to the parent control
                    owner = owner.Parent;
                }
                
                Debug.WriteLine($"Calculated absolute location: {location.X}, {location.Y}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error calculating absolute location: {ex.Message}");
                // Fallback to a safe default position
                location = new Point(100, 100);
            }
            
            return location;
        }

        private void AddressTextBox_KeyDown(object? sender, KeyEventArgs e)
        {
            if (suggestionListBox == null || !isShowing)
                return;
                
            switch (e.KeyCode)
            {
                case Keys.Down:
                    // Navigate down through suggestions
                    if (suggestionListBox.Items.Count > 0)
                    {
                        // Set flag before making any changes
                        isNavigatingSuggestions = true;
                        
                        if (suggestionListBox.SelectedIndex < 0)
                        {
                            // First time pressing down - select first item
                            suggestionListBox.SelectedIndex = 0;
                        }
                        else if (suggestionListBox.SelectedIndex < suggestionListBox.Items.Count - 1)
                        {
                            // Move to next item
                            suggestionListBox.SelectedIndex++;
                        }
                        
                        // UpdateAddressBarWithSuggestion will be called by SelectedIndexChanged
                        e.Handled = true;
                    }
                    break;
                    
                case Keys.Up:
                    // Navigate up through suggestions
                    if (suggestionListBox.Items.Count > 0)
                    {
                        // Set flag before making any changes
                        isNavigatingSuggestions = true;
                        
                        if (suggestionListBox.SelectedIndex <= 0)
                        {
                            // At first item or no selection - restore original text
                            RestoreOriginalText();
                            suggestionListBox.SelectedIndex = -1;
                        }
                        else
                        {
                            // Move to previous item
                            suggestionListBox.SelectedIndex--;
                            // UpdateAddressBarWithSuggestion will be called by SelectedIndexChanged
                        }
                        
                        e.Handled = true;
                    }
                    break;
                    
                case Keys.Enter:
                    // Accept current suggestion (or original text if no selection)
                    if (suggestionListBox.SelectedIndex >= 0 && suggestionListBox.SelectedItem != null)
                    {
                        AcceptSuggestion(suggestionListBox.SelectedItem.ToString() ?? "");
                    }
                    else if (!string.IsNullOrEmpty(originalAddressText))
                    {
                        // No suggestion selected, use original text
                        isNavigatingProgrammatically = true;
                        addressTextBox.Text = originalAddressText;
                        isNavigatingProgrammatically = false;
                    }
                    HideSuggestions();
                    e.Handled = true;
                    break;
                    
                case Keys.Escape:
                    // Cancel suggestions and restore original text
                    RestoreOriginalText();
                    HideSuggestions();
                    e.Handled = true;
                    break;
                    
                case Keys.Tab:
                    // Accept first suggestion
                    if (suggestionListBox.Items.Count > 0)
                    {
                        suggestionListBox.SelectedIndex = 0;
                        AcceptSuggestion(suggestionListBox.Items[0].ToString() ?? "");
                        e.Handled = true;
                    }
                    break;
            }
        }

        private void SuggestionListBox_Click(object? sender, EventArgs e)
        {
            if (suggestionListBox?.SelectedItem != null)
            {
                // Update address bar with the clicked suggestion first
                UpdateAddressBarWithSuggestion();
                
                // Then accept the suggestion
                AcceptSuggestion(suggestionListBox.SelectedItem.ToString() ?? "");
            }
        }

        private void SuggestionListBox_KeyDown(object? sender, KeyEventArgs e)
        {
            if (suggestionListBox == null)
                return;
                
            switch (e.KeyCode)
            {
                case Keys.Up:
                    if (suggestionListBox.SelectedIndex <= 0)
                    {
                        // At first item - restore original text and return focus to address bar
                        RestoreOriginalText();
                        suggestionListBox.SelectedIndex = -1;
                        addressTextBox?.Focus();
                    }
                    else
                    {
                        // Set flag before making changes
                        isNavigatingSuggestions = true;
                        // Move to previous item
                        suggestionListBox.SelectedIndex--;
                        // UpdateAddressBarWithSuggestion will be called by SelectedIndexChanged
                    }
                    e.Handled = true;
                    break;
                    
                case Keys.Down:
                    if (suggestionListBox.SelectedIndex < suggestionListBox.Items.Count - 1)
                    {
                        // Set flag before making changes
                        isNavigatingSuggestions = true;
                        // Move to next item
                        suggestionListBox.SelectedIndex++;
                        // UpdateAddressBarWithSuggestion will be called by SelectedIndexChanged
                    }
                    e.Handled = true;
                    break;
                    
                case Keys.Enter:
                    if (suggestionListBox.SelectedItem != null)
                    {
                        AcceptSuggestion(suggestionListBox.SelectedItem.ToString() ?? "");
                        e.Handled = true;
                    }
                    break;
                    
                case Keys.Escape:
                    RestoreOriginalText();
                    HideSuggestions();
                    addressTextBox?.Focus();
                    e.Handled = true;
                    break;
                    
                case Keys.Back:
                case Keys.Delete:
                    // Return focus to address bar for editing
                    RestoreOriginalText();
                    addressTextBox?.Focus();
                    e.Handled = true;
                    break;
            }
        }

        private void AcceptSuggestion(string suggestion)
        {
            if (addressTextBox == null)
                return;
                
            try
            {
                // Extract the actual path from suggestions that include drive labels
                string path = suggestion;
                if (suggestion.Contains(" (") && suggestion.EndsWith(")"))
                {
                    // This is a drive suggestion like "C:\ (Local Disk)"
                    path = suggestion.Substring(0, suggestion.IndexOf(" ("));
                }
                
                // End suggestion navigation before setting final text
                EndSuggestionNavigation();
                
                isNavigatingProgrammatically = true;
                addressTextBox.Text = path;
                isNavigatingProgrammatically = false;
                
                // Position cursor at the end
                if (addressTextBox.Control is TextBox textBox)
                {
                    textBox.SelectionStart = textBox.Text.Length;
                    textBox.SelectionLength = 0;
                }
                
                HideSuggestions();
                addressTextBox.Focus();
                
                Debug.WriteLine($"Accepted suggestion: '{path}'");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error accepting suggestion: {ex.Message}");
            }
        }

        private void AddressTextBox_LostFocus(object? sender, EventArgs e)
        {
            // Use a small delay to allow clicking on suggestions
            if (parentForm != null)
            {
                var timer = new System.Windows.Forms.Timer { Interval = 150 };
                timer.Tick += (s, args) =>
                {
                    timer.Stop();
                    timer.Dispose();
                    if (!suggestionListBox?.Focused == true)
                    {
                        HideSuggestions();
                    }
                };
                timer.Start();
            }
        }

        public void HideSuggestions()
        {
            if (suggestionListBox != null)
            {
                suggestionListBox.Visible = false;
                isShowing = false;
                
                // Only clear tracking if we're not in the middle of navigation
                if (!isNavigatingSuggestions)
                {
                    lastSuggestionPath = "";
                    originalAddressText = "";
                }
                
                // Always reset the navigation flag when hiding
                isNavigatingSuggestions = false;
                
                Debug.WriteLine($"Hidden suggestions. Navigation flag reset.");
            }
        }

        public void SetNavigatingProgrammatically(bool isNavigating)
        {
            isNavigatingProgrammatically = isNavigating;
            if (isNavigating)
            {
                HideSuggestions();
            }
        }

        public void Dispose()
        {
            if (suggestionListBox != null)
            {
                suggestionListBox.Dispose();
                suggestionListBox = null;
            }
        }

        private void UpdateAddressBarWithSuggestion()
        {
            if (suggestionListBox?.SelectedItem != null && addressTextBox != null)
            {
                string suggestion = suggestionListBox.SelectedItem.ToString() ?? "";
                
                // Extract the actual path from suggestions that include drive labels
                string path = suggestion;
                if (suggestion.Contains(" (") && suggestion.EndsWith(")"))
                {
                    // This is a drive suggestion like "C:\ (Local Disk)"
                    path = suggestion.Substring(0, suggestion.IndexOf(" ("));
                }
                
                // Temporarily disable text change events to prevent recursive calls
                bool wasNavigatingProgrammatically = isNavigatingProgrammatically;
                isNavigatingProgrammatically = true;
                
                addressTextBox.Text = path;
                
                // Restore the previous programmatic navigation state
                isNavigatingProgrammatically = wasNavigatingProgrammatically;
                
                // Position cursor at the end
                if (addressTextBox.Control is TextBox textBox)
                {
                    textBox.SelectionStart = textBox.Text.Length;
                    textBox.SelectionLength = 0;
                }
                
                Debug.WriteLine($"Updated address bar with suggestion: '{path}'");
            }
        }

        private void RestoreOriginalText()
        {
            if (addressTextBox != null && !string.IsNullOrEmpty(originalAddressText))
            {
                // End suggestion navigation before restoring text
                EndSuggestionNavigation();
                
                // Temporarily disable text change events to prevent recursive calls
                isNavigatingProgrammatically = true;
                addressTextBox.Text = originalAddressText;
                isNavigatingProgrammatically = false;
                
                // Position cursor at the end
                if (addressTextBox.Control is TextBox textBox)
                {
                    textBox.SelectionStart = textBox.Text.Length;
                    textBox.SelectionLength = 0;
                }
                
                Debug.WriteLine($"Restored original text: '{originalAddressText}'");
            }
        }

        private void SuggestionListBox_SelectedIndexChanged(object? sender, EventArgs e)
        {
            // Update address bar when selection changes (including mouse hover)
            if (isShowing && suggestionListBox?.SelectedIndex >= 0)
            {
                // Set the flag BEFORE updating the address bar to prevent interference
                isNavigatingSuggestions = true;
                UpdateAddressBarWithSuggestion();
            }
        }

        private void EndSuggestionNavigation()
        {
            isNavigatingSuggestions = false;
            // Reset the last suggestion path to allow fresh suggestions
            lastSuggestionPath = "";
            Debug.WriteLine("Ended suggestion navigation - ready for new suggestions");
        }
    }
}