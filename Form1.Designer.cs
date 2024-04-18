using Microsoft.VisualBasic.Devices;
using System.Windows.Forms;
using System.Windows.Input;
using Newtonsoft.Json;

namespace Macro
{
    public partial class Form1 : Form
    {
        private string[] row1 = { "`", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "-", "=", "Backspace" };
        private string[] row2 = { "Tab", "q", "w", "e", "r", "t", "y", "u", "i", "o", "p", "[", "]" };
        private string[] row3 = { "Caps Lock", "a", "s", "d", "f", "g", "h", "j", "k", "l", ";", "'", "Enter"};
        private string[] row4 = { "Shift", "z", "x", "c", "v", "b", "n", "m", ",", ".", "/", "Shift"};
        private string[] row5 = { "Ctrl", "Alt", "Space", "Alt", "Ctrl", "Left Arrow", "Up Arrow", "Down Arrow", "Right Arrow" };

        private string[] row1Caps = { "~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "+", "Backspace" };
        private string[] row2Caps = { "Tab", "Q", "W", "E", "R", "T", "Y", "U", "I", "O", "P", "{", "}"};
        private string[] row3Caps = { "Caps Lock", "A", "S", "D", "F", "G", "H", "J", "K", "L", ":", "\"", "Enter"};
        private string[] row4Caps = { "Shift", "Z", "X", "C", "V", "B", "N", "M", "<", ">", "?", "Shift"};

        private bool capsLockEnabled = false;
        private string keyholder;
        private bool isMacroKeyPressed = false;

        protected List<string> macros = new List<string>(); // List to store macros
        private Dictionary<string, bool> macroStates = new Dictionary<string, bool>();
        private System.Windows.Forms.Timer macroExecutionTimer;
        private System.Windows.Forms.Timer macroExecutionTimer2;
        private System.ComponentModel.IContainer components = null;

        public void InitializeForm()
        {
            LoadDataFromFile("data.json");
            InitializeComponent(1100, 350);
            InitializeKeyboard();
            InitializeMacroList();
            this.KeyDown += Form1_KeyDown;
            macroExecutionTimer = new System.Windows.Forms.Timer();
            macroExecutionTimer.Interval = 1000; // Check for macro updates every second
            macroExecutionTimer.Tick += (sender, e) =>
            {
                RunEnabledMacros();
            };
            macroExecutionTimer.Start();
            macroExecutionTimer2 = new System.Windows.Forms.Timer();
            macroExecutionTimer2.Interval = 10000;
            macroExecutionTimer2.Tick += (sender, e) =>
            {
                SaveDataToFile("data.json");
            };
            macroExecutionTimer2.Start();
        }

        private void SaveDataToFile(string filePath)
        {
            try
            {
                // Create a data object to hold macros and macro states
                var data = new
                {
                    Macros = macros,
                    MacroStates = macroStates
                };

                // Serialize the data to JSON format
                string jsonData = JsonConvert.SerializeObject(data, Formatting.Indented);

                // Write the JSON data to the file
                File.WriteAllText(filePath, jsonData);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving data to file: {ex.Message}");
            }
        }

        // Load macros and macro states from a JSON file
        private void LoadDataFromFile(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    // Read the JSON data from the file
                    string jsonData = File.ReadAllText(filePath);

                    // Deserialize the JSON data into a dynamic object
                    dynamic data = JsonConvert.DeserializeObject(jsonData);

                    // Update macros list
                    macros = data.Macros.ToObject<List<string>>();

                    // Update macro states dictionary
                    macroStates = data.MacroStates.ToObject<Dictionary<string, bool>>();

                    // Update the macro list on the form
                    UpdateMacroList();
                }
                else
                {

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading data from file: {ex.Message}");
            }
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }
        public void InitializeComponent(int width, int height)
        {
            SuspendLayout();
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(width, height);
            Text = $"Macro Machine";
            Load += Form1_Load;
            ResumeLayout(false);
        }
        private void InitializeKeyboard()
        {

            string[][] rows = { row1, row2, row3, row4, row5 };
            string[][] rowsCaps = { row1Caps, row2Caps, row3Caps, row4Caps, row5 };

            int buttonWidth = 50;
            int buttonHeight = 50;
            int panelWidth = 800; // Width of the panel to contain the buttons
            int panelHeight = 50; // Height of each panel
            int startY = 5;
            for (int row = 0; row < rows.Length; row++)
            {
                // Clear existing controls
                foreach (Control control in Controls)
                {
                    if (control is Panel)
                    {
                        Controls.Remove(control);
                        control.Dispose();
                    }
                }
            }

            for (int row = 0; row < rows.Length; row++)
            {

                string[] currentRow;
                if (capsLockEnabled && row < 4) // Check if caps lock is enabled and it's one of the first 4 rows
                {
                    currentRow = rowsCaps[row];
                }
                else
                {
                    currentRow = rows[row];
                }

                // Additional spacing for row 4 buttons
                int extraSpacing = 5;

                // Calculate the total width occupied by buttons in the current row
                int totalButtonWidth = currentRow.Length * (buttonWidth + extraSpacing) - extraSpacing;
                int startX = (panelWidth - totalButtonWidth) / 2; // Calculate the left position to center buttons horizontally

                // Calculate starting X position for the bottom row buttons separately to ensure centering
                if (row == 4)
                {
                    extraSpacing = 23;
                    startX = (panelWidth - (buttonWidth * currentRow.Length + extraSpacing * (currentRow.Length - 1))) / 2;
                }

                // Create a Panel to contain the buttons in each row
                Panel panel = new Panel();
                panel.Width = panelWidth;
                panel.Height = panelHeight;
                panel.Left = 10; // Anchor to the left side of the form
                panel.Top = startY + row * (buttonHeight + 10); // Adjust the top position
                panel.BackColor = Color.Transparent; // Make the panel background transparent

                for (int col = 0; col < currentRow.Length; col++)
                {
                    Button button = new Button();
                    button.Width = buttonWidth;
                    button.Height = buttonHeight;

                    // Adjust the left position to center buttons horizontally with additional spacing for row 4 buttons
                    button.Left = startX + col * (buttonWidth + extraSpacing);
                    button.Top = 0; // Align buttons to the top of the panel

                    // Increase width for buttons in the bottom row
                    if (row == 4)
                    {
                        button.Width = buttonWidth + 20;
                    }

                    button.Text = currentRow[col];
                    button.Click += Button_Click;
                    panel.Controls.Add(button);
                }

                this.Controls.Add(panel); // Add the panel to the form
            }

            // Add button to switch between lowercase and uppercase
            Button btnCapsLock = new Button();
            btnCapsLock.Text = "Toggle Keyboard Capitalization";
            btnCapsLock.Location = new Point(350, 300);
            btnCapsLock.AutoSize = true;
            btnCapsLock.Click += BtnCapsLock_Click;
            this.Controls.Add(btnCapsLock);
        }

        private void BtnCapsLock_Click(object sender, EventArgs e)
        {
            capsLockEnabled = !capsLockEnabled; // Toggle caps lock
            InitializeKeyboard(); // Reinitialize keyboard layout to update the case of letters
        }

        private ListBox macroListBox;
        private void InitializeMacroList()
        {
            macroListBox = new ListBox();
            macroListBox.Name = "macroListBox"; // Set the name to find it later
            macroListBox.Width = 250;
            macroListBox.Left = this.Width - macroListBox.Width - 10;
            macroListBox.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;

            // Set the top position to align with the top of the form
            macroListBox.Top = 0;
            // Set the height to span from the top to the bottom of the form
            macroListBox.Height = this.ClientSize.Height;

            macroListBox.DataSource = macros;
            macroListBox.ScrollAlwaysVisible = true;

            macroListBox.DrawMode = DrawMode.OwnerDrawFixed; // Set draw mode to custom

            // Subscribe to the DrawItem event
            macroListBox.DrawItem += MacroListBox_DrawItem;

            // Subscribe to the MouseClick event
            macroListBox.MouseClick += MacroListBox_MouseClick;

            this.Controls.Add(macroListBox);
        }

        private void MacroListBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();
            e.DrawFocusRectangle();

            if (e.Index >= 0 && e.Index < macroListBox.Items.Count)
            {
                // Retrieve the macro text
                string macroText = macroListBox.Items[e.Index].ToString();

                // Calculate the bounds for the checkbox
                Rectangle checkBoxRect = new Rectangle(e.Bounds.Left + 5, e.Bounds.Top + 3, 15, e.Bounds.Height - 6);

                // Draw the checkbox
                CheckBox checkBox = new CheckBox();
                checkBox.Location = checkBoxRect.Location;
                checkBox.Size = checkBoxRect.Size;

                // Set the checkbox state based on the corresponding state in the macroStates dictionary
                string macro = macroText;
                if (macroStates.ContainsKey(macro))
                {
                    checkBox.Checked = macroStates[macro];
                }
                else
                {
                    checkBox.Checked = false; // Default state
                }

                // Handle the checkbox state change
                checkBox.CheckedChanged += (s, args) =>
                {
                    // Update the corresponding state in the macroStates dictionary
                    macroStates[macro] = checkBox.Checked;
                };

                // Calculate the bounds for the "X" button
                Rectangle closeButtonRect = new Rectangle(e.Bounds.Right - 20, e.Bounds.Top, 20, e.Bounds.Height);

                // Draw the "X" button
                e.Graphics.DrawString("X", e.Font, Brushes.Red, closeButtonRect, StringFormat.GenericDefault);

                // Draw the macro text with some offset to avoid overlapping with the checkbox
                int textOffset = 25;
                Rectangle macroTextRect = new Rectangle(e.Bounds.Left + textOffset, e.Bounds.Top, e.Bounds.Width - textOffset - 20, e.Bounds.Height);
                e.Graphics.DrawString(macroText, e.Font, Brushes.Black, macroTextRect);

                // Add the checkbox to the ListBox
                macroListBox.Controls.Add(checkBox);
            }
        }




        private void MacroListBox_MouseClick(object sender, MouseEventArgs e)
        {
            // Check if the click is within the bounds of the "X" button
            for (int i = 0; i < macroListBox.Items.Count; i++)
            {
                Rectangle rect = new Rectangle(macroListBox.GetItemRectangle(i).Right - 20, macroListBox.GetItemRectangle(i).Top, 20, macroListBox.GetItemRectangle(i).Height);
                if (rect.Contains(e.Location))
                {
                    // Remove the corresponding macro from the list
                    string macro = macros[i];
                    macros.RemoveAt(i);

                    // Remove the corresponding macro state
                    if (macroStates.ContainsKey(macro))
                    {
                        macroStates.Remove(macro);
                    }

                    // Remove the corresponding checkbox from the ListBox controls
                    if (macroListBox.Controls.Count > i)
                    {
                        Control checkBox = macroListBox.Controls[i];
                        macroListBox.Controls.Remove(checkBox);
                        checkBox.Dispose(); // Dispose the checkbox to release resources
                    }

                    // Update the ListBox
                    macroListBox.DataSource = null;
                    macroListBox.DataSource = macros;
                    Controls.Remove(macroListBox);
                    InitializeMacroList();
                    break;
                }
            }
        }



        private void Button_Click(object sender, EventArgs e)
        {
            Button button = sender as Button;
            string key = button.Text;

            // Open a new form for creating a macro for the clicked key
            using (MacroCreationForm macroForm = new MacroCreationForm(key))
            {
                DialogResult result = macroForm.ShowDialog();

                // Check if the macro form was closed with OK result
                if (result == DialogResult.OK)
                {
                    // Handle the macro creation process
                    string macroName = macroForm.MacroName;
                    string macroContent = macroForm.MacroContent;
                    string triggerKey = macroForm.TriggerKey;
                    double cooldownSeconds = macroForm.CooldownSeconds;

                    // Construct the macro string based on the user's choices
                    string macro;
                    if (!string.IsNullOrEmpty(triggerKey))
                    {
                        macro = $"{key}: {macroName} Trigger: {triggerKey}";
                    }
                    else
                    {
                        macro = $"{key}: {macroName} (Cooldown: {cooldownSeconds} seconds)";
                    }

                    // Add the macro to the list
                    macros.Add(macro);

                    // Update the macro list
                    UpdateMacroList();

                    // Add or update the macro state in the dictionary
                    if (!macroStates.ContainsKey(macro))
                    {
                        macroStates.Add(macro, false); // Add with default state of false
                    }
                    else
                    {
                        macroStates[macro] = false; // Update the state to false
                    }
                }
            }
        }


        protected void UpdateMacroList()
        {
            if (macroListBox != null)
            {
                macroListBox.DataSource = null;
                macroListBox.DataSource = macros;
            }
        }
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            Enum.TryParse(keyholder, out Keys parsedEnumValue);
            if (e.KeyCode == parsedEnumValue)
            {
                isMacroKeyPressed = true;
            }
        }
        private async Task ExecuteMacro(string macro)
        {
            // Extract the key and macro content from the macro string
            string[] parts = macro.Split(':');
            if (parts.Length != 3)
            {
                // Invalid macro format, ignore
                return;
            }

            string key = parts[0].Trim();
            string macroContent = parts[1].Trim();
            string macroEnd = parts[2].Trim();

            // Check if the macro is enabled
            if (!macroStates.TryGetValue(macro, out bool isEnabled) || !isEnabled)
            {
                // Macro is disabled, ignore
                return;
            }

            // Execute the macro based on its settings
            if (macroContent.Contains("Cooldown"))
            {
                // Extract the cooldown value from the macro content
                double cooldownSeconds = GetCooldownFromMacroContent(macroEnd);

                // Continuously send the key while the macro is enabled
                while (macroStates[macro])
                {
                    SendKey(key);
                    await Task.Delay((int)(cooldownSeconds * 1000)); // Delay for cooldown duration
                }
            }
            else if (macroContent.Contains("Trigger"))
            {
                // Extract the trigger key from the macro content
                keyholder = macroEnd;
                // Wait for the trigger key to be pressed
                while (!isMacroKeyPressed)
                {
                    await Task.Delay(100); // Polling interval of 100 milliseconds
                }

                // Trigger key pressed, send the key
                SendKey(key);
            }
            else
            {
                // Invalid macro content, ignore
                return;
            }
        }

        private void SendKey(string key)
        {
            SendKeys.Send(key);
        }


        private double GetCooldownFromMacroContent(string macroContent)
        {
            // Extract the cooldown value from the macro content
            string[] parts = macroContent.Split(' ');
            if (parts.Length != 2)
            {
                // Invalid macro content format, return default cooldown
                return 1.0; // Default cooldown of 1 second
            }

            string cooldownStr = parts[0].Trim();
            if (double.TryParse(cooldownStr, out double cooldownSeconds))
            {
                return cooldownSeconds;
            }
            else
            {
                // Invalid cooldown value, return default
                return 1.0; // Default cooldown of 1 second
            }
        }

        private async void RunEnabledMacros()
        {
            // Run all enabled macros in separate tasks
            List<Task> tasks = new List<Task>();
            foreach (string macro in macroStates.Keys)
            {
                if (macroStates[macro])
                {
                    tasks.Add(ExecuteMacro(macro));
                }
            }
            await Task.WhenAll(tasks);
        }

        private void macroListBox_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            // Update the state of the macro when the checkbox state changes
            string macro = macroListBox.Items[e.Index].ToString();
            if (e.NewValue == CheckState.Checked)
            {
                // Enable the macro
                macroStates[macro] = true;
            }
            else
            {
                // Disable the macro
                macroStates[macro] = false;
            }
        }
}

    public partial class MacroCreationForm : Form
    {
        // Properties for macro name, content, and behavior
        public string MacroName { get; private set; }
        public string MacroContent { get; private set; }
        public bool PressRepeatedly { get; private set; } // Indicates if the button should be pressed repeatedly
        public bool TriggerOnKeyPress { get; private set; } // Indicates if the macro should trigger on another key press
        public string TriggerKey { get; private set; } // Trigger key if Trigger on key press is selected
        public double CooldownSeconds { get; private set; } // Cooldown time in seconds

        // Controls
        private TextBox textBoxName;
        private CheckBox checkBoxPressRepeatedly;
        private CheckBox checkBoxTriggerOnKeyPress;
        private TextBox textBoxTriggerKey;
        private NumericUpDown numericUpDownCooldown;

        // Constructor
        public MacroCreationForm(string key)
        {
            Text = $"Create Macro for '{key}'";
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(350, 200);
            InitializeControls();
        }

        // Method to initialize controls
        private void InitializeControls()
        {
            // Add label for macro name
            Label labelMacroName = new Label();
            labelMacroName.Text = "Macro Name:";
            labelMacroName.Location = new Point(10, 20);
            labelMacroName.AutoSize = true;
            Controls.Add(labelMacroName);

            // Add TextBox for macro name
            textBoxName = new TextBox();
            textBoxName.Location = new Point(110, 20); // Adjust the X position to align with the label
            textBoxName.Size = new Size(200, 20);
            Controls.Add(textBoxName);

            // Add CheckBoxes for mutually exclusive options
            checkBoxTriggerOnKeyPress = new CheckBox();
            checkBoxTriggerOnKeyPress.Text = "Trigger on key press:";
            checkBoxTriggerOnKeyPress.Location = new Point(20, 50);
            checkBoxTriggerOnKeyPress.AutoSize = true;
            checkBoxTriggerOnKeyPress.CheckedChanged += (sender, e) =>
            {
                TriggerOnKeyPress = checkBoxTriggerOnKeyPress.Checked;
                if (TriggerOnKeyPress)
                {
                    checkBoxPressRepeatedly.Checked = false; // Ensure only one option is selected
                    textBoxTriggerKey.Enabled = true; // Enable the TextBox
                    numericUpDownCooldown.Enabled = false; // Disable the NumericUpDown
                }
                else
                {
                    textBoxTriggerKey.Enabled = false; // Disable the TextBox
                    if (!checkBoxPressRepeatedly.Checked)
                    {
                        numericUpDownCooldown.Enabled = true;
                    }
                }
            };
            Controls.Add(checkBoxTriggerOnKeyPress);

            checkBoxPressRepeatedly = new CheckBox();
            checkBoxPressRepeatedly.Text = "Press repeatedly:";
            checkBoxPressRepeatedly.AutoSize = true;
            checkBoxPressRepeatedly.Location = new Point(20, 80);
            checkBoxPressRepeatedly.CheckedChanged += (sender, e) =>
            {
                PressRepeatedly = checkBoxPressRepeatedly.Checked;
                if (PressRepeatedly)
                {
                    checkBoxTriggerOnKeyPress.Checked = false; // Ensure only one option is selected
                    textBoxTriggerKey.Enabled = false; // Disable the TextBox
                    numericUpDownCooldown.Enabled = true; // Enable the NumericUpDown
                }
                else
                {
                    numericUpDownCooldown.Enabled = false;
                }
            };
            Controls.Add(checkBoxPressRepeatedly);

            // Add TextBox for trigger key (enabled when "Trigger on key press" is checked)
            textBoxTriggerKey = new TextBox();
            textBoxTriggerKey.Location = new Point(160, 50); // Adjust the X position to align with the label
            textBoxTriggerKey.Size = new Size(40, 20);
            textBoxTriggerKey.Enabled = false; // Initially disabled
            Controls.Add(textBoxTriggerKey);

            // Add NumericUpDown for cooldown time
            Label labelCooldown = new Label();
            labelCooldown.AutoSize = true;
            labelCooldown.Text = "Cooldown (seconds):";
            labelCooldown.Location = new Point(20, 110);
            Controls.Add(labelCooldown);

            numericUpDownCooldown = new NumericUpDown();
            numericUpDownCooldown.Location = new Point(160, 108);
            numericUpDownCooldown.Size = new Size(70, 20);
            numericUpDownCooldown.Minimum = 0;
            numericUpDownCooldown.Maximum = 1000; // Maximum cooldown time in seconds
            numericUpDownCooldown.DecimalPlaces = 1; // Allow decimal values (e.g., 0.1 seconds)
            numericUpDownCooldown.Enabled = true; // Initially enabled
            Controls.Add(numericUpDownCooldown);

            // Add Save and Cancel buttons
            Button btnSave = new Button();
            btnSave.Text = "Save";
            btnSave.Location = new Point(20, 150);
            btnSave.Click += btnSave_Click;
            Controls.Add(btnSave);

            Button btnCancel = new Button();
            btnCancel.Text = "Cancel";
            btnCancel.Location = new Point(100, 150);
            btnCancel.Click += btnCancel_Click;
            Controls.Add(btnCancel);
        }

        // Save button click event handler
        private void btnSave_Click(object sender, EventArgs e)
        {
            // Check if "Trigger on key press" is selected and TextBox is empty
            if (checkBoxTriggerOnKeyPress.Checked && string.IsNullOrWhiteSpace(textBoxTriggerKey.Text))
            {
                MessageBox.Show("Please enter a key for the trigger key.");
                return; // Exit the method without saving
            }

            // Check if "Press repeatedly" is selected and NumericUpDown value is zero
            if (checkBoxPressRepeatedly.Checked && numericUpDownCooldown.Value == 0)
            {
                MessageBox.Show("Please enter a non-zero value for cooldown.");
                return; // Exit the method without saving
            }

            // Retrieve data from controls
            MacroName = textBoxName.Text;
            TriggerKey = textBoxTriggerKey.Text;
            CooldownSeconds = (double)numericUpDownCooldown.Value;
            MacroContent = ""; // Placeholder for macro content

            // Close the form with OK result
            DialogResult = DialogResult.OK;
            Close();
        }

        // Cancel button click event handler
        private void btnCancel_Click(object sender, EventArgs e)
        {
            // Close the form with Cancel result
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
