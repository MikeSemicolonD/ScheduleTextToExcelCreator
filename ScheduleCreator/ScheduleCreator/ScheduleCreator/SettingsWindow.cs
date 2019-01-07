using System;
using System.Windows.Forms;

namespace ScheduleCreator
{
    public partial class SettingsWindow : Form
    {
        /// <summary>
        /// Constructor
        /// Sets the location to the main window's location
        /// Applys the settings that currently exist to the UI in the window
        /// </summary>
        public SettingsWindow()
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterScreen;
            ApplySettingsToUI(MainWindow.instance.RuntimeSettings);
        }

        /// <summary>
        /// Gets called when the window is being called.
        /// Allows to have only a single instance of this window.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SettingsWindow_FormClosing(Object sender, FormClosingEventArgs e)
        {
            MainWindow.instance.ClosedSettingsWindowCallback();
        }

        /// <summary>
        /// Apply button
        /// Sends changes to the main window and closes the settings window
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            MainWindow.instance.ApplyNewSettings(checkBox1.Checked,checkBox2.Checked,checkBox3.Checked,comboBox1.SelectedIndex);
            Close();
        }

        /// <summary>
        /// Set to Default button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            MainWindow.instance.ApplyDefaultSettings();
            
            ApplySettingsToUI(MainWindow.instance.Default);
        }

        /// <summary>
        /// Takes a new settings struct and applies it to the UI
        /// </summary>
        /// <param name="newSettings"></param>
        private void ApplySettingsToUI(MainWindow.Settings newSettings)
        {
            checkBox1.Checked = newSettings.OutputDataAsRawInExcel;
            checkBox2.Checked = newSettings.OutputCreditTotalInExcel;
            checkBox3.Checked = newSettings.SwitchDayAndMonthPositionInExcel;
            comboBox1.SelectedIndex = newSettings.SelectedParserTemplate;
        }
    }
}
