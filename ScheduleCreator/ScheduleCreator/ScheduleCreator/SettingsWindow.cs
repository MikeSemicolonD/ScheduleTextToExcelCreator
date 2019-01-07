using System;
using System.Windows.Forms;

namespace ScheduleCreator
{
    public partial class SettingsWindow : Form
    {

        public SettingsWindow()
        {
            InitializeComponent();
            ApplySettingsToUI(MainWindow.instance.RuntimeSettings);
        }

        /// <summary>
        /// Apply button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            MainWindow.instance.ApplyNewSettings(checkBox1.Checked,checkBox2.Checked,checkBox3.Checked,comboBox1.SelectedIndex);
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
