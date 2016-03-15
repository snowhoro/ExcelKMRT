using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        string filename;
        public Form1()
        {
            InitializeComponent();
            monthComboBox.SelectedIndex = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;

            string _filepath = filename;
            int _monthNumber = monthComboBox.SelectedIndex;
            string _monthName = monthComboBox.SelectedItem.ToString();
            int _yearNumber = (int)yearNumber.Value;
            int _minValue = (int)numericUpDown1.Value;
            int _maxValue = (int)numericUpDown2.Value;
            bool _RunTotalKM = randomKM.Checked;
            int _mondayKM = (int)numericUpDown3.Value;
            Excel excel = new Excel(_filepath, _monthNumber, _monthName, _yearNumber,
                _minValue, _maxValue, _RunTotalKM, _mondayKM);
            excel.StartExcel();
            excel.CloseExcel();

            button1.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel|*.xls*;*.xlsx*|All Files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.Multiselect = false;

            DialogResult userClickedOK = openFileDialog1.ShowDialog();

            if (userClickedOK == DialogResult.OK)
            {
                filename = openFileDialog1.InitialDirectory + openFileDialog1.FileName;
                textBox1.Text = Path.GetFileName(filename);
            }
        }
    }
}
