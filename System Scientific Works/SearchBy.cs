using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace System_Scientific_Works
{
    public partial class SearchBy : Form
    {
        
        Data q;
        DB db;
        int ind;
        
        public SearchBy(DB db, int indicator)
        {
            InitializeComponent();

            ind = indicator;
            this.db = db;
            q = new Data();
            
            switch (ind)
            {
                case 6:
                    q = db.GetNameId("Faculty");
                    numericUpDown1.Value = DateTime.Now.Year;
                    break;
                case 7:
                    numericUpDown1.Visible= false;
                    label2.Visible= false;
                    q = db.GetNameId("Faculty");
                    break;
                case 8:
                    label1.Text = "Выберите кафедру";
                    numericUpDown1.Visible = false;
                    label2.Visible = false;
                    q = db.GetNameId("Department");
                    break;
            }
            
            foreach(var testc in q.name) comboBox1.Items.Add(testc);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            switch (ind)
            {
                case 6:
                    if (comboBox1.SelectedIndex != -1)
                    {

                        Year = Convert.ToInt32(numericUpDown1.Value);
                        Faculty = q.id[comboBox1.SelectedIndex];
                        name = q.name[comboBox1.SelectedIndex];
                        
                        DialogResult = DialogResult.OK;
                    }
                    break;
                case 7:
                    if (comboBox1.SelectedIndex != -1)
                    {
                        Faculty = q.id[comboBox1.SelectedIndex];
                        name = q.name[comboBox1.SelectedIndex];
                        DialogResult = DialogResult.OK;
                    }
                    break;
                case 8:
                    if (comboBox1.SelectedIndex != -1)
                    {
                        Faculty = q.id[comboBox1.SelectedIndex];
                        name = q.name[comboBox1.SelectedIndex];

                        DialogResult = DialogResult.OK;
                    }
                    break;
            }
            

            
        }
        public int Faculty;
        public int Year;
        public string name;

    }
}
