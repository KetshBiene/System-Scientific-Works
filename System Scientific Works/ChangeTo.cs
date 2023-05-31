using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace System_Scientific_Works
{
    public partial class ChangeTo : Form
    {
        public ChangeTo(DB db, string from, string element)
        {
            InitializeComponent();

            this.db= db;

            dataGridView1.DataSource = db.FullTable($"SELECT {element} FROM {from}").Tables[0];


            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
        }

        public ChangeTo(DB db, string query)
        {
            InitializeComponent();

            this.db = db;

            dataGridView1.DataSource = db.FullTable(query).Tables[0];

            dataGridView1.Columns[0].Visible = false;

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
        }

        DB db;

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Columns.Count == 2)
            {
                newId = Convert.ToInt32(dataGridView1[0, e.RowIndex].Value);
                newName = dataGridView1[1, e.RowIndex].Value.ToString();
            }
            else 
            { 
                newId = Convert.ToInt32(dataGridView1[2, e.RowIndex].Value);
                newName = dataGridView1[3, e.RowIndex].Value.ToString();
                newName2 = dataGridView1[1, e.RowIndex].Value.ToString();
            }

            DialogResult= DialogResult.OK;
        }

        public int newId;
        public string newName;
        public string newName2;

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;

                dataGridView1.Visible = true;

                string search = textBox1.Text.Trim().ToLower();

                if (search == null) return;


                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {

                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        if (dataGridView1[j, i].Value.ToString().Trim().ToLower().Contains(search))
                        {
                            dataGridView1.Rows[i].Visible = true;
                            break;
                        }
                        else
                        {
                            dataGridView1.CurrentCell = null;
                            dataGridView1.Rows[i].Visible = false;
                        }

                    }
                }
            }
        }
    }
}
