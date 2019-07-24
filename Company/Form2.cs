using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Company
{
    public partial class Form2 : Form
    {
        private System.Data.OleDb.OleDbDataAdapter dAdapter;
        private System.Data.DataSet dSet;
        private System.Data.DataTable dTable;
        BindingSource bs = new BindingSource();
        String nameDep = "", strConnect = "";

        public Form2(String name, String con)
        {
            InitializeComponent();
            nameDep = name;
            strConnect = con;
            cn.ConnectionString = strConnect;
            cn.Open();
            textBox1.Text = name;
            dataGridView1.CellClick += DataGridView1_CellClick;
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font(dataGridView1.ColumnHeadersDefaultCellStyle.Font.FontFamily, 12f, FontStyle.Bold);
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            this.Activated += Form2_Activated;
            CreateTable();
        }

        //
        // Create table of employees.
        //
        public void CreateTable()
        {
            dSet = new DataSet();
            String Query = "SELECT ID, SurName, FirstName, Patronymic FROM Empoyee WHERE DepaRtmentID = (SELECT ID FROM Department WHERE Name = '" + nameDep + "')";
            dAdapter = new OleDbDataAdapter(Query, cn);
            dAdapter.Fill(dSet, "Empoyee");
            dTable = dSet.Tables["Empoyee"];
            dTable.Columns.Add(" ", typeof(String));
            bs.DataSource = dTable;
            dataGridView1.DataSource = bs;
            dataGridView1.Columns[0].HeaderCell.Value = "ID";
            dataGridView1.Columns[1].HeaderCell.Value = "Фамилия";
            dataGridView1.Columns[2].HeaderCell.Value = "Имя";
            dataGridView1.Columns[3].HeaderCell.Value = "Отчество";
            for (int i = 0; i < dTable.Rows.Count; i++)
            {
                dataGridView1.Rows[i].Cells[4].Value = "Подробнее...";
            }
        }

        //
        // Open employee details.
        //
        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int curCol = dataGridView1.CurrentCell.ColumnIndex;
            int curRow = dataGridView1.CurrentRow.Index;
            String ID =" ", SurName = " ", FirstName = " ", Patronymic = " ";
            String curValue = dataGridView1.Rows[curRow].Cells[curCol].Value.ToString();
            if (curCol == 4 && curValue != "" && curRow > -1)
            {
                ID = dataGridView1.Rows[curRow].Cells[0].Value.ToString();
                SurName = dataGridView1.Rows[curRow].Cells[1].Value.ToString(); 
                FirstName = dataGridView1.Rows[curRow].Cells[2].Value.ToString(); 
                Patronymic = dataGridView1.Rows[curRow].Cells[3].Value.ToString();
                Form3 f1 = new Form3(ID, FirstName, SurName, Patronymic, strConnect);
                f1.ShowDialog();

            }
        }

        //
        // Add new employee.
        //
        private void Button1_Click(object sender, EventArgs e)
        {
            Form4 f1 = new Form4(nameDep, strConnect);
            f1.ShowDialog();
        }

        private void Form2_Activated(object sender, EventArgs e)
        {
            if (Class1.DataIsRecieved == true)
            {
                CreateTable();
            }
            Class1.DataIsRecieved = false;
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = false;
        }

    }
}
