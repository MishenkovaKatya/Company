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
    public partial class Form4 : Form
    {
        private System.Data.OleDb.OleDbDataAdapter dAdapter;
        private System.Data.DataSet dSet;
        private System.Data.DataTable dTable;
        String nameDep = "";
        public Form4(String name, String con)
        {
            InitializeComponent();
            cn.ConnectionString = con;
            cn.Open();
            nameDep = name;
            textBox1.KeyPress += TextBox1_KeyPress;
            textBox2.KeyPress += TextBox1_KeyPress;
            textBox3.KeyPress += TextBox1_KeyPress;
        }

        //
        // Enter only letters.
        //
        private void TextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsLetter(e.KeyChar)) && !(char.IsControl(e.KeyChar)))
                e.Handled = true;
        }

        //
        // Insert a new record in database.
        //
        private void Button1_Click(object sender, EventArgs e)
        {
            dSet = new DataSet();
            String strSQL = "SELECT count(*) from Empoyee WHERE FirstName = '" + textBox1.Text + " ' AND SurName = '" + textBox2.Text + "' AND DocSeries = '" + maskedTextBox2.Text + "' AND DocNumber = '" + maskedTextBox3.Text + "'";
            dAdapter = new OleDbDataAdapter(strSQL, cn);
            dAdapter.Fill(dSet, "Department");
            dTable = dSet.Tables["Department"];
            String res = dTable.Rows[0][0].ToString();
            int repeatEmp = Convert.ToInt32(res);

            if (textBox1.Text == "")
                MessageBox.Show("Введите имя!", "Некорректные данные", MessageBoxButtons.OK);
            else if (textBox2.Text == "")
                MessageBox.Show("Введите фамилию!", "Некорректные данные", MessageBoxButtons.OK);
            else if (textBox3.Text == "")
                MessageBox.Show("Введите отчество!", "Некорректные данные", MessageBoxButtons.OK);
            else if (maskedTextBox1.Text == "  .  .")
                MessageBox.Show("Введите дату рождения!", "Некорректные данные", MessageBoxButtons.OK);
            else if (maskedTextBox2.Text == "")
                MessageBox.Show("Введите серию паспорта!", "Некорректные данные", MessageBoxButtons.OK);
            else if (maskedTextBox3.Text == "")
                MessageBox.Show("Введите номер паспорта!", "Некорректные данные", MessageBoxButtons.OK);
            else if (textBox7.Text == "")
                MessageBox.Show("Введите должность!", "Некорректные данные", MessageBoxButtons.OK);
            else if (repeatEmp > 0)
            {
                MessageBox.Show("Сотрудник с таким данными уже существует!");
            }
            else
            {
                dSet = new DataSet();
                strSQL = "SELECT ID FROM Department WHERE NAME = '" + nameDep + "'";
                dAdapter = new OleDbDataAdapter(strSQL, cn);
                dAdapter.Fill(dSet, "Department");
                dTable = dSet.Tables["Department"];
                String idDep =  dTable.Rows[0][0].ToString();


                strSQL = "INSERT INTO Empoyee(FirstName, SurName, Patronymic, DateOfBirth, DocSeries, DocNumber, Position, DepartmentID) VALUES(?,?,?,?,?,?,?,?)";
                OleDbCommand cmdIC = new OleDbCommand(strSQL, cn);
                cmdIC.Parameters.Add("@p1", OleDbType.VarChar, 50);
                cmdIC.Parameters.Add("@p2", OleDbType.VarChar, 50);
                cmdIC.Parameters.Add("@p3", OleDbType.VarChar, 50);
                cmdIC.Parameters.Add("@p4", OleDbType.Date);
                cmdIC.Parameters.Add("@p5", OleDbType.VarChar, 4);
                cmdIC.Parameters.Add("@p6", OleDbType.VarChar, 6);
                cmdIC.Parameters.Add("@p7", OleDbType.VarChar, 50);
                cmdIC.Parameters.Add("@p8", OleDbType.VarChar, 50);

                cmdIC.Parameters[0].Value = textBox1.Text;
                cmdIC.Parameters[1].Value = textBox2.Text;
                cmdIC.Parameters[2].Value = textBox3.Text;
                cmdIC.Parameters[3].Value = maskedTextBox1.Text;
                cmdIC.Parameters[4].Value = maskedTextBox2.Text;
                cmdIC.Parameters[5].Value = maskedTextBox3.Text;
                cmdIC.Parameters[6].Value = textBox7.Text;
                cmdIC.Parameters[7].Value = idDep;

                try
                {
                    cmdIC.ExecuteNonQuery();
                    MessageBox.Show("Новый сотрудник добавлен!");
                    this.Close();
                }
                catch (OleDbException exc)
                {
                    MessageBox.Show(exc.ToString());
                }
            }
        }
    }
}
