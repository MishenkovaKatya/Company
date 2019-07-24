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
    public partial class Form3 : Form
    {
        private System.Data.OleDb.OleDbDataAdapter dAdapter;
        private System.Data.DataSet dSet;
        private System.Data.DataTable dTable;
        String idEmp = "";
        public Form3(String ID, String FirstName, String SurName, String Patronymic, String con)
        {
            InitializeComponent();
            idEmp = ID;
            textBox7.Text = SurName ;
            textBox8.Text = FirstName;
            textBox9.Text = Patronymic;
            cn.ConnectionString = con;
            cn.Open();
            CreateEmployeeInformation();
            textBox7.KeyPress += TextBox1_KeyPress;
            textBox8.KeyPress += TextBox1_KeyPress;
            textBox9.KeyPress += TextBox1_KeyPress;
        }

        //
        // Employee details.
        //
        void CreateEmployeeInformation()
        {
            DataSet dSet = RunQuery("SELECT CAST(DateOfBirth AS date), DocSeries, DocNumber, Position FROM Empoyee WHERE ID = '" + idEmp + "'");

            DateTime curDate = DateTime.Today;
            String bD = dTable.Rows[0][0].ToString();
            DateTime birthDate = DateTime.Parse(bD);
            int age = curDate.Year - birthDate.Year;
            DateTime tempDate = curDate.AddYears(-age);
            if (birthDate > tempDate)
                age--;

            textBox2.Text = age.ToString();
            maskedTextBox1.Text = dTable.Rows[0]["DocSeries"].ToString();
            maskedTextBox2.Text = dTable.Rows[0]["DocNumber"].ToString();
            maskedTextBox3.Text = birthDate.ToString("d");
            textBox6.Text = dTable.Rows[0]["Position"].ToString();
        }

        //
        // Enable edit.
        //
        private void Button1_Click(object sender, EventArgs e)
        {
            maskedTextBox3.ReadOnly = false;
            maskedTextBox1.ReadOnly = false;
            maskedTextBox2.ReadOnly = false;
            textBox6.ReadOnly = false;
            textBox7.ReadOnly = false;
            textBox8.ReadOnly = false;
            textBox9.ReadOnly = false;
            button2.Visible = true;

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
        // Update information of employee in database.
        //
        private void Button2_Click(object sender, EventArgs e)
        {
            DateTime correctDate;
            if (textBox8.Text == "")
                MessageBox.Show("Введите имя!", "Некорректные данные", MessageBoxButtons.OK);
            else if (textBox7.Text == "")
                MessageBox.Show("Введите фамилию!", "Некорректные данные", MessageBoxButtons.OK);
            else if (textBox9.Text == "")
                MessageBox.Show("Введите отчество!", "Некорректные данные", MessageBoxButtons.OK);
            else if (maskedTextBox3.Text == "  .  .")
                MessageBox.Show("Введите дату рождения!", "Некорректные данные", MessageBoxButtons.OK);
            else if (!DateTime.TryParse(maskedTextBox3.Text, out correctDate))
                MessageBox.Show("Дата рождения введена некорректна! Введите дату в формате dd.mm.yyyy.", "Некорректные данные", MessageBoxButtons.OK);
            else if (maskedTextBox1.Text == "")
                MessageBox.Show("Введите серию паспорта!", "Некорректные данные", MessageBoxButtons.OK);
            else if (maskedTextBox2.Text == "")
                MessageBox.Show("Введите номер паспорта!", "Некорректные данные", MessageBoxButtons.OK);
            else if (textBox6.Text == "")
                MessageBox.Show("Введите должность!", "Некорректные данные", MessageBoxButtons.OK);
            else
            {
                dSet = new DataSet();
                String strSQL = "UPDATE Empoyee SET FirstName=?, SurName=?, Patronymic=?, DateOfBirth=?, DocSeries=?, DocNumber=?, Position=? WHERE ID = '" + idEmp + "'";
                OleDbCommand cmdIC = new OleDbCommand(strSQL, cn);
                cmdIC.Parameters.Add("@name", OleDbType.VarChar, 50);
                cmdIC.Parameters.Add("@sur", OleDbType.VarChar, 50);
                cmdIC.Parameters.Add("@pat", OleDbType.VarChar, 50);
                cmdIC.Parameters.Add("@DoB", OleDbType.Date);
                cmdIC.Parameters.Add("@Ser", OleDbType.VarChar, 4);
                cmdIC.Parameters.Add("@Num", OleDbType.VarChar, 6);
                cmdIC.Parameters.Add("@pos", OleDbType.VarChar, 50);

                String dr = maskedTextBox3.Text;
                DateTime parserDate = DateTime.Parse(dr);
                cmdIC.Parameters[0].Value = textBox8.Text;
                cmdIC.Parameters[1].Value = textBox7.Text;
                cmdIC.Parameters[2].Value = textBox9.Text;
                cmdIC.Parameters[3].Value = parserDate;
                cmdIC.Parameters[4].Value = maskedTextBox1.Text;
                cmdIC.Parameters[5].Value = maskedTextBox2.Text;
                cmdIC.Parameters[6].Value = textBox6.Text;

                try
                {
                    cmdIC.ExecuteNonQuery();
                    MessageBox.Show("Информация о сотруднике обновлена!");
                    Class1.DataIsRecieved = true;
                    this.Close();
                }
                catch (OleDbException exc)
                {
                    MessageBox.Show(exc.ToString());
                }
            }
        }

        //
        // Processing request.
        //
        DataSet RunQuery(String Query)
        {
            dSet = new DataSet();
            dAdapter = new OleDbDataAdapter(Query, cn);
            dAdapter.Fill(dSet, "Department");
            dTable = dSet.Tables["Department"];
            return dSet;
        }

    }
}
