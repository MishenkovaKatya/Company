using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Microsoft.VisualBasic;


namespace Company
{
    public partial class Form1 : Form
    {
        private System.Data.OleDb.OleDbDataAdapter dAdapter;
        private System.Data.DataSet dSet;
        private System.Data.DataTable dTable;
        String strConnect = "";

        public Form1()
        {
            InitializeComponent();
            treeView1.BeforeExpand += TreeView1_BeforeExpand;
            treeView1.MouseDoubleClick += TreeView1_MouseDoubleClick;
        }

        //
        //Search root of tree.
        //
        void CreateRoot()
        {
            if (strConnect != "")
            {
                DataSet dSet = RunQuery("SELECT ID, Name FROM Department WHERE ParentDepartmentID IS NULL");
                int numRoot = dTable.Rows.Count;
                for (int i = 0; i < numRoot; i++)
                {
                    TreeNode root = new TreeNode(dTable.Rows[i][1].ToString());
                    CreateNode(root, dTable.Rows[i][0].ToString());
                    treeView1.Nodes.Add(root);
                }
            }
        }

        //
        //Search child elements of tree.
        //
        void CreateNode(TreeNode root, string rootID)
        {
            DataSet dSet2 = RunQuery("SELECT ID, Name FROM Department WHERE ParentDepartmentID = '" + rootID + " '");
            int numChild = dTable.Rows.Count;
            for (int i = 0; i < numChild; i++)
            {
                root.Nodes.Add(dTable.Rows[i][1].ToString());
            }
        }

        //
        //Event before opening the node.
        //
        void TreeView1_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            int numChildNode = e.Node.Nodes.Count;
            TreeNode[] node = new TreeNode[numChildNode];
            String[] nodeName = new String[numChildNode];
            String[] nodeID = new String[numChildNode];
            TreeNode curNode = e.Node.FirstNode;
            for (int i = 0; i < numChildNode; i++)
            {
                node[i] = curNode;
                nodeName[i] = curNode.Text;
                dSet = RunQuery("SELECT ID FROM Department WHERE Name = '" + curNode.Text + "'");
                nodeID[i] = dTable.Rows[0][0].ToString();
                curNode.NodeFont = new Font("Times New Roman", 14, FontStyle.Regular);
                curNode = curNode.NextNode;
            }

            for (int i = 0; i < numChildNode; i++)
            {
                if (node[i].Nodes.Count == 0)
                {
                    dSet = RunQuery("SELECT ID, Name FROM Department WHERE ParentDepartmentID = '" + nodeID[i] + " '");
                    int numChild = dTable.Rows.Count;
                    for (int j = 0; j < numChild; j++)
                    {
                        node[i].Nodes.Add(dTable.Rows[j][1].ToString());
                    }
                }
            }

        }

        //
        //Event after double click - open information about employees of department.
        //
        void TreeView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            TreeNode selectedNode = treeView1.HitTest(e.Location).Node;
            if (selectedNode != null)
            {
                Form2 f1 = new Form2(selectedNode.Text, strConnect);
                f1.ShowDialog();
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

        //
        // Connection database.
        //
        private void ConnectToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            strConnect = Microsoft.VisualBasic.Interaction.InputBox("Введите строку подключения к базе данных.", "Соединение с БД", "Provider= SQLOLEDB.1;Data Source=DESKTOP-HFE45PV" + "\\" + "MYSERVER;Persist Security Info=True;Password=sa;User ID=sa;Initial Catalog=TestDB", 600, 350);
            if (strConnect.Length < 8 || strConnect.Substring(0, 8) != "Provider")
            {
                MessageBox.Show("Ошибка. Введены некорректные данные для подключения к базе!", "Connection", MessageBoxButtons.OK);
            }
            else
            {
                cn.ConnectionString = strConnect;
                try
                {
                    cn.Open();
                    MessageBox.Show("Соединение установлено!", "Connection", MessageBoxButtons.OK);
                }
                catch (Exception)
                {
                    MessageBox.Show("Ошибка. Отсутствует соединение с БД!", "Connection", MessageBoxButtons.OK);
                }
                if (cn.State == ConnectionState.Open)
                {
                    CreateRoot();
                }
            }
        }

        //
        // Application Information.
        //
        private void HelpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Данная программа предназначена для просмотра структуры предприятия и его сотрудников. " +
                "Основная функциональность - просмотр, добавление и редактирование сотрудников в отделах.", "Информация", MessageBoxButtons.OK);
        }

        //
        // Close application.
        //
        private void ExitToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Application.Exit();
        }

        //
        // Close connection.
        //
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            cn.Close();
        }
    }
}

