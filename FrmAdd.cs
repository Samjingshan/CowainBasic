using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MyFuncDll;
using System.Configuration;

namespace Sharetek
{
    public partial class FrmAdd : Form
    {
        FrmMain frm = new FrmMain("","");
        string customConnect = "";
        public FrmAdd(FrmMain fr)
        {
            customConnect = ConfigurationManager.ConnectionStrings["ConnectionString"].ToString();
            InitializeComponent();
            frm = fr;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Trim().ToString() == "")
            {
                MessageBox.Show("用户名不能为空！！", "警告");
                textBox1.Focus();
                return;
            }
            //判断用户名唯一性
            string querystr = "select * from uname where userID='" + textBox1.Text.Trim().ToString() + "'";
            DataSet ds = SQLServer.Query(querystr,customConnect);
            if (ds.Tables[0].Rows.Count > 0)
            {
                MessageBox.Show("该用户名已经存在，请重新输入！！", "警告");
                textBox1.Text = "";
                textBox1.Focus();
                return;
            }
            if (textBox2.Text.Trim().ToString() == "")
            {
                MessageBox.Show("密码不能为空！！", "警告");
                return;
            }
            if (textBox2.Text.Trim().ToString() != textBox3.Text.Trim().ToString())
            {
                MessageBox.Show("密码不一致，请重新输入！！", "警告");
                textBox2.Text = "";
                textBox3.Text = "";
                textBox2.Focus();
                return;
            }
            try
            {
                string insertstr = "insert into uname(userID,pwd,status_user,permision) values('" + textBox1.Text + "','" + textBox2.Text + "','" + comboBox1.Text + "','" + comboBox2.Text + "')";
                int i = SQLServer.ExecuteSql(insertstr,customConnect);
                MessageBox.Show("保存成功！！", "提示");
                frm.showdata();
                this.Hide();
            }
            catch
            {
                return;
            }
        }
    }
}
