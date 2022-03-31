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
    public partial class Frmupdate : Form
    {
        string customConnect = "";
        FrmMain frm = new FrmMain("","");
        public Frmupdate(FrmMain fr)
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
            if (textBox3.Text.Trim().ToString() == "")
            {
                MessageBox.Show("新密码不能为空！！", "警告");
                return;
            }
            if (textBox4.Text.Trim().ToString() == "")
            {
                MessageBox.Show("确认密码不能为空！！", "警告");
                return;
            }
            if (textBox4.Text.Trim().ToString() != textBox3.Text.Trim().ToString())
            {
                MessageBox.Show("输入的密码不一致，请重新输入！！", "警告");
                textBox3.Text = "";
                textBox4.Text = "";
                textBox3.Focus();
                return;
            }
            if (comboBox1.Text.Trim().ToString() == "")
            {

                MessageBox.Show("请选择用户状态！！", "警告");
                return;
            }
            if (MessageBox.Show("是否确定修改记录？", "修改", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                string updatestr = "update uname set pwd='" + textBox3.Text + "',status_user='" + comboBox1.Text + "' ,permision='" + comboBox2.Text + "' where userID='" + textBox1.Text.Trim().ToString() + "'";
                int i = SQLServer.ExecuteSql(updatestr,customConnect);
                MessageBox.Show("修改成功！！", "提示");
                this.Hide();
                frm.showdata();

            }

        }

        private void Frmupdate_Load_1(object sender, EventArgs e)
        {
            textBox1.Text = frm.yhm;
            textBox2.Text = frm.pwd;
            comboBox1.Text = frm.zt;
            textBox1.Enabled = false;
            textBox2.Enabled = false;
            if (textBox1.Text.Trim().ToString() == "admin")
            {
                comboBox1.Enabled = false;
            }
        }
    }
}
