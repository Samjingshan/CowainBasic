using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MyFuncDll;
using System.Configuration;
namespace Sharetek
{
    public partial class Form_LogIn : Form
    {       
       
        public Form_LogIn()
        {
            InitializeComponent();
           
        }

        private void button1_Click(object sender, EventArgs e)
        {

            string status_user;
            string permision = "";
            string querystr = "select  top 1 * from uname where userID='" + txt_username.Text.Trim().ToString() + "' and pwd='" + txt_password.Text.Trim().ToString() + "' order by create_time desc ";
            DataSet ds = SQLServer.Query(querystr,  MainClass.customConnect);
            if (ds.Tables.Count > 0)
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        status_user = ds.Tables[0].Rows[i]["status_user"].ToString();
                        permision = ds.Tables[0].Rows[i]["permision"].ToString();
                        status_user = status_user.Replace(" ","");
                        permision = permision.Replace(" ","");
                        if (status_user != "启用")
                        {
                            txt_username.Text = "";
                            txt_password.Text = "";
                            txt_username.Focus();
                            MessageBox.Show("该用户已停用，请联系管理员！！");
                            return;
                        }
                      
                        FrmMain fm = new FrmMain(txt_username.Text.Trim().ToString(), permision);
                        fm.Show();
                        

                        this.Hide();
                    }
                }
                else
                {
                    //txt_username.Text = "";
                    txt_password.Text = "";
                    txt_username.Focus();
                    MessageBox.Show("用户名或密码错误，请重新输入！！");
                    return;
                }
            }
           
            
        }

        

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                txt_password.Focus();//当在文本框1中检查到回车键时，直接将焦点转入密码框
            }
        }
        private int CheckRunning()
        {
            int ProceedingCount = 0;
            string ss;

            System.Diagnostics.Process[] Processes;
            Processes = System.Diagnostics.Process.GetProcesses();
            foreach (System.Diagnostics.Process IsProcedding in Processes)
            {
                ss = IsProcedding.ProcessName;
                if (ss == "Sharetek")
                {
                    ProceedingCount += 1;
                }
            }
            if (ProceedingCount > 1)
            {
                DialogResult result;
                result = MessageBox.Show("程序已经运行, 不允许重复打开!!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return -1;
            }
            else
            {
                return 0;
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
           // int ii = CheckRunning();

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            System.Environment.Exit(0);
        }
    }
}
