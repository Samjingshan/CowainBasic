using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;

using MyFuncDll;
using System.Data.Sql;
using System.Data.SqlClient;
using System.IO;
using System.Configuration;
using System.Net;
using System.Net.Sockets;
using System.Collections.Specialized;
using System.Windows.Forms.DataVisualization.Charting;
using System.Media;
using System.Collections;
using Aspose.Cells;
using PLC_Adapt;
using Sharetek_Common;


namespace Sharetek
{


    public partial class FrmMain : Form
    {

        public static PLCManager plcManager;
        public System.Timers.Timer HeratBeatT = new System.Timers.Timer(1000);
        public string Log_Path = "";

        public string Logger_str = "";
        FTP_Fun Ftp_Connect = null;


      



        //IPAddress ipAddr2;
        // IPEndPoint ipEndPoint2;

        int beginPosition = 0;



        /// <summary>
        /// 从CustomSocket读写操作的数组 
        /// </summary>
        byte[] customBytes = new byte[1024];



        
     

      
        string CustomConnect = "";

        string permision = "";
        string person = "";
        public string yhm;
        public string pwd;
        public string zt;
        private string sql_BCR = "";
        private string sql_ERR = "";
        private string sql_Log = "";
        private string sql_Huo_Log = "";


    





        public FrmMain(string Name, string Permision)
        {
            InitializeComponent();

            this.person = Name;
            this.permision = Permision;

            CustomConnect = ConfigurationManager.ConnectionStrings["ConnectionString"].ToString();
          
            NameValueCollection appSettings = ConfigurationManager.AppSettings;

            if (appSettings == null || appSettings.Count <= 0)
            {
                PublicFunc.WritelnTxt("AppSetting", "app.config 中无AppSetting节点", "Error.txt");
            }
            else
            {




                

                //Info.FileName = ConfigurationManager.AppSettings["FileName"];


            }

    
            Dictionary<string, string> configDict = new Dictionary<string, string>();
            string _plcmemoryfile = ConfigurationManager.AppSettings["PLCMemoryFileName"] ?? "PLCMemory_s7.txt";
            configDict.Add("PLCMemoryFileName", _plcmemoryfile);

            PLCAdapter adapter = null;

            string _adapter = ConfigurationManager.AppSettings["PLCAdapter"] ?? "FileAdapter";
            string _ip = ConfigurationManager.AppSettings["IP"] ?? "192.168.1.10";
            string _port = ConfigurationManager.AppSettings["Port"] ?? "6000";
            string _PLCFile = ConfigurationManager.AppSettings["PLC_ConfigFile"] ?? "SiemensS7.txt";
            string PLCFile_Info= ConfigurationManager.AppSettings["PLC_ConfigFile_Info "] ?? "Siemens_Info.txt ";
            string _resolution = ConfigurationManager.AppSettings["resolution"] ?? "500";
            string _s7BaseAddress = ConfigurationManager.AppSettings["SiemensS7BasedAddress"] ?? "D15";
            string _s7Offset = ConfigurationManager.AppSettings["SiemensS7Offset"] ?? "77";
            string _SiemensMode = ConfigurationManager.AppSettings["SiemensMode"] ?? "3";
            Log_Path = ConfigurationManager.AppSettings["Log_path"];

            string FTP_IP = ConfigurationManager.AppSettings["FTP_IP"];
            string USER = ConfigurationManager.AppSettings["USER"];
            string PASSWORD = ConfigurationManager.AppSettings["PASSWORD"];
            string RemoteGetFilePath = ConfigurationManager.AppSettings["RemoteGetFilePath"];




            configDict.Add("SiemensS7Offset", _s7Offset);
            configDict.Add("SiemensS7BasedAddress", _s7BaseAddress);
            configDict.Add("IP", _ip);
            configDict.Add("PLCAdapter", _adapter);
            configDict.Add("Port", _port);
            configDict.Add("PLCConfigFile", _PLCFile);
            configDict.Add("PLC_ConfigFile_Info", PLCFile_Info);

            configDict.Add("resolution", _resolution);
            configDict.Add("SiemensMode", _SiemensMode);
            configDict.Add("Log_path", Log_Path);
            configDict.Add("PASSWORD", PASSWORD);
            configDict.Add("FTP_IP", FTP_IP);
            configDict.Add("USER", USER);
            configDict.Add("RemoteGetFilePath", RemoteGetFilePath);
            if (_adapter == "FileAdapter")
            {
                adapter = new FilePLCAdapter();
                adapter.Init(configDict);
            }
            else if (_adapter == "Melsec")
            {
                adapter = new MelsecPLCAdapter();
                adapter.Init(configDict);
            }
            else if (_adapter == "S7Siemens")
            {
                adapter = new SiemensPLCAdapter();
                adapter.Init(configDict);
            }
            #region
           
            plcManager = new PLCManager(Log_Path, Logger_str, Int32.Parse(_resolution),
                                      adapter, _PLCFile, PLCFile_Info);
            plcManager.Init();//将定义PLC交互信号的2个文件 里所有的字段放到字典，key 字段名 value:PLCItem    key 字段名 value:PLCAlarm
            plcManager.Connect();//PLCEvent绑定绑定方法，若键值有变化，则触发方法
            plcManager.StartReading();
           
            HeratBeatT.Elapsed += new System.Timers.ElapsedEventHandler(HeratBeatT_Elapsed);
            HeratBeatT.Start();


            

            #endregion


            Task.Run(() => {
                try
                {
                    System.Threading.Thread.Sleep(1000);

                    DoWithPLC doWithPLC = new DoWithPLC(plcManager,Log_Path);


                }
                catch (Exception ex)
                {
                    PublicFunc.WritelnTxt(Log_Path,"异常发生行号："+ex.StackTrace+"异常："+ex.Message,"ERROR");
                }
            });
         







        }
        private void HeratBeatT_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            HeratBeatT.Enabled = true;
            //if (plcManager.isConnected)
            //{
            //    plc_label.ForeColor = Color.Green;
            //    plc_label.Text = "已连接";

            //}
            //else
            //{
            //    plc_label.ForeColor = Color.Red;
            //    plc_label.Text = "连接失败,请检查设备连接！";


            //}
        }

            private void PaintColorToRtxtbx(System.Windows.Forms.RichTextBox rtxtBx, string inputstr, ref int _iBegin,ref int PaintColor_count)
        {
            rtxtBx.AppendText(inputstr);
            rtxtBx.Select(_iBegin, inputstr.Length);
            rtxtBx.SelectionColor = Color.Black;
            rtxtBx.Select(_iBegin, 0);
            _iBegin = _iBegin + inputstr.Length;
           
            PaintColor_count++;
            if (PaintColor_count > 1000)
            {
                rtxtBx.Clear();
                PaintColor_count = 0;

            }
        }
        private double SubDateTimeByMill(DateTime dt1, DateTime dt2)
        {
            double i = 0;
            TimeSpan ts1 = new TimeSpan(dt1.Ticks);
            TimeSpan ts2 = new TimeSpan(dt2.Ticks);
            TimeSpan ts = ts1.Subtract(ts2).Duration();
            i = Convert.ToInt32(ts.TotalMilliseconds);
            return i;
        }
        private void FrmMain_Load(object sender, EventArgs e)
        {
            string Txt_title =  "三花（" + permision + "）";
            label3.Visible = false;
            label5.Visible = false;
            //permision = this.Text.ToString();

            //string[] arr = permision.Split(',');
            //person = arr[0];
            //permision = arr[1];
            if (permision != "管理员")
            {

                tabPage5.Parent = null;
            }
          
            

           

            #region
            //使用自定义格式
            dateTimePicker1.CustomFormat = "yyyy/MM/dd HH:mm：ss";
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.Value = DateTime.Now.Date;
            dateTimePicker1.Value = DateTime.Now.Date.AddHours(00).AddMinutes(00).AddSeconds(00);
            //使用自定义格式
            dateTimePicker2.CustomFormat = "yyyy/MM/dd HH:mm：ss";          
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.Value = DateTime.Now.Date;
            dateTimePicker2.Value = DateTime.Now.Date.AddHours(23).AddMinutes(59).AddSeconds(59);
          
       






            #endregion

            this.Text = Txt_title;
            username.Text = "当前用户：" + person;
            username.ForeColor = Color.Green;
            timer1.Enabled = false;
            timer2.Enabled = false;
            timer3.Enabled = true;
            
          
           
         
          
           
        }
     

        


      



     

        private void button2_Click(object sender, EventArgs e)
        {
          
            if (MessageBox.Show("是否确定关闭程序？", "关闭程序", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
               // myResetEvent1.Reset();
               
                
              
                System.Threading.Thread.Sleep(500);
               

                this.timer1.Enabled = true;
                this.timer1.Dispose();
                this.timer2.Enabled = false;
                this.timer2.Dispose();

             
                string hint = "当前用户：" + person;
                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "Close", hint, "txt");
                //try { Proc.Kill(); }
                //catch { ;}
                System.Environment.Exit(0);

            }
        }
        private void  Byte_Arr_to_Int(int n,Byte[] arr,ref string result)
        {
            
            int a = Convert.ToInt16(arr[0 + 2 * n].ToString());
            int b = Convert.ToInt16(arr[1 + 2 * n].ToString());
            if (a == 255)//负数
            {
                result = (b - 256).ToString();
            }
            else
            {
                result = (a * 256 + b).ToString();
            }
          

           
        }

        private void Deal_with_Log(string  palno,  string num, string port, string errormsg ,  string connectStr)
        {
            SqlParameter[] sqlParameters_Log ={
                                                  new SqlParameter("@palno",SqlDbType.NVarChar,100),
                                                  new SqlParameter("@num",SqlDbType.NVarChar,100),
                                                  new SqlParameter("@port",SqlDbType.NVarChar,100),
                                                  new SqlParameter("@errormsg",SqlDbType.NVarChar,100)
                                                
                                                 
                                              };
            sqlParameters_Log[0].Value = palno;
            sqlParameters_Log[1].Value = num;
            sqlParameters_Log[2].Value = port;
            sqlParameters_Log[3].Value = errormsg;
           

            string[] paraName = { "@palno" ,"@num", "@port", "@errormsg" };
            object[] paraOutValue = new object[4];
            bool _result_weight = SQLServer.RunProc("[sh_Insert_Log]", sqlParameters_Log, paraName, ref paraOutValue, connectStr);


            //over_Weight = Convert.ToInt32(paraOutValue[0]);
        }
        private void Deal_with_HDR_Log(string palno, string G2P_NO, string Station, string STATUS, string errormsg, string connectStr)
        {
            SqlParameter[] sqlParameters_Log ={
                                                  new SqlParameter("@palno",SqlDbType.NVarChar,100),
                                                  new SqlParameter("@G2P_NO",SqlDbType.NVarChar,100),
                                                  new SqlParameter("@Station",SqlDbType.NVarChar,100),
                                                  new SqlParameter("@STATUS",SqlDbType.NVarChar,100),
                                                  new SqlParameter("@errormsg",SqlDbType.NVarChar,100)


                                              };
            sqlParameters_Log[0].Value = palno;
            sqlParameters_Log[1].Value = G2P_NO;
            sqlParameters_Log[2].Value = Station;
            sqlParameters_Log[3].Value = STATUS;
            sqlParameters_Log[4].Value = errormsg;


            string[] paraName = { "@palno", "@G2P_NO", "@Station", "@STATUS", "@errormsg" };
            object[] paraOutValue = new object[5];
            bool _result_weight = SQLServer.RunProc("[sh_Insert_HDR_Log]", sqlParameters_Log, paraName, ref paraOutValue, connectStr);


            //over_Weight = Convert.ToInt32(paraOutValue[0]);
        }
        private void Deal_with_CCS_to_WMS_Log(string Box_num, string Lpn, int  num,  string connectStr)
        {
            SqlParameter[] sqlParameters_Log ={
                                                  new SqlParameter("@Box_num",SqlDbType.NVarChar,100),
                                                  new SqlParameter("@Lpn",SqlDbType.NVarChar,100),
                                                  new SqlParameter("@num",SqlDbType.Int)                                               
                                              };
            sqlParameters_Log[0].Value = Box_num;
            sqlParameters_Log[1].Value = Lpn;
            sqlParameters_Log[2].Value = num;
           


            string[] paraName = { "@Box_num", "@Lpn", "@num" };
            object[] paraOutValue = new object[3];
            bool _result_weight = SQLServer.RunProc("[sh_Insert_CCS_TO_WMS]", sqlParameters_Log, paraName, ref paraOutValue, connectStr);


            
        }

        private void Deal_with_ERRINFO(string bar_code, string num, string ERR_INFO, string connectStr)
        {
            SqlParameter[] sqlParameters_Log ={
                                                  new SqlParameter("@bar_code",SqlDbType.NVarChar,100),
                                                  new SqlParameter("@num",SqlDbType.NVarChar,100),
                                                  new SqlParameter("@ERR_INFO",SqlDbType.NVarChar,500)
                                              };
            sqlParameters_Log[0].Value = bar_code;
            sqlParameters_Log[1].Value = num;
            sqlParameters_Log[2].Value = ERR_INFO;
          

            string[] paraName = { "@bar_code", "@num", "@ERR_INFO" };
            object[] paraOutValue = new object[3];
            bool _result_weight = SQLServer.RunProc("[sh_Insert_ERRINFO]", sqlParameters_Log, paraName, ref paraOutValue, connectStr);


            //over_Weight = Convert.ToInt32(paraOutValue[0]);
        }
        private string  Deal_with_SOKET_Log(string bar_code, string Date, string connectStr)
        {
            string send_id = "";
            SqlParameter[] sqlParameters_Log ={
                                                  new SqlParameter("@Date",SqlDbType.NVarChar,100),
                                                  new SqlParameter("@bar_code",SqlDbType.NVarChar,100),
                                                   new SqlParameter("@Send_Id",SqlDbType.NVarChar,100)


                                              };
            sqlParameters_Log[0].Value = Date;
            sqlParameters_Log[1].Value = bar_code;
            sqlParameters_Log[2].Direction = ParameterDirection.Output;



            string[] paraName = { "@Send_Id" };
            object[] paraOutValue = new object[1];
            bool _result_Log = SQLServer.RunProc("[sh_Insert_Socket_Log]", sqlParameters_Log, paraName, ref paraOutValue, connectStr);

       
            try
            {
                send_id = paraOutValue[0].ToString();
          
            }
            catch
            {
                send_id = "";
            }
            return send_id;
        }


        public static string AddZero(string addZeroStr, int strLen)
        {
            int _len = addZeroStr.Length;
            if (_len >= strLen)
            {
                return addZeroStr;
            }
            else
            {
                while (_len < strLen)
                {
                    addZeroStr = "0" + addZeroStr;
                    _len = addZeroStr.Length;
                }
            }
            return addZeroStr;
        }
      
      
    



   









        private string format_barcode(string barcode)
        {
          
            char strtempa = (char)0x02;
            char strtempb = (char)0x03;
            //我们要求c---g之间的字符串，也就是：defghi
            //求得strtempa 和 strtempb 出现的位置:
            int IndexofA = barcode.IndexOf(strtempa);
            int IndexofB = barcode.IndexOf(strtempb);
            if(IndexofA>0 && IndexofB>0)
                barcode = barcode.Substring(IndexofA + 1, IndexofB - IndexofA - 1);
            return barcode;
        }




        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;


          
            timer1.Enabled = true;

        }

       

        private void FrmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
           
     

            if (MessageBox.Show("是否确定关闭程序？", "关闭程序", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {


             
                
                System.Threading.Thread.Sleep(500);


                this.timer1.Enabled = false;
                this.timer1.Dispose();
                this.timer2.Enabled = false;
                this.timer2.Dispose();

                //try { Proc.Kill(); }
                //catch { ;}
                string hint = "当前用户：" + person;
                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "Close", hint, "txt");

                System.Environment.Exit(0);

            }

            //Application.Exit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
         

            //writeSocket("12345", 1, out sendSck_XJ1, senderCustomSocket1, "XJ1Sck", myResetEvent1, readCustomThread1);

        }

        private void button4_Click(object sender, EventArgs e)
        {
            label3.Visible = true;
            button14.Visible=false;
            sql_Log = "select  palno as '条码号'"+
            ",num as '位置',port as '目的地',tkdat as '交互时间',errormsg as '错误信息' from View_Box_Log_All where 1=1 ";

            if (checkBox1.Checked)
            {
                if (string.IsNullOrEmpty(textBox2.Text.Trim()))
                {
                    MessageBox.Show("请输入查询条码！！", "提示");
                    textBox2.Focus();
                    return;
                }
                else
                {
                    sql_Log = sql_Log + " and palno like '%" + textBox2.Text.Trim() + "%'";
                }
            }
            if (checkBox8.Checked)
            {
                if (string.IsNullOrEmpty(textBox3.Text.Trim()))
                {
                    MessageBox.Show("请输入扫码器编号！！", "提示");
                    textBox3.Focus();
                    return;
                }
                else
                {
                    sql_Log = sql_Log + " and num like '%" + textBox3.Text.Trim() + "%'";
                }
            }
            //if (checkBox2.Checked)
            //{
                sql_Log = sql_Log + " and tkdat>='" + dateTimePicker1.Value.ToString() + "' and tkdat<='" + dateTimePicker2.Value.ToString() + "'";
            //}
            sql_Log = sql_Log + "  order by tkdat desc";
            label3.Text = "";
            DataSet ds1 = SQLServer.Query(sql_Log, MainClass.customConnect);
            label3.Text = "记录数:" + ds1.Tables[0].Rows.Count.ToString();
            dataGridView1.DataSource = ds1.Tables["ds"];
            // 设定包括Header和所有单元格的列宽自动调整
          //  dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            // 设定包括Header和所有单元格的行高自动调整
            //dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            //dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            //dataGridView1.Columns[0].Width = 200;
            ////dataGridView1.Columns[1].Width = 70;
            ////dataGridView1.Columns[2].Width = 70;
            ////dataGridView1.Columns[3].Width = 70;
            //dataGridView1.Columns[1].Width = 100;// dataGridView3.RowHeadersWidth
           
            //dataGridView1.Columns[2].Width = 100;// dataGridView3.RowHeadersWidth
            //dataGridView1.Columns[3].Width = 150;// dataGridView3.RowHeadersWidth
            //dataGridView1.Columns[4].Width = 310;// dataGridView3.RowHeadersWidth
            ////dataGridView1.Columns[9].Width = 200;// dataGridView3.RowHeadersWidth
            //dataGridView1.RowHeadersWidth = 50;
            ////添加行号
            //for (int i = 0; i < dataGridView1.Rows.Count; i++)
            //{
            //    dataGridView1.Rows[i].HeaderCell.Value = (i + 1).ToString();
            //}
            button14.Visible = true;
        }

        private void Button9_Click(object sender, EventArgs e)
        {
            showdata();
        }

        public void showdata()
        {
            this.Text = permision;
            string sql = "select userID as 用户名,status_user as 状态,permision as 权限 from uname";
            DataSet ds = SQLServer.Query(sql, CustomConnect);
            dataGridView11.DataSource = ds.Tables["ds"];
            for (int i = 0; i < this.dataGridView11.Rows.Count; i++)
            {
                DataGridViewRow r = this.dataGridView11.Rows[i];
                r.HeaderCell.Value = string.Format("{0}", i + 1);
            }
            this.dataGridView11.Refresh();
        }

        private void Button8_Click(object sender, EventArgs e)
        {

            FrmAdd fr = new FrmAdd(this);
            fr.Show();
        }

        private void Button7_Click(object sender, EventArgs e)
        {
            try
            {
                //string updatestr = "update uname set userID='" + textBox1.Text + "',pwd='" + textBox2.Text + "',status_user='" + comboBox1.Text + "' where userID='" + dataGridView1[0, dataGridView1.CurrentRow.Index].Value.ToString() + "'";
                //int i = SQLServer.ExecuteSql(updatestr);
                string sql = "select userID as 用户名,pwd as 密码,status_user as 状态 from uname where userID='" + dataGridView11[0, dataGridView11.CurrentRow.Index].Value.ToString() + "'";
                DataSet ds = SQLServer.Query(sql, CustomConnect);
                if (ds.Tables.Count > 0)
                {
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            yhm = ds.Tables[0].Rows[i][0].ToString();
                            pwd = ds.Tables[0].Rows[i][1].ToString();
                            zt = ds.Tables[0].Rows[i][2].ToString();
                        }
                    }
                }
                Frmupdate frm = new Frmupdate(this);
                frm.Text = permision;
                frm.Show();
            }
            catch
            {

                return;
            }
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            try
            {


                string s = dataGridView11[0, dataGridView11.CurrentRow.Index].Value.ToString();
                s = s.Replace(" ", "");
                if (MessageBox.Show("是否确定删除用户名为:'" + s + "'的用户？", "删除", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    string updatestr = "delete from  uname where userID='" + dataGridView11[0, dataGridView11.CurrentRow.Index].Value.ToString() + "'";
                    int i = SQLServer.ExecuteSql(updatestr, CustomConnect);
                    this.Text = permision;
                    string sql = "select userID as 用户名,status_user as 状态,permision as 权限  from uname";
                    DataSet ds = SQLServer.Query(sql, CustomConnect);
                    dataGridView11.DataSource = ds.Tables["ds"];
                   
                }
            }
            catch 
            {
                return;
            }
        }

        private void Button5_Click(object sender, EventArgs e)
        {
           
            if (MessageBox.Show("是否确定关闭程序？", "关闭程序", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
               
              
         
             
                System.Threading.Thread.Sleep(500);
              

                this.timer1.Enabled = false;
                this.timer1.Dispose();
                this.timer2.Enabled = false;
                this.timer2.Dispose();

              
                //try { Proc.Kill(); }
                //catch { ;}
                string hint = "当前用户：" + person;
                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "Close", hint, "txt");
               
                System.Environment.Exit(0);

            }
        }

     

     

       


        private void ExportInvoiceExcel(string sqlstr)
        {
            Workbook workbook = new Workbook();
            Worksheet sheet = (Worksheet)workbook.Worksheets[0];
            //设置表头样式
            Aspose.Cells.Style style = new Aspose.Cells.Style();
            style.Font.IsBold = true;
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.Pattern = BackgroundType.Solid;
            //设置表头单元格背景颜色
            style.ForegroundColor = System.Drawing.Color.FromArgb(166, 166, 166);
            Cells cells = sheet.Cells;//单元格
            string columnName = string.Empty;
            string rowName = string.Empty;
            string taxRat = string.Empty;
            string createDate = string.Empty;
            var colIndex = "A";
            try
            {
                //调用数据集
                #region
                string s = sqlstr;
                DataSet ds = SQLServer.Query(sqlstr,CustomConnect);
                if (ds.Tables.Count > 0)
                {
                    for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                    {
                        columnName = ds.Tables[0].Columns[i].ColumnName;
                        ////修改EXCEL列标题
                        if (ds.Tables[0].Columns[i].ColumnName == "BCR_NAME")
                        {
                            columnName = "条码阅读器编号";
                        }
                        if (ds.Tables[0].Columns[i].ColumnName == "Read_Count")
                        {
                            columnName = "识读条码数";
                        }
                        if (ds.Tables[0].Columns[i].ColumnName == "Not_Read_count")
                        {
                            columnName = "未识读条码";
                        }
                        if (ds.Tables[0].Columns[i].ColumnName == "PENCENT")
                        {
                            columnName = "条码识读率(%)";
                        }
                        if (ds.Tables[0].Columns[i].ColumnName == "insert_date")
                        {
                            columnName = "时间";
                        }
                        if (ds.Tables[0].Columns[i].ColumnName == "palno")
                        {
                            columnName = "条码号";
                        }
                        if (ds.Tables[0].Columns[i].ColumnName == "Length")
                        {
                            columnName = "长";
                        }
                        if (ds.Tables[0].Columns[i].ColumnName == "width")
                        {
                            columnName = "宽";
                        }
                        if (ds.Tables[0].Columns[i].ColumnName == "height")
                        {
                            columnName = "高";
                        }
                        if (ds.Tables[0].Columns[i].ColumnName == "Mode")
                        {
                            columnName = "母器具型号";
                        }
                        if (ds.Tables[0].Columns[i].ColumnName == "port")
                        {
                            columnName = "信息点";
                        }
                        if (ds.Tables[0].Columns[i].ColumnName == "num")
                        {
                            columnName = "目的地";
                        }
                        if (ds.Tables[0].Columns[i].ColumnName == "tkdat")
                        {
                            columnName = "更新时间";
                        }
                        if (ds.Tables[0].Columns[i].ColumnName == "tkdat")
                        {
                            columnName = "更新时间";
                        }
                        if (ds.Tables[0].Columns[i].ColumnName == "errormsg")
                        {
                            columnName = "错误信息";
                        }

                        sheet.Cells[colIndex + 1].PutValue((object)columnName);
                        sheet.Cells[colIndex + 1].SetStyle(style);
                        colIndex = ((char)(colIndex[0] + 1)).ToString();
                        for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
                        {
                            DataRow rows = ds.Tables[0].Rows[j];
                            cells[j + 1, i].PutValue(rows[ds.Tables[0].Columns[i].ColumnName].ToString());
                        }
                    }
                    #endregion
                    //大纲显示（折叠）
                    #region 设置单元格样式
                    cells.SetColumnWidth(0, 16);//设置列宽
                    cells.SetColumnWidth(1, 16);
                    cells.SetColumnWidth(2, 16);
                    cells.SetColumnWidth(3, 16);//设置列宽
                    cells.SetColumnWidth(4, 24);
                    //cells.SetColumnWidth(5, 20);
                    //cells.SetColumnWidth(6, 30);//设置列宽
                    //cells.SetColumnWidth(7, 20);
                    //cells.SetColumnWidth(8, 20);
                    //cells.SetColumnWidth(9, 30);//设置列宽
                    //cells.SetColumnWidth(10, 20);
                    //cells.SetColumnWidth(11, 20);
                    //cells.SetColumnWidth(12, 30);//设置列宽
                    //cells.SetColumnWidth(13, 20);
                    //cells.SetColumnWidth(14, 20);
                    //cells.SetColumnWidth(15, 30);//设置列宽
                    //cells.SetColumnWidth(16, 30);
                    //cells.SetColumnWidth(17, 30);
                    //cells.SetColumnWidth(16, 30);
                    //cells.SetColumnWidth(17, 30);
                    //cells.SetColumnWidth(18, 30);
                    //cells.SetColumnWidth(19, 30);
                    #endregion
                    string ExcelFileName = SaveExcelName();
                    if (ExcelFileName.Trim() == "")
                    { return; }
                    workbook.Save(ExcelFileName);
                    MessageBox.Show("导出完成！！", "提示");
                    //workbook.Save("d:\\Export\\ss.xls");
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        private string SaveExcelName()
        {
            //申明保存对话框
            SaveFileDialog dlg = new SaveFileDialog();
            //默然文件后缀
            dlg.DefaultExt = "xls";
            //文件后缀列表
            dlg.Filter = "EXCEL文件(*.XLS)|*.xls";
            //默然路径是系统当前路径
            dlg.InitialDirectory = Directory.GetCurrentDirectory();
            //打开保存对话框
            dlg.ShowDialog();
            //返回文件路径
            return dlg.FileName;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            ExportInvoiceExcel(sql_BCR);
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            timer2.Enabled = false;
          
            //Thread.Sleep(500);

            //try
            //{
            //    string sql = "select  top 1 Use_PC_Prepared from Choose_PC  order by update_time  ";
            //    DataSet ds_sql = new DataSet();
            //    ds_sql = SQLServer.Query(sql, CustomConnect);
            //    if (ds_sql.Tables.Count > 0 && ds_sql.Tables[0].Rows.Count > 0)
            //    {
            //        can_use_pc_prepared = ds_sql.Tables[0].Rows[0][0].ToString();
            //    }
            //}
            //catch
            //{
            //    can_use_pc_prepared = "";
            //}
            //Thread.Sleep(1200);
            //if ((can_use_pc_prepared.Equals("2") && SC_PC.Equals("2")) || (SC_PC.Equals("1")))
            //{
            //   // timer1.Enabled = true;
            //    Status.Text = "程序状态:运行";

            //    Status.ForeColor = Color.Green;
            //    timer1.Enabled = true;
            //    timer3.Enabled = true;
            //}
            //else
            //{
            //    Status.Text = "程序状态：待机";
            //    Status.ForeColor = Color.Red;
            //}

            //string sql_updete = "";
            //if (SC_PC.Equals("1"))
            //{
            //    sql_updete = "update Choose_PC set Use_PC_Prepared ='1'";

            //}
            //else
            //{
            //    sql_updete = "update Choose_PC set Use_PC_Prepared ='2'";


            //}

            //int aa = SQLServer.ExecuteSql(sql_updete, CustomConnect);


            timer2.Enabled = true;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            ExportInvoiceExcel(sql_ERR);
        }

        private void button14_Click(object sender, EventArgs e)
        {
            ExportInvoiceExcel(sql_Log);
        }

      


        private void button1_Click_1(object sender, EventArgs e)
        {
            if (MessageBox.Show("是否确定关闭程序？", "关闭程序", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                
               
             
          
                System.Threading.Thread.Sleep(500);
                

                this.timer1.Enabled = false;
                this.timer1.Dispose();
                this.timer2.Enabled = false;
                this.timer2.Dispose();

              
                //try { Proc.Kill(); }
                //catch { ;}
                string hint = "当前用户：" + person;
                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "Close", hint, "txt");
                System.Environment.Exit(0);

            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("是否确定关闭程序？", "关闭程序", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
               
               
           
          
                System.Threading.Thread.Sleep(500);
                

                this.timer1.Enabled = false;
                this.timer1.Dispose();
                this.timer2.Enabled = false;
                this.timer2.Dispose();

                
                //try { Proc.Kill(); }
                //catch { ;}
                string hint = "当前用户："+person;
                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "Close" , hint, "txt");

                System.Environment.Exit(0);

            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("是否确定关闭程序？", "关闭程序", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
               
              
           
          
                System.Threading.Thread.Sleep(500);
               

                this.timer1.Enabled = false;
                this.timer1.Dispose();
                this.timer2.Enabled = false;
                this.timer2.Dispose();

                
                //try { Proc.Kill(); }
                //catch { ;}
                string hint = "当前用户：" + person;
                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "Close", hint, "txt");
                System.Environment.Exit(0);

            }
        }

      

       
       
      

        private void button18_Click(object sender, EventArgs e)
        {
            ExportInvoiceExcel(sql_Huo_Log);
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            if (plcManager.isConnected)
            {
                plc_label.ForeColor = Color.Green;
                plc_label.Text = "已连接";

            }
            else
            {
                plc_label.ForeColor = Color.Red;
                plc_label.Text = "连接失败,请检查设备连接！";


            }
            timer1.Enabled = true ;
        }
    }
}
