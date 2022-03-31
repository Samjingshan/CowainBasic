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
using S7.Net;
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
using PLC_Adapt_;



namespace Sharetek
{


    public partial class FrmMain : Form
    {

        public static PLCManager plcManager;

        public string Log_Path = "";

        public string Logger_str = "";
        Plc S71500_1;

        string plc_ip_1;



        //IPAddress ipAddr2;
        // IPEndPoint ipEndPoint2;

        int beginPosition = 0;



        /// <summary>
        /// 从CustomSocket读写操作的数组 
        /// </summary>
        byte[] customBytes = new byte[1024];



        struct T_port
        {
            public bool readcode;
            public Byte[] scode;
            public Byte[] SIZE;
            public int read_count;
            public int Not_read_count;

            public int read_count_last;
            public int Not_read_count_last;

        }
        T_port[] percode = new T_port[100];

        private bool read_count_tmp = false;
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


        private int Print_Count1 = 0;





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

            plc_ip_1 = PublicFunc.ReadXML("PLC.xml", "//Config/IPAddress", "value", PublicFunc.PathType.AppPath);
            Dictionary<string, string> configDict = new Dictionary<string, string>();
            string _plcmemoryfile = ConfigurationManager.AppSettings["PLCMemoryFileName"] ?? "PLCMemory_s7.txt";
            configDict.Add("PLCMemoryFileName", _plcmemoryfile);

            PLCAdapter adapter = null;

            string _adapter = ConfigurationManager.AppSettings["PLCAdapter"] ?? "FileAdapter";
            string _ip = ConfigurationManager.AppSettings["IP"] ?? "192.168.1.10";
            string _port = ConfigurationManager.AppSettings["Port"] ?? "6000";
            string _stockerfile = ConfigurationManager.AppSettings["PLC_ConfigFile"] ?? "SiemensS7.txt";
            string _otherPLCFile= ConfigurationManager.AppSettings["PLC_ConfigFile_Other "] ?? "Siemens_Info.txt ";
            string _resolution = ConfigurationManager.AppSettings["resolution"] ?? "500";
            string _s7BaseAddress = ConfigurationManager.AppSettings["SiemensS7BasedAddress"] ?? "D15";
            string _s7Offset = ConfigurationManager.AppSettings["SiemensS7Offset"] ?? "77";
            string _SiemensMode = ConfigurationManager.AppSettings["SiemensMode"] ?? "3";
            Log_Path = ConfigurationManager.AppSettings["Log_path"];

            configDict.Add("SiemensS7Offset", _s7Offset);
            configDict.Add("SiemensS7BasedAddress", _s7BaseAddress);
            configDict.Add("IP", _ip);
            configDict.Add("PLCAdapter", _adapter);
            configDict.Add("Port", _port);
            configDict.Add("PLCConfigFile", _stockerfile);
            configDict.Add("resolution", _resolution);
            configDict.Add("SiemensMode", _SiemensMode);
            configDict.Add("Log_path", Log_Path);

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
                                      adapter, _stockerfile);
            plcManager.Init();//将定义PLC交互信号的2个文件 里所有的字段放到字典，key 字段名 value:PLCItem    key 字段名 value:PLCAlarm
            plcManager.Connect();//PLCEvent绑定绑定方法，若键值有变化，则触发方法
            plcManager.StartReading();




            #endregion

            INIT();
            Task.Run(() => {
                try
                {
                    System.Threading.Thread.Sleep(1000);

                    DoWithPLC doWithPLC = new DoWithPLC(plcManager);


                }
                catch (Exception ex)
                {
                    //CoreHelper.Logger.Error(CoreHelper.GetExceptionInfo(ex).Message);
                    //CoreHelper.Logger.Error(CoreHelper.GetExceptionInfo(ex).StackTrace);
                }
            });
            connect_plc1();







        }

        public void INIT()//将PLC的配置参数加载到字典里
        {
            

         
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
            tabPage4.Parent = null;
            

           

            #region
            //使用自定义格式
            dateTimePicker1.CustomFormat = "yyyy/MM/dd HH:mm：ss";
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.Value = DateTime.Now.Date;
            dateTimePicker1.Value = DateTime.Now.Date.AddHours(00).AddMinutes(00).AddSeconds(00);
            dateTimePicker2.CustomFormat = "yyyy/MM/dd HH:mm：ss";
            //使用自定义格式
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.Value = DateTime.Now.Date;
            dateTimePicker2.Value = DateTime.Now.Date.AddHours(23).AddMinutes(59).AddSeconds(59);
            dateTimePicker4.CustomFormat = "yyyy/MM/dd HH:mm：ss";
            //使用自定义格式
            dateTimePicker4.Format = DateTimePickerFormat.Custom;
            dateTimePicker4.Value = DateTime.Now.Date;
            dateTimePicker4.Value = DateTime.Now.Date.AddHours(00).AddMinutes(00).AddSeconds(00);
            dateTimePicker3.CustomFormat = "yyyy/MM/dd HH:mm：ss";
            //使用自定义格式
            dateTimePicker3.Format = DateTimePickerFormat.Custom;
            dateTimePicker3.Value = DateTime.Now.Date;
            dateTimePicker3.Value = DateTime.Now.Date.AddHours(23).AddMinutes(59).AddSeconds(59);






            #endregion

            this.Text = Txt_title;
            username.Text = "当时用户：" + person;
            username.ForeColor = Color.Green;
            timer1.Enabled = false;
            timer2.Enabled = false;
            
          
           
         
          
           
        }
     

        private void connect_plc1()
        {
            #region 


            try
            {
                S71500_1 = new Plc(CpuType.S71500, plc_ip_1, 0, 1);
                S71500_1.Open();
                
                if (!S71500_1.IsConnected)
                {
                    //plc1_label.ForeColor = Color.Red;
                    //plc1_label.Text = "PLC_1连接状态：异常";
                    PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "connect_plc", "connect_plc:与PLC1连接异常", "txt");


                }
                else
                {
                    //plc1_label.ForeColor = Color.Green;
                    //plc1_label.Text = "PLC_1连接状态：正常";

                    PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "connect_plc", "connect_plc:与PLC1连接正常", "txt");

                }

  

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "connect_plc", "异常：connect_plc:" + ex.ToString(), "txt");              


            }
            #endregion



        }


      



        private void read_readcode(int code_type, int scan_no, int barcode_db, int strat_add)
        //code_type :1代表阅读器条码，2代表货到人条码,3代表分拣机前条码,4代表多穿入口条码,scan_no条码编号，scan_no条码编号，barcode_db条码所在DB块，strat_add条码起始地址
        {


            //1-14 读条码标志   1-14为条码阅读器，11-16为1号货到人，21-26为2号货到人
            try
            {
                if (code_type == 1)
                {
                    for (int i = 1; i < 15; i++)
                    {

                        percode[i].scode = (byte[])S71500_1.ReadBytes(DataType.DataBlock, barcode_db, 16 * (i - 1), 1);
                        if (percode[i].scode[0] == 1 || percode[i].scode[0] == 3 )
                        {

                            percode[i].readcode = true;
                        }
                        else
                        {
                            percode[i].readcode = false;
                        }
                    }

                }
                else if (code_type == 2)
                {
                   

                        percode[scan_no+10].scode = (byte[])S71500_1.ReadBytes(DataType.DataBlock, barcode_db, strat_add, 1);
                        if (percode[scan_no+10].scode[0] == 1 || percode[scan_no + 10].scode[0] == 3 )
                        {

                            percode[scan_no+10].readcode = true;
                        }
                        else
                        {
                            percode[scan_no+10].readcode = false;
                        }
                    

                }
                
                else if (code_type == 4)
                {


                    percode[scan_no + 15].scode = (byte[])S71500_1.ReadBytes(DataType.DataBlock, barcode_db, strat_add, 1);
                    if (percode[scan_no + 15].scode[0] == 1 || percode[scan_no + 15].scode[0] == 3 )
                    {

                        percode[scan_no + 15].readcode = true;
                    }
                    else
                    {
                        percode[scan_no + 15].readcode = false;
                    }


                }


            }
            catch
            {
                return;
            }





        }

        private string read_barcode(int code_type,int scan_no,int barcode_db,int strat_add)
        //code_type :1代表阅读器条码，2代表货到人条码,3代表分拣机前条码,4代表多穿入口条码,scan_no条码编号，barcode_db条码所在DB块，strat_add条码起始地址
        {
            string _barCode = "";
            if (code_type == 1)
            {

                percode[scan_no].scode = (byte[])S71500_1.ReadBytes(DataType.DataBlock, barcode_db , 30* (scan_no - 1), 30);//读取30个字节
                _barCode = System.Text.Encoding.Default.GetString(percode[scan_no].scode).Trim().Trim((char)0x00).Trim((char)0x02).Trim((char)0x03);//.Trim((char)0x02).Trim((char)0x03).Trim((char)0x0A).Trim((char)0x0D);

            }
            else if (code_type == 2)
            {
                percode[scan_no+10].scode = (byte[])S71500_1.ReadBytes(DataType.DataBlock, barcode_db, strat_add, 30);//读取30个字节
                _barCode = System.Text.Encoding.Default.GetString(percode[scan_no+10].scode).Trim().Trim((char)0x00).Trim((char)0x02).Trim((char)0x03);//.Trim((char)0x02).Trim((char)0x03).Trim((char)0x0A).Trim((char)0x0D);


            }
           
            if (code_type == 4)
            {

                percode[scan_no + 15].scode = (byte[])S71500_1.ReadBytes(DataType.DataBlock, barcode_db, strat_add, 30);//读取30个字节
                _barCode = System.Text.Encoding.Default.GetString(percode[scan_no+15].scode).Trim().Trim((char)0x00).Trim((char)0x02).Trim((char)0x03);//.Trim((char)0x02).Trim((char)0x03).Trim((char)0x0A).Trim((char)0x0D);

            }

           // _barCode = System.Text.Encoding.Default.GetString(percode[scan_no].scode).Trim().Trim((char)0x00).Trim((char)0x02).Trim((char)0x03);//.Trim((char)0x02).Trim((char)0x03).Trim((char)0x0A).Trim((char)0x0D);
            if (!string.IsNullOrEmpty(_barCode))
            {
                //_barCode = _barCode.Substring(1,_barCode.Length-1);
                _barCode = format_barcode(_barCode);  //以02开头03结尾的字符 
            }
            _barCode = _barCode.Replace(" ", "");
            _barCode = _barCode.Replace("\n", "");
            _barCode = _barCode.Replace("\r", "");
            _barCode = _barCode.Replace("\t", "");
            if (!(scan_no == 1 || scan_no == 15) && _barCode.Length > 8)//非原箱条码
            {
                _barCode = _barCode.Substring(0, 8);

            }
            //if (_barCode.Length > 20)
            //{

            //    _barCode = _barCode.Substring(0, 20);
            //}

            // if (_barCode.Length > 8 && scan_no==1)
            //{
            //   string  _barCode_tmp1 = _barCode.Substring(0, 8);
            //    string _barCode_tmp2 = _barCode.Substring(8, 8);
            //    if (_barCode_tmp1.Equals(_barCode_tmp2))
            //    {
            //        _barCode = _barCode_tmp1;
            //    }
            //}
            return _barCode;

        }

        private void read_Barcode_Count()
        {
           
            string data_time = DateTime.Now.ToLongTimeString();

            string[] arr_T = data_time.Split(':');
            string Min = arr_T[1];
            string Sec = arr_T[2];
            if (Min.Equals("00") && Sec.Equals("00") && !read_count_tmp )

            {
                read_count_tmp = true;
                //若读到书和未读到数和上个扫描周期不一样，那么往数据库插入一条数据
                for (int i = 0; i < 10; i++)
                {
                    percode[i].read_count = Convert.ToInt16(S71500_1.Read("db2.dbw" + (10 + 18 * i).ToString()));//读出数
                    percode[i].Not_read_count = Convert.ToInt16(S71500_1.Read("db2.dbw" + (12 + 18 * i).ToString()));//未读出数
                    string pencent = "0";
                    try
                    {
                        pencent = (100 * percode[i].read_count / (percode[i].read_count + percode[i].Not_read_count)).ToString();
                    }
                    catch
                    {
                        pencent = "0";
                    }

                    if ((percode[i].read_count == percode[i].read_count_last) && (percode[i].Not_read_count == percode[i].Not_read_count_last))
                    {
                        continue;
                    }
                    else
                    {
                        string sql = "insert into BCR_COUNT (BCR_NO,BCR_NAME,Read_Count,Not_Read_count,PENCENT,insert_date)" +
                    "values(" + (i + 1) + ", 'BCR" + (i + 1) + "', " + percode[i].read_count + ", " + percode[i].Not_read_count + ", '" + pencent + "', GETDATE())";
                    int aaa = SQLServer.ExecuteSql(sql, CustomConnect);
                    if (aaa > 0)
                    {
                            percode[i].read_count_last = percode[i].read_count;
                            percode[i].Not_read_count_last = percode[i].Not_read_count;
                            PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "条码识读率", i + "号识读率更新成功", "txt");
                    }
                    else
                    {
                        PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "条码识读率", i + "号识读率更新失败，请检查网络连接", "txt");
                    }


                    }

                }
                read_count_tmp = false;
            }
         

        }

        private void button2_Click(object sender, EventArgs e)
        {
          
            if (MessageBox.Show("是否确定关闭程序？", "关闭程序", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
               // myResetEvent1.Reset();
               
                S71500_1.Close();
              
                System.Threading.Thread.Sleep(500);
               

                this.timer1.Enabled = false;
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
      
      
        private void do_check_BCR1(int scan_no, int read_code_DB, int read_code_Add, int Weight_Add)
        {

            string _barCode = "";
            string _Weight = "";
            string Area = "";
            try
            {
                DateTime dtStar_1 = DateTime.Now;


                //先取到条码,再更新到wcs_to_ccs表   

                //_barCode = read_barcode(1, scan_no, barcode_DB, 0);

                int Route = 0;
                string ERR_String = "";
                string Route_tmp = "";
                string box_num = "";
                double Weight_tmp = 0;


                //percode[scan_no].SIZE = (byte[])S71500_1.ReadBytes(DataType.DataBlock, barcode_DB, Weight_Add, 2);//一个int 占2个字节，4个int就是4*2=8个字节

                Byte_Arr_to_Int(0, percode[scan_no].SIZE, ref _Weight);



                string sql_select = "select * from OPENQUERY(REFINE, 'select  Area1, box_num  from WMS_TO_CCS where box_num=''" + _barCode + "''  or  Lpn=''" + _barCode+"'' order by  UpDate_Time desc limit 0,1')";


                #region 获取路向，判断异常


                try
                {
                    if (!(string.IsNullOrEmpty(_barCode) || _barCode.Equals("NoRead")))
                    {


                        DataSet ds = SQLServer.Query(sql_select, CustomConnect);
                      
                            #region  判断是否WMS有记录，无则写异常信息
                            if (ds.Tables[0].Rows.Count > 0 && ds.Tables.Count > 0) //wms_to_ccs有记录
                            {
                                #region 判断路向
                                try
                                {
                                    Route_tmp = ds.Tables[0].Rows[0][0].ToString();
                                    box_num = ds.Tables[0].Rows[0][1].ToString();
                                 
                                    if (!string.IsNullOrEmpty(Route_tmp))
                                    {
                                        if (Route_tmp.Equals("0"))
                                        {
                                       
                                            Route = 1;
                                            Area = "分拣机区域";                                       
                                        }
                                        else if (Route_tmp.Equals("1"))

                                        {
                                        if (Weight_tmp > 60)//超重
                                        {
                                            Route = 8;
                                            Area = "换箱区域";
                                        }
                                        else
                                        {
                                            Route = 2;
                                            Area = "换箱区域";
                                        }
                                           
                                           
                                        }
                                        else
                                        {

                                            Route = 9;
                                            Area = "换箱区域";
                                            ERR_String = "地址绑定错误";                                                                                        
                                          
                                        }

                                    }
                                    else
                                    {
                                        Route = 9;
                                        Area = "换箱区域";
                                        ERR_String ="地址未绑定";
                                        
                                           
                                        
                                }


                                }
                                catch (Exception ex)
                                {
                                
                                    Route = 9;
                                    Area = "换箱区域";
                                    ERR_String = ex.Message;
                               
                            }
                                #endregion


                            }
                            else  //wms_to_ccs没有记录
                            {
                                Area = "换箱区域";
                                Route = 9;
                                ERR_String = "未查到该条码任务";
                            
                           }
                            #endregion

                        



                    }
                    else //条码为空，当异常处理
                    {
                        Route = 9;
                        ERR_String = "条码未读取";                       
                        Area = "换箱区域";
                       
                    }
                }
                catch (Exception ex)
                {
                    Route = 9;
                    ERR_String = ex.Message;                    
                    Area = "换箱区域";
                 
                }
                #endregion


                Deal_with_Log(_barCode, "BH" + scan_no, Route.ToString(), ERR_String, CustomConnect);//流水记录
                if (Area.Equals("换箱区域"))
                Deal_with_CCS_to_WMS_Log(_barCode, "Null", 1, CustomConnect);//往CCS_TO_WMS插流水记录，并复位wms_to_ccs地址


                if (!string.IsNullOrEmpty(ERR_String))
                {
                    if (!(ERR_String.Equals("条码未读取") || ERR_String.Equals("未查到该条码任务") || ERR_String.Equals("地址未绑定") || ERR_String.Equals("地址绑定错误")))
                    {
                        ERR_String = "BH" + scan_no + "处系统异常";
                    }


                    Deal_with_ERRINFO(_barCode, "BH" + scan_no, ERR_String, CustomConnect);
                    PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "ERROR_BH" + scan_no, ERR_String, "txt");

                }


                Byte[] write_stackinfo = new byte[4];
                write_stackinfo[0] = Convert.ToByte(2); //清READCODE  置READDONE 0000 0100
                write_stackinfo[1] = Convert.ToByte(0);
                write_stackinfo[2] = Convert.ToByte(Route / 256);
                write_stackinfo[3] = Convert.ToByte(Route % 256);
                S71500_1.WriteBytes(DataType.DataBlock, read_code_DB, read_code_Add, write_stackinfo);


                DateTime dtEnd = DateTime.Now;
                double dSub = SubDateTimeByMill(dtStar_1, dtEnd) / 1000;
                string hint = "扫描器=BCR" + scan_no + "，条码=" + _barCode + "，Route=" + Route + ",目的地 = " + Area + ",时间：" + dtEnd + "，耗时:" + dSub.ToString() + "秒" + " " + "\n";
                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "do_check_BCR" + scan_no, hint, "txt");
                PaintColorToRtxtbx(richTextBox1, hint, ref beginPosition, ref Print_Count1);


            }
            catch (Exception ex)
            {

                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "ERROR_BH" + scan_no, _barCode + ",发生异常:" + ex.ToString(), "txt");

                return;
            }

        }

        private void do_check_Area2(int scan_no, int read_code_DB, int read_code_Add)
        {

            string _barCode = "";
            string Area = "";
            try
            {
                DateTime dtStar_1 = DateTime.Now;


                //先取到条码,再更新到wcs_to_ccs表   

                //_barCode = read_barcode(1, scan_no, barcode_DB, 0);

                int Route = 0;
                string ERR_String = "";
                string Route_tmp = "";
               
                string Lpn = "";
               





                string sql_select = "select * from OPENQUERY(REFINE, 'select  Area2,Lpn  from WMS_TO_CCS where  Lpn=''" + _barCode + "'' order by  UpDate_Time desc limit 0,1')";


                #region 获取路向，判断异常


                try
                {
                    if (!(string.IsNullOrEmpty(_barCode) || _barCode.Equals("NoRead")))
                    {


                        DataSet ds = SQLServer.Query(sql_select, CustomConnect);
                       
                            #region  判断是否WMS有记录，无则写异常信息
                            if (ds.Tables[0].Rows.Count > 0 && ds.Tables.Count > 0) //wms_to_ccs有记录
                            {
                                #region 判断路向
                                try
                                {
                                    Route_tmp = ds.Tables[0].Rows[0][0].ToString();
                                    Lpn = ds.Tables[0].Rows[0][1].ToString();

                                    if (!string.IsNullOrEmpty(Route_tmp))
                                    {
                                        if (Route_tmp.Equals("0"))
                                        {
                                            Route = 2;
                                            Area = "直行";
                                        }
                                        else if (Route_tmp.Equals("1"))

                                        {
                                            Route = 1;
                                            Area = "提升机区域";
                                           
                                        }
                                        else
                                        {
                                            Route = 2;
                                            Area = "直行";
                                            ERR_String = "地址绑定错误";
                                        }

                                    }
                                    else
                                    {
                                        Route = 2;
                                        Area = "直行";
                                        ERR_String = "地址未绑定";
                                    }


                                }
                                catch (Exception ex)
                                {
                                    Route = 2;
                                    Area = "直行";
                                    ERR_String = ex.Message;
                                }
                                #endregion


                            }
                            else  //wms_to_ccs没有记录
                            {
                                ERR_String = "未查到该条码任务";
                                Route = 2;
                                Area = "直行";
                        }
                            #endregion

                        



                    }
                    else //条码为空，当异常处理
                    {
                        ERR_String = "条码未读取";
                        Route = 2;
                        Area = "直行";
                    }
                }
                catch (Exception ex)
                {
                    ERR_String = ex.Message;
                    Route = 2;
                    Area = "直行";
                }
                #endregion

              
                Deal_with_Log(_barCode, "BH"+scan_no , Route.ToString(), ERR_String, CustomConnect);//流水记录
               

                if (!string.IsNullOrEmpty(ERR_String))
                {
                    if (!(ERR_String.Equals("条码未读取") || ERR_String.Equals("未查到该条码任务") || ERR_String.Equals("地址未绑定") || ERR_String.Equals("地址绑定错误")))
                    {
                       
                            ERR_String = "BH"+scan_no + "处系统异常";
                    }


                   

                        Deal_with_ERRINFO(_barCode, "BH"+scan_no, ERR_String, CustomConnect);
                        PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "ERROR_BH"+scan_no, ERR_String, "txt");
                     
                  

                }


                Byte[] write_stackinfo = new byte[4];
                write_stackinfo[0] = Convert.ToByte(2); //清READCODE  置READDONE 0000 0100
                write_stackinfo[1] = Convert.ToByte(0);
                write_stackinfo[2] = Convert.ToByte(Route / 256);
                write_stackinfo[3] = Convert.ToByte(Route % 256);
                S71500_1.WriteBytes(DataType.DataBlock, read_code_DB, read_code_Add, write_stackinfo);


                DateTime dtEnd = DateTime.Now;
                double dSub = SubDateTimeByMill(dtStar_1, dtEnd) / 1000;
                string hint = "";
               
                    hint = "扫描器=BCR"+scan_no  + "，条码=" + _barCode + "，Route=" + Route + ",目的地 = " + Area + ",时间：" + dtEnd + "，耗时:" + dSub.ToString() + "秒" + " " + "\n";
                    PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "do_check_BCR"+scan_no , hint, "txt");
                
               
                PaintColorToRtxtbx(richTextBox1, hint, ref beginPosition, ref Print_Count1);


            }
            catch (Exception ex)
            {

                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "ERROR_BH" + scan_no, _barCode + ",发生异常:" + ex.ToString(), "txt");

                return;
            }

        }


        private void DC_In(int scan_no, int read_code_DB, int read_code_Add)
        {//功能：查询wcs_to_ccs表，查找路向信息 入库的时候记个日志
            string _barCode = "";
            string AREA = "DC_IN_" + (scan_no );

            try
            {
                //scan_no代表扫描器编号    
                DateTime dtStar = DateTime.Now;
                //if (scan_no == 1)
                //     _barCode = read_barcode(4, scan_no, barcode_DB, barcode_DC_1);
                //else if (scan_no == 2)
                //    _barCode = read_barcode(4, scan_no, barcode_DB, barcode_DC_2);
                Deal_with_Log(_barCode, AREA, "", "", CustomConnect);//流水记录
                Deal_with_CCS_to_WMS_Log("Null", _barCode, 2, CustomConnect);//往CCS_TO_WMS插流水记录，并复位wms_to_ccs地址

                Byte[] write_stackinfo = new byte[1];
                write_stackinfo[0] = Convert.ToByte(2); //清READCODE  置READDONE 0000 0100
                S71500_1.WriteBytes(DataType.DataBlock, read_code_DB, read_code_Add, write_stackinfo);

                DateTime dtEnd = DateTime.Now;
                double dSub = SubDateTimeByMill(dtStar, dtEnd) / 1000;
                string hint = "区域=" + AREA + "，条码=" + _barCode + ",时间：" + dtEnd + "，耗时:" + dSub.ToString() + "秒" + " " + "\n";
                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + AREA , hint, "txt");
                PaintColorToRtxtbx(richTextBox1, hint, ref beginPosition, ref Print_Count1);





            }
            catch (Exception ex)

            {
                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + AREA + "_ERR" , _barCode + ",do_check_area:" + ex.ToString(), "txt");
               
                return;
            }


        }
        private void do_check_BCR6(int scan_no, int read_code_DB, int read_code_Add)
        {

            string _barCode = "";
            string Area = "";
            try
            {
                DateTime dtStar_1 = DateTime.Now;


                //先取到条码,再更新到wcs_to_ccs表   

                //_barCode = read_barcode(1, scan_no, barcode_DB, 0);

                int Route = 0;
                string ERR_String = "";
                string Route_tmp = "";
                string box_num = "";
               





                string sql_select = "select * from OPENQUERY(REFINE, 'select  Area1, box_num  from WMS_TO_CCS where box_num=''" + _barCode + "''  or  Lpn=''" + _barCode + "'' order by  UpDate_Time desc limit 0,1')";


                #region 获取路向，判断异常


                try
                {
                    if (!(string.IsNullOrEmpty(_barCode) || _barCode.Equals("NoRead")))
                    {


                        DataSet ds = SQLServer.Query(sql_select, CustomConnect);

                        #region  判断是否WMS有记录，无则写异常信息
                        if (ds.Tables[0].Rows.Count > 0 && ds.Tables.Count > 0) //wms_to_ccs有记录
                        {
                            #region 判断路向
                            try
                            {
                                Route_tmp = ds.Tables[0].Rows[0][0].ToString();
                                box_num = ds.Tables[0].Rows[0][1].ToString();

                                if (!string.IsNullOrEmpty(Route_tmp))
                                {
                                    if (Route_tmp.Equals("0"))
                                    {
                                        
                                            Route = 2;
                                            Area = "直行";
                                       


                                    }
                                    else if (Route_tmp.Equals("1"))

                                    {
                                        Route = 1;
                                        Area = "换箱区域";
                                       
                                    }
                                    else
                                    {

                                        Route = 2;
                                        ERR_String = "地址绑定错误";
                                        
                                            Area = "直行";
                                    }

                                }
                                else
                                {
                                    Route = 2;
                                    ERR_String = "地址未绑定";
                                   
                                        Area = "直行";
                                }


                            }
                            catch (Exception ex)
                            {
                               
                                    Route = 2;
                                    Area = "直行";
                                ERR_String = ex.Message;
                                
                            }
                            #endregion


                        }
                        else  //wms_to_ccs没有记录
                        {
                            Route = 2;
                            ERR_String = "未查到该条码任务";
                            
                                Area = "直行";


                        }
                        #endregion





                    }
                    else //条码为空，当异常处理
                    {
                        Route = 2;
                       
                            Area = "直行";
                    }
                }
                catch (Exception ex)
                {
                    Route = 2;
                    ERR_String = ex.Message;
                    
                        Area = "直行";
                }
                #endregion


                Deal_with_Log(_barCode, "BH" + scan_no, Route.ToString(), ERR_String, CustomConnect);//流水记录
                if (Area.Equals("换箱区域"))
                    Deal_with_CCS_to_WMS_Log("Null", _barCode, 1, CustomConnect);//往CCS_TO_WMS插流水记录，并复位wms_to_ccs地址


                if (!string.IsNullOrEmpty(ERR_String))
                {
                    if (!(ERR_String.Equals("条码未读取") || ERR_String.Equals("未查到该条码任务") || ERR_String.Equals("地址未绑定") || ERR_String.Equals("地址绑定错误")))
                    {
                        ERR_String = "BH" + scan_no + "处系统异常";
                    }


                    Deal_with_ERRINFO(_barCode, "BH" + scan_no, ERR_String, CustomConnect);
                    PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "ERROR_BH" + scan_no, ERR_String, "txt");

                }


                Byte[] write_stackinfo = new byte[4];
                write_stackinfo[0] = Convert.ToByte(2); //清READCODE  置READDONE 0000 0100
                write_stackinfo[1] = Convert.ToByte(0);
                write_stackinfo[2] = Convert.ToByte(Route / 256);
                write_stackinfo[3] = Convert.ToByte(Route % 256);
                S71500_1.WriteBytes(DataType.DataBlock, read_code_DB, read_code_Add, write_stackinfo);


                DateTime dtEnd = DateTime.Now;
                double dSub = SubDateTimeByMill(dtStar_1, dtEnd) / 1000;
                string hint = "扫描器=BCR" + scan_no + "，条码=" + _barCode + "，Route=" + Route + ",目的地 = " + Area + ",时间：" + dtEnd + "，耗时:" + dSub.ToString() + "秒" + " " + "\n";
                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "do_check_BCR" + scan_no, hint, "txt");
                PaintColorToRtxtbx(richTextBox1, hint, ref beginPosition, ref Print_Count1);


            }
            catch (Exception ex)
            {

                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "ERROR_BH" + scan_no, _barCode + ",发生异常:" + ex.ToString(), "txt");

                return;
            }

        }

        private void do_check_BCR10(int scan_no, int read_code_DB, int read_code_Add)
        {

            string _barCode = "";
            string Area = "";
            try
            {
                DateTime dtStar_1 = DateTime.Now;


                //先取到条码,再更新到wcs_to_ccs表   

                //_barCode = read_barcode(1, scan_no, barcode_DB, 0);

                int Route = 0;
                string ERR_String = "";
                string Route_tmp_1 = "";
                string Route_tmp_2 = "";
   





                string sql_select = "select * from OPENQUERY(REFINE, 'select  Area10,Area11  from WMS_TO_CCS where box_num=''" + _barCode + "''  or  Lpn=''" + _barCode + "'' order by  UpDate_Time desc limit 0,1')";


                #region 获取路向，判断异常


                try
                {
                    if (!(string.IsNullOrEmpty(_barCode) || _barCode.Equals("NoRead")))
                    {


                        DataSet ds = SQLServer.Query(sql_select, CustomConnect);

                        #region  判断是否WMS有记录，无则写异常信息
                        if (ds.Tables[0].Rows.Count > 0 && ds.Tables.Count > 0) //wms_to_ccs有记录
                        {
                            #region 判断路向
                            try
                            {
                                Route_tmp_1 = ds.Tables[0].Rows[0][0].ToString();
                                Route_tmp_2 = ds.Tables[0].Rows[0][1].ToString();

                                if (!(string.IsNullOrEmpty(Route_tmp_1)&& string.IsNullOrEmpty(Route_tmp_2)))
                                {
                                    
                                    if (Route_tmp_1.Equals("1")|| Route_tmp_2.Equals("1"))

                                    {
                                        Route = 1;
                                        Area = "复核区域或者立库区域";

                                    }
                                    else
                                    {

                                        Route = 2;
                                        

                                        Area = "直行";
                                    }

                                }
                                else
                                {
                                    Route = 2;
                                    ERR_String = "地址未绑定";

                                    Area = "直行";
                                }


                            }
                            catch (Exception ex)
                            {

                                Route = 2;
                                Area = "直行";
                                ERR_String = ex.Message;

                            }
                            #endregion


                        }
                        else  //wms_to_ccs没有记录
                        {
                            Route = 2;
                            ERR_String = "未查到该条码任务";

                            Area = "直行";


                        }
                        #endregion





                    }
                    else //条码为空，当异常处理
                    {
                        Route = 2;
                        ERR_String = "条码未读取";
                        Area = "直行";
                    }
                }
                catch (Exception ex)
                {
                    Route = 2;
                    ERR_String = ex.Message;

                    Area = "直行";
                }
                #endregion


                Deal_with_Log(_barCode, "BH" + scan_no, Route.ToString(), ERR_String, CustomConnect);//流水记录
                //Deal_with_CCS_to_WMS_Log("Null", Lpn, scan_no, CustomConnect);//往CCS_TO_WMS插流水记录，并复位wms_to_ccs地址


                if (!string.IsNullOrEmpty(ERR_String))
                {
                    if (!(ERR_String.Equals("条码未读取") || ERR_String.Equals("未查到该条码任务") || ERR_String.Equals("地址未绑定") || ERR_String.Equals("地址绑定错误")))
                    {
                        ERR_String = "BH" + scan_no + "处系统异常";
                    }


                    Deal_with_ERRINFO(_barCode, "BH" + scan_no, ERR_String, CustomConnect);
                    PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "ERROR_BH" + scan_no, ERR_String, "txt");

                }


                Byte[] write_stackinfo = new byte[4];
                write_stackinfo[0] = Convert.ToByte(2); //清READCODE  置READDONE 0000 0100
                write_stackinfo[1] = Convert.ToByte(0);
                write_stackinfo[2] = Convert.ToByte(Route / 256);
                write_stackinfo[3] = Convert.ToByte(Route % 256);
                S71500_1.WriteBytes(DataType.DataBlock, read_code_DB, read_code_Add, write_stackinfo);


                DateTime dtEnd = DateTime.Now;
                double dSub = SubDateTimeByMill(dtStar_1, dtEnd) / 1000;
                string hint = "扫描器=BCR" + scan_no + "，条码=" + _barCode + "，Route=" + Route + ",目的地 = " + Area + ",时间：" + dtEnd + "，耗时:" + dSub.ToString() + "秒" + " " + "\n";
                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "do_check_BCR" + scan_no, hint, "txt");
                PaintColorToRtxtbx(richTextBox1, hint, ref beginPosition, ref Print_Count1);


            }
            catch (Exception ex)
            {

                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "ERROR_BH" + scan_no, _barCode + ",发生异常:" + ex.ToString(), "txt");

                return;
            }

        }

        private void do_check_BCR9(int scan_no, int read_code_DB, int read_code_Add)
        {

            string _barCode = "";
            string Area = "";
            try
            {
                DateTime dtStar_1 = DateTime.Now;


                //先取到条码,再更新到wcs_to_ccs表   

                //_barCode = read_barcode(1, scan_no, barcode_DB, 0);

                int Route = 0;
                string ERR_String = "";
                string Route_tmp_1 = "";
                string Route_tmp_2 = "";






                string sql_select = "select * from OPENQUERY(REFINE, 'select  Area10,Area11  from WMS_TO_CCS where box_num=''" + _barCode + "''  or  Lpn=''" + _barCode + "'' order by  UpDate_Time desc limit 0,1')";


                #region 获取路向，判断异常


                try
                {
                    if (!(string.IsNullOrEmpty(_barCode) || _barCode.Equals("NoRead")))
                    {


                        DataSet ds = SQLServer.Query(sql_select, CustomConnect);

                        #region  判断是否WMS有记录，无则写异常信息
                        if (ds.Tables[0].Rows.Count > 0 && ds.Tables.Count > 0) //wms_to_ccs有记录
                        {
                            #region 判断路向
                            try
                            {
                                Route_tmp_1 = ds.Tables[0].Rows[0][0].ToString();
                                Route_tmp_2 = ds.Tables[0].Rows[0][1].ToString();

                                if (!(string.IsNullOrEmpty(Route_tmp_1) && string.IsNullOrEmpty(Route_tmp_2)))
                                {

                                    if ( Route_tmp_2.Equals("1"))

                                    {
                                        Route = 2;
                                        Area = "立库区域";

                                    }
                                    else
                                    {

                                        Route = 1;
                                        Area = "复核区域";
                                    }

                                }
                                else
                                {
                                    Route = 1;
                                    Area = "复核区域";
                                    ERR_String = "地址未绑定";

                                   
                                }


                            }
                            catch (Exception ex)
                            {

                                Route = 1;
                                Area = "复核区域";
                                ERR_String = ex.Message;

                            }
                            #endregion


                        }
                        else  //wms_to_ccs没有记录
                        {
                            Route = 1;
                            Area = "复核区域";
                            ERR_String = "未查到该条码任务";

                           


                        }
                        #endregion





                    }
                    else //条码为空，当异常处理
                    {
                        Route = 1;
                        Area = "复核区域";
                        ERR_String = "条码未读取";
                        
                    }
                }
                catch (Exception ex)
                {
                    Route = 1;
                    Area = "复核区域";
                    ERR_String = ex.Message;

                   
                }
                #endregion


                Deal_with_Log(_barCode, "BH" + scan_no, Route.ToString(), ERR_String, CustomConnect);//流水记录
                if (Route == 1)

                    Deal_with_CCS_to_WMS_Log("Null", _barCode, 10, CustomConnect);//往CCS_TO_WMS插流水记录，并复位wms_to_ccs地址
                else if (Route == 2)

                    Deal_with_CCS_to_WMS_Log("Null", _barCode, 11, CustomConnect);//往CCS_TO_WMS插流水记录，并复位wms_to_ccs地址


                if (!string.IsNullOrEmpty(ERR_String))
                {
                    if (!(ERR_String.Equals("条码未读取") || ERR_String.Equals("未查到该条码任务") || ERR_String.Equals("地址未绑定") || ERR_String.Equals("地址绑定错误")))
                    {
                        ERR_String = "BH" + scan_no + "处系统异常";
                    }


                    Deal_with_ERRINFO(_barCode, "BH" + scan_no, ERR_String, CustomConnect);
                    PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "ERROR_BH" + scan_no, ERR_String, "txt");

                }


                Byte[] write_stackinfo = new byte[4];
                write_stackinfo[0] = Convert.ToByte(2); //清READCODE  置READDONE 0000 0100
                write_stackinfo[1] = Convert.ToByte(0);
                write_stackinfo[2] = Convert.ToByte(Route / 256);
                write_stackinfo[3] = Convert.ToByte(Route % 256);
                S71500_1.WriteBytes(DataType.DataBlock, read_code_DB, read_code_Add, write_stackinfo);


                DateTime dtEnd = DateTime.Now;
                double dSub = SubDateTimeByMill(dtStar_1, dtEnd) / 1000;
                string hint = "扫描器=BCR" + scan_no + "，条码=" + _barCode + "，Route=" + Route + ",目的地 = " + Area + ",时间：" + dtEnd + "，耗时:" + dSub.ToString() + "秒" + " " + "\n";
                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "do_check_BCR" + scan_no, hint, "txt");
                PaintColorToRtxtbx(richTextBox1, hint, ref beginPosition, ref Print_Count1);


            }
            catch (Exception ex)
            {

                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "ERROR_BH" + scan_no, _barCode + ",发生异常:" + ex.ToString(), "txt");

                return;
            }

        }


        

        private void do_check_Area_Huo_up(int scan_no, int AREA,int read_code_DB, int read_code_Add)
        {

            string _barCode = "";
            string Area = "Area" + AREA;
            try
            {
                DateTime dtStar_1 = DateTime.Now;


                //先取到条码,再更新到wcs_to_ccs表   

                //_barCode = read_barcode(1, scan_no, barcode_DB, 0);
                //_barCode = "1234567890";


                int Route = 0;
                string ERR_String = "";
                string Route_tmp = "";

                string Lpn = "";


                string sql_select = "select * from OPENQUERY(REFINE, 'select  "+ Area + ", Lpn  from WMS_TO_CCS where  Lpn=''" + _barCode + "'' order by  UpDate_Time desc limit 0,1')";


                #region 获取路向，判断异常


                try
                {
                    if (!(string.IsNullOrEmpty(_barCode) || _barCode.Equals("NoRead")))
                    {


                        DataSet ds = SQLServer.Query(sql_select, CustomConnect);

                        #region  判断是否WMS有记录，无则写异常信息
                        if (ds.Tables[0].Rows.Count > 0 && ds.Tables.Count > 0) //wms_to_ccs有记录
                        {
                            #region 判断路向
                            try
                            {
                                Route_tmp = ds.Tables[0].Rows[0][0].ToString();
                                Lpn = ds.Tables[0].Rows[0][1].ToString();

                                if (!string.IsNullOrEmpty(Route_tmp))
                                {
                                    if (Route_tmp.Equals("0"))
                                    {
                                        Route = 2;
                                        Area = "直行";
                                    }
                                    else if (Route_tmp.Equals("1"))

                                    {
                                        Route = 1;
                                        Area = "货到人区域";

                                    }
                                    else
                                    {
                                        Route = 2;
                                        Area = "直行";
                                        ERR_String = "地址绑定错误";
                                    }

                                }
                                else
                                {
                                    Route = 2;
                                    Area = "直行";
                                    ERR_String = "地址未绑定";
                                }


                            }
                            catch (Exception ex)
                            {
                                Route = 2;
                                Area = "直行";
                                ERR_String = ex.Message;
                            }
                            #endregion


                        }
                        else  //wms_to_ccs没有记录
                        {
                            ERR_String = "未查到该条码任务";
                            Route = 2;
                            Area = "直行";
                        }
                        #endregion





                    }
                    else //条码为空，当异常处理
                    {
                        ERR_String = "条码未读取";
                        Route = 2;
                        Area = "直行";
                    }
                }
                catch (Exception ex)
                {
                    ERR_String = ex.Message;
                    Route = 2;
                    Area = "直行";
                }
                #endregion


                Deal_with_Log(_barCode, "BH" + scan_no, Route.ToString(), ERR_String, CustomConnect);//流水记录
                //Deal_with_CCS_to_WMS_Log(box_num, Lpn, 2, CustomConnect);//往CCS_TO_WMS插流水记录，并复位wms_to_ccs地址


                if (!string.IsNullOrEmpty(ERR_String))
                {
                    if (!(ERR_String.Equals("条码未读取") || ERR_String.Equals("未查到该条码任务") || ERR_String.Equals("地址未绑定") || ERR_String.Equals("地址绑定错误")))
                    {
                        ERR_String = "BH" + scan_no + "处系统异常";
                    }


                    Deal_with_ERRINFO(_barCode, "BH" + scan_no, ERR_String, CustomConnect);
                    PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "ERROR_BH" + scan_no, ERR_String, "txt");

                }


                Byte[] write_stackinfo = new byte[4];
                write_stackinfo[0] = Convert.ToByte(2); //清READCODE  置READDONE 0000 0100
                write_stackinfo[1] = Convert.ToByte(0);
                write_stackinfo[2] = Convert.ToByte(Route / 256);
                write_stackinfo[3] = Convert.ToByte(Route % 256);
                S71500_1.WriteBytes(DataType.DataBlock, read_code_DB, read_code_Add, write_stackinfo);


                DateTime dtEnd = DateTime.Now;
                double dSub = SubDateTimeByMill(dtStar_1, dtEnd) / 1000;
                string hint = "扫描器=BCR" + scan_no + "，条码=" + _barCode + "，Route=" + Route + ",目的地 = " + Area + ",时间：" + dtEnd + "，耗时:" + dSub.ToString() + "秒" + " " + "\n";
                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "do_check_BCR" + scan_no, hint, "txt");
                PaintColorToRtxtbx(richTextBox1, hint, ref beginPosition, ref Print_Count1);


            }
            catch (Exception ex)
            {

                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "ERROR_BH" + scan_no, _barCode + ",发生异常:" + ex.ToString(), "txt");

                return;
            }

        }



        private void do_check_Area_Huo_down_In(int scan_no, int AREA, int read_code_DB, int read_code_Add)
        {

            string _barCode = "";
            string Area = "Area" + AREA;
            string  Area_TMP = "Area1,Area2,Area3,Area4,,Area5,,Area6,,Area7,,Area8,,Area9,,Area10,,Area10,,Area11";
            try
            {
                DateTime dtStar_1 = DateTime.Now;


                //先取到条码,再更新到wcs_to_ccs表   

                //_barCode = read_barcode(1, scan_no, barcode_DB, 0);
                //_barCode = "1234567890";


                int Route = 0;
                string ERR_String = "";
                string Route_tmp = "";
                string Empty_flag = "";
                string Lpn = "";


                string sql_select = "select * from OPENQUERY(REFINE, 'select * from WMS_TO_CCS where  Lpn=''" + _barCode + "'' order by  UpDate_Time desc limit 0,1')";


                #region 获取路向，判断异常


                try
                {
                    if (!(string.IsNullOrEmpty(_barCode) || _barCode.Equals("NoRead")))
                    {


                        DataSet ds = SQLServer.Query(sql_select, CustomConnect);

                        #region  判断是否WMS有记录，无则写异常信息
                        if (ds.Tables[0].Rows.Count > 0 && ds.Tables.Count > 0) //wms_to_ccs有记录
                        {
                            #region 判断路向
                            try
                            {
                                Route_tmp = ds.Tables[0].Rows[0][Area].ToString();
                                Lpn = ds.Tables[0].Rows[0]["Lpn"].ToString();
                                if (ds.Tables[0].Rows[0]["Area1"].ToString().Equals("0") && ds.Tables[0].Rows[0]["Area2"].ToString().Equals("0") 
                                    && ds.Tables[0].Rows[0]["Area4"].ToString().Equals("0") && ds.Tables[0].Rows[0]["Area5"].ToString().Equals("0") && ds.Tables[0].Rows[0]["Area6"].ToString().Equals("0")
                                    && ds.Tables[0].Rows[0]["Area7"].ToString().Equals("0") && ds.Tables[0].Rows[0]["Area8"].ToString().Equals("0") && ds.Tables[0].Rows[0]["Area9"].ToString().Equals("0")
                                    && ds.Tables[0].Rows[0]["Area10"].ToString().Equals("0") && ds.Tables[0].Rows[0]["Area11"].ToString().Equals("0"))

                                    Empty_flag = "0";
                                else
                                    Empty_flag = "1";
                                //如果有任务，直接进；没任务的话，如果空箱，则1、2货到人，谁空进谁；没任务又不是空箱，那么不进
                                if (!string.IsNullOrEmpty(Route_tmp))
                                {
                                    if (Route_tmp.Equals("0"))
                                    {
                                        if (Empty_flag.Equals("0"))
                                        { 
                                            Route = 1;
                                            Area = "货到人区域";
                                        }
                                        else
                                        {
                                            Route = 2;
                                            Area = "直行";
                                        }
                                    }
                                    else if (Route_tmp.Equals("1"))

                                    {
                                        Route = 1;
                                        Area = "货到人区域";

                                    }
                                    else
                                    {
                                        Route = 2;
                                        Area = "直行";
                                        ERR_String = "地址绑定错误";
                                    }

                                }
                                else
                                {
                                    Route = 2;
                                    Area = "直行";
                                    ERR_String = "地址未绑定";
                                }


                            }
                            catch (Exception ex)
                            {
                                Route = 2;
                                Area = "直行";
                                ERR_String = ex.Message;
                            }
                            #endregion


                        }
                        else  //wms_to_ccs没有记录
                        {
                            ERR_String = "未查到该条码任务";
                            Route = 2;
                            Area = "直行";
                        }
                        #endregion





                    }
                    else //条码为空，当异常处理
                    {
                        ERR_String = "条码未读取";
                        Route = 2;
                        Area = "直行";
                    }
                }
                catch (Exception ex)
                {
                    ERR_String = ex.Message;
                    Route = 2;
                    Area = "直行";
                }
                #endregion


                Deal_with_Log(_barCode, "BH" + scan_no, Route.ToString(), ERR_String, CustomConnect);//流水记录
                //Deal_with_CCS_to_WMS_Log(box_num, Lpn, 2, CustomConnect);//往CCS_TO_WMS插流水记录，并复位wms_to_ccs地址


                if (!string.IsNullOrEmpty(ERR_String))
                {
                    if (!(ERR_String.Equals("条码未读取") || ERR_String.Equals("未查到该条码任务") || ERR_String.Equals("地址未绑定") || ERR_String.Equals("地址绑定错误")))
                    {
                        ERR_String = "BH" + scan_no + "处系统异常";
                    }


                    Deal_with_ERRINFO(_barCode, "BH" + scan_no, ERR_String, CustomConnect);
                    PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "ERROR_BH" + scan_no, ERR_String, "txt");

                }


                Byte[] write_stackinfo = new byte[4];
                write_stackinfo[0] = Convert.ToByte(2); //清READCODE  置READDONE 0000 0100
                write_stackinfo[1] = Convert.ToByte(0);
                write_stackinfo[2] = Convert.ToByte(Route / 256);
                write_stackinfo[3] = Convert.ToByte(Route % 256);
                S71500_1.WriteBytes(DataType.DataBlock, read_code_DB, read_code_Add, write_stackinfo);


                DateTime dtEnd = DateTime.Now;
                double dSub = SubDateTimeByMill(dtStar_1, dtEnd) / 1000;
                string hint = "扫描器=BCR" + scan_no + "，条码=" + _barCode + "，Route=" + Route + ",目的地 = " + Area + ",时间：" + dtEnd + "，耗时:" + dSub.ToString() + "秒" + " " + "\n";
                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "do_check_BCR" + scan_no, hint, "txt");
                PaintColorToRtxtbx(richTextBox1, hint, ref beginPosition, ref Print_Count1);


            }
            catch (Exception ex)
            {

                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "ERROR_BH" + scan_no, _barCode + ",发生异常:" + ex.ToString(), "txt");

                return;
            }

        }


        private void do_check_Area_Huo_down_Out(int scan_no, int AREA, int read_code_DB, int read_code_Add)
        {

            string _barCode = "";
            string Area = "Area" + AREA;
            string Area_TMP = "Area1,Area2,Area3,Area4,,Area5,,Area6,,Area7,,Area8,,Area9,,Area10,,Area10,,Area11";
            try
            {
                DateTime dtStar_1 = DateTime.Now;


                //先取到条码,再更新到wcs_to_ccs表   

                //_barCode = read_barcode(1, scan_no, barcode_DB, 0);
                //_barCode = "1234567890";


                int Route = 0;
                string ERR_String = "";
                string Route_tmp = "";
                string Empty_flag = "";
                string Lpn = "";


                string sql_select = "select * from OPENQUERY(REFINE, 'select * from WMS_TO_CCS where  Lpn=''" + _barCode + "'' order by  UpDate_Time desc limit 0,1')";


                #region 获取路向，判断异常


                try
                {
                    if (!(string.IsNullOrEmpty(_barCode) || _barCode.Equals("NoRead")))
                    {


                        DataSet ds = SQLServer.Query(sql_select, CustomConnect);

                        #region  判断是否WMS有记录，无则写异常信息
                        if (ds.Tables[0].Rows.Count > 0 && ds.Tables.Count > 0) //wms_to_ccs有记录
                        {
                            #region 判断路向
                            try
                            {
                                Route_tmp = ds.Tables[0].Rows[0][Area].ToString();
                                Lpn = ds.Tables[0].Rows[0]["Lpn"].ToString();
                                if (ds.Tables[0].Rows[0]["Area1"].ToString().Equals("0") && ds.Tables[0].Rows[0]["Area2"].ToString().Equals("0")
                                    && ds.Tables[0].Rows[0]["Area4"].ToString().Equals("0") && ds.Tables[0].Rows[0]["Area5"].ToString().Equals("0") && ds.Tables[0].Rows[0]["Area6"].ToString().Equals("0")
                                    && ds.Tables[0].Rows[0]["Area7"].ToString().Equals("0") && ds.Tables[0].Rows[0]["Area8"].ToString().Equals("0") && ds.Tables[0].Rows[0]["Area9"].ToString().Equals("0")
                                    && ds.Tables[0].Rows[0]["Area10"].ToString().Equals("0") && ds.Tables[0].Rows[0]["Area11"].ToString().Equals("0"))

                                    Empty_flag = "0";
                                else
                                    Empty_flag = "1";
                                //如果有任务，直接进；没任务的话，如果空箱，则1、2货到人，谁空进谁；没任务又不是空箱，那么不进
                                if (!string.IsNullOrEmpty(Route_tmp))
                                {
                                    if (Route_tmp.Equals("0"))
                                    {
                                        if (Empty_flag.Equals("0"))
                                        {
                                            Route = 2;
                                            Area = "货到人区域";
                                        }
                                        else
                                        {
                                            Route = 3;
                                            Area = "直行";
                                        }
                                    }
                                    else if (Route_tmp.Equals("1"))

                                    {
                                        Route = 1;
                                        Area = "货到人区域";

                                    }
                                    else
                                    {
                                        Route = 3;
                                        Area = "直行";
                                        ERR_String = "地址绑定错误";
                                    }

                                }
                                else
                                {
                                    Route = 3;
                                    Area = "直行";
                                    ERR_String = "地址未绑定";
                                }


                            }
                            catch (Exception ex)
                            {
                                Route = 3;
                                Area = "直行";
                                ERR_String = ex.Message;
                            }
                            #endregion


                        }
                        else  //wms_to_ccs没有记录
                        {
                            ERR_String = "未查到该条码任务";
                            Route = 3;
                            Area = "直行";
                        }
                        #endregion





                    }
                    else //条码为空，当异常处理
                    {
                        ERR_String = "条码未读取";
                        Route = 3;
                        Area = "直行";
                    }
                }
                catch (Exception ex)
                {
                    ERR_String = ex.Message;
                    Route = 3;
                    Area = "直行";
                }
                #endregion


                Deal_with_Log(_barCode, "BH" + scan_no, Route.ToString(), ERR_String, CustomConnect);//流水记录
                //Deal_with_CCS_to_WMS_Log(box_num, Lpn, 2, CustomConnect);//往CCS_TO_WMS插流水记录，并复位wms_to_ccs地址


                if (!string.IsNullOrEmpty(ERR_String))
                {
                    if (!(ERR_String.Equals("条码未读取") || ERR_String.Equals("未查到该条码任务") || ERR_String.Equals("地址未绑定") || ERR_String.Equals("地址绑定错误")))
                    {
                        ERR_String = "BH" + scan_no + "处系统异常";
                    }


                    Deal_with_ERRINFO(_barCode, "BH" + scan_no, ERR_String, CustomConnect);
                    PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "ERROR_BH" + scan_no, ERR_String, "txt");

                }


                Byte[] write_stackinfo = new byte[4];
                write_stackinfo[0] = Convert.ToByte(2); //清READCODE  置READDONE 0000 0100
                write_stackinfo[1] = Convert.ToByte(0);
                write_stackinfo[2] = Convert.ToByte(Route / 256);
                write_stackinfo[3] = Convert.ToByte(Route % 256);
                S71500_1.WriteBytes(DataType.DataBlock, read_code_DB, read_code_Add, write_stackinfo);


                DateTime dtEnd = DateTime.Now;
                double dSub = SubDateTimeByMill(dtStar_1, dtEnd) / 1000;
                string hint = "扫描器=BCR" + scan_no + "，条码=" + _barCode + "，Route=" + Route + ",目的地 = " + Area + ",时间：" + dtEnd + "，耗时:" + dSub.ToString() + "秒" + " " + "\n";
                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "do_check_BCR" + scan_no, hint, "txt");
                PaintColorToRtxtbx(richTextBox1, hint, ref beginPosition, ref Print_Count1);


            }
            catch (Exception ex)
            {

                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "ERROR_BH" + scan_no, _barCode + ",发生异常:" + ex.ToString(), "txt");

                return;
            }

        }

        private void In_HuoDaoRen(int scan_no,  int G2P_NO, int AREA, string Station, int Barcode_Add,int read_code_DB, int read_code_Add)
        {

            string   _barCode = "";
            
            try
            {
                DateTime dtStar_1 = DateTime.Now;


                //先取到条码,再更新到wcs_to_ccs表   

                //_barCode = read_barcode(2, scan_no, barcode_DB, Barcode_Add);              
                string ERR_String = "";
 
                string sql_select = "update openquery(REFINE, 'select  * from CCS_TO_WMS_G2P where Station=''"+Station+"'' AND  G2P_NO="+ G2P_NO + "') "+
                "set Status = 1, tkdate = GETDATE(),Lpn='"+ _barCode+"'";

                int a = SQLServer.ExecuteSql(sql_select,CustomConnect);
                if (a < 1)
                {
                    ERR_String = "更新到位状态失败";

                }

                Deal_with_HDR_Log(_barCode, G2P_NO.ToString() , Station,"入", ERR_String, CustomConnect);//流水记录
                Deal_with_CCS_to_WMS_Log("Null", _barCode, AREA, CustomConnect);//往CCS_TO_WMS插流水记录，并复位wms_to_ccs地址


                if (!string.IsNullOrEmpty(ERR_String))
                {
                    Deal_with_ERRINFO(_barCode, "HDR_" + G2P_NO + "_" + Station, ERR_String, CustomConnect);
                    PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "ERROR_HDR_" + G2P_NO + "_" + Station, ERR_String, "txt");

                }

                Byte[] write_stackinfo = new byte[1];
                write_stackinfo[0] = Convert.ToByte(2); //清READCODE  置READDONE 0000 0100
               
                S71500_1.WriteBytes(DataType.DataBlock, read_code_DB, read_code_Add, write_stackinfo);


                DateTime dtEnd = DateTime.Now;
                double dSub = SubDateTimeByMill(dtStar_1, dtEnd) / 1000;
                string hint = "货到人组=" + G2P_NO + "，工位=" + Station + ",条码="+_barCode+"到位，时间：" + dtEnd + "，耗时:" + dSub.ToString() + "秒" + " " + "\n";
                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "do_check_HDR_" + G2P_NO + "_" + Station, hint, "txt");
                //PaintColorToRtxtbx(richTextBox2, hint, ref beginPosition_2, ref Print_Count2);


            }
            catch (Exception ex)
            {

                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "ERROR_HDR_" + G2P_NO + "_" + Station, _barCode + ",发生异常:" + ex.ToString(), "txt");

                return;
            }

        }

        private void Out_HuoDaoRen( int G2P_NO, string Station,int read_code_DB, int read_code_Add)
        {

            string _barCode = "";

            try
            {
                DateTime dtStar_1 = DateTime.Now;


                //先取到条码,再更新到wcs_to_ccs表   

              
                string ERR_String = "";
                int Route = 0;
                string sql_select = "SELECT* FROM  openquery(REFINE, 'select G2P_NO,Station,Lpn, UpDate_Time from ccs_to_wms_g2p"+
                " WHERE Status =2  and Station=''"+ Station+ "''  and    G2P_NO="+ G2P_NO+" order by  UpDate_Time desc limit 0, 1')";
               
                DataSet ds = new DataSet();
                ds = SQLServer.Query(sql_select,CustomConnect);
                if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {

                    Route = 2;
                    _barCode = ds.Tables[0].Rows[0]["Lpn"].ToString();
               

               
                Deal_with_HDR_Log(_barCode, G2P_NO.ToString(), Station, "出",ERR_String, CustomConnect);//流水记录
                    string NULL_BARCODE = "";
                    string sql_update = "update openquery(REFINE, 'select  * from CCS_TO_WMS_G2P where Station=''" + Station + "'' AND  G2P_NO=" + G2P_NO + "') " +
                                   "set Status = 0, tkdate = GETDATE(),Lpn='" + NULL_BARCODE + "'";

                    int a = SQLServer.ExecuteSql(sql_update, CustomConnect);
                    if (a < 1)
                    {
                        ERR_String = "更新到位状态失败";

                    }
                    //if (!string.IsNullOrEmpty(ERR_String))
                    //{
                    //    Deal_with_ERRINFO(_barCode, "HDR_" + G2P_NO + "_" + Station, ERR_String, CustomConnect);
                    //    PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "ERROR_HDR_" + G2P_NO + "_" + Station, ERR_String, "txt");

                    //}

                    Byte[] write_stackinfo = new byte[4];
                write_stackinfo[0] = Convert.ToByte(0); //清READCODE  置READDONE 0000 0100
                write_stackinfo[1] = Convert.ToByte(0);
                write_stackinfo[2] = Convert.ToByte(Route / 256);
                write_stackinfo[3] = Convert.ToByte(Route % 256);
                S71500_1.WriteBytes(DataType.DataBlock, read_code_DB, read_code_Add, write_stackinfo);

                DateTime dtEnd = DateTime.Now;
                double dSub = SubDateTimeByMill(dtStar_1, dtEnd) / 1000;
                string hint = "货到人组=" + G2P_NO + "，工位=" + Station + ",条码=" + _barCode + "拣货完成，时间：" + dtEnd + "，耗时:" + dSub.ToString() + "秒" + " " + "\n";
                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "Out_HDR_" + G2P_NO + "_" + Station, hint, "txt");
                //PaintColorToRtxtbx(richTextBox2, hint, ref beginPosition_2, ref Print_Count2);

                }
                else
                {
                    Route = 1;
                    
                }
            }
            catch (Exception ex)
            {

                PublicFunc.WritelnTxt(MainClass.Path, MainClass.sFile + "ERROR_Out_HDR_" + G2P_NO + "_" + Station, _barCode + ",发生异常:" + ex.ToString(), "txt");

                return;
            }

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



            #region MyRegion


            if (S71500_1.IsConnected )
            {
                #region 读条码标志


                //read_readcode(1, 0, Readcode_DB, 0);
                //read_readcode(4, 1, Readcode_DB, Readcode_DC1);//DC入口读条码标志
                //read_readcode(4, 2, Readcode_DB, Readcode_DC2);//DC入口读条码标志


                //read_readcode(2, 11, Readcode_Huo_DB, Readcode_Huo_11);
                //read_readcode(2, 12, Readcode_Huo_DB, Readcode_Huo_12);
                //read_readcode(2, 13, Readcode_Huo_DB, Readcode_Huo_13);
                //read_readcode(2, 14, Readcode_Huo_DB, Readcode_Huo_14);
                //read_readcode(2, 15, Readcode_Huo_DB, Readcode_Huo_15);
                //read_readcode(2, 16, Readcode_Huo_DB, Readcode_Huo_16);

                //read_readcode(2, 21, Readcode_Huo_DB, Readcode_Huo_21);
                //read_readcode(2, 22, Readcode_Huo_DB, Readcode_Huo_22);
                //read_readcode(2, 23, Readcode_Huo_DB, Readcode_Huo_23);
                //read_readcode(2, 24, Readcode_Huo_DB, Readcode_Huo_24);
                //read_readcode(2, 25, Readcode_Huo_DB, Readcode_Huo_25);
                //read_readcode(2, 26, Readcode_Huo_DB, Readcode_Huo_26);

               
                #endregion

                // read_Barcode_Count();

                #region  扫码器区域


                //if (percode[1].readcode)
                //{
                //    do_check_BCR1(1, Readcode_DB, Readcode_BCR1, Weight);
                //}
                //if (percode[3].readcode)
                //{
                //    do_check_Area2(3, Readcode_DB, Readcode_BCR3);
                //}
              
              
              


                #endregion


              


            }
            else
            {
                connect_plc1();
            }

            
            #endregion

                     
            
            //if (can_use==1)
            //{
            //    timer1.Enabled = true;
            //}

        }

       

        private void FrmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
           
     

            if (MessageBox.Show("是否确定关闭程序？", "关闭程序", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {


                S71500_1.Close();
                
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
               
              
                S71500_1.Close();
             
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

        private void button10_Click(object sender, EventArgs e)
        {
            sql_ERR = "";
            label5.Visible = true;
            button13.Visible = false;
            sql_ERR = "select Num as '地点' ,bar_code as '条码号',Insert_date as '发生时间',IF_Check as '是否处理',Check_date as '处理时间',ERR_INFO as '错误信息' from View_ERR_INFO_All where 1=1 ";

            if (checkBox4.Checked)
            {
                if (string.IsNullOrEmpty(textBox1.Text.Trim()))
                {
                    MessageBox.Show("请输入查询区域！！", "提示");
                    textBox1.Focus();
                    return;
                }
                else
                {
                    sql_ERR = sql_ERR + " and Num like '%" + textBox1.Text.Trim() + "%'";
                }
            }

            if (checkBox7.Checked)
            {
                if (string.IsNullOrEmpty(comboBox2.Text.Trim()))
                {
                    MessageBox.Show("请选择该异常是否已处理！！", "提示");
                    comboBox2.Focus();
                    return;
                }
                else
                {
                    if (comboBox2.Text.Trim().Equals("已处理"))
                    {
                        sql_ERR = sql_ERR + " and IF_Check ='1'";
                    }
                    else
                    {
                        sql_ERR = sql_ERR + " and (IF_Check ='0' or IF_Check  is null)";
                    }
                    
                }
            }

            //if (checkBox3.Checked)
            //{
                sql_ERR = sql_ERR + " and Insert_date>='" + dateTimePicker4.Value.ToString() + "' and Insert_date<='" + dateTimePicker3.Value.ToString() + "'";
            //}
            sql_ERR = sql_ERR + "  order by Insert_date  desc ";
            label5.Text = "";
            DataSet ds1 = SQLServer.Query(sql_ERR, MainClass.customConnect);
            label5.Text = "记录数:"+ds1.Tables[0].Rows.Count.ToString();
            dataGridView2.DataSource = ds1.Tables["ds"];
            // 设定包括Header和所有单元格的列宽自动调整
            //dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            // 设定包括Header和所有单元格的行高自动调整
            //dataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            //dataGridView2.Columns[0].Width = 70;
            //dataGridView2.Columns[1].Width = 150;
            //dataGridView2.Columns[2].Width = 150;
            //dataGridView2.Columns[3].Width = 80;
            //dataGridView2.Columns[4].Width = 150;
            //dataGridView2.Columns[5].Width = 238;// dataGridView3.RowHeadersWidth
            //dataGridView2.RowHeadersWidth = 50;
            //添加行号
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                dataGridView2.Rows[i].HeaderCell.Value = (i + 1).ToString();
                if (i < dataGridView2.Rows.Count-1)
                {
                    string tmp1 = "";
                    try
                    {
                        tmp1 = dataGridView2.Rows[i].Cells[3].Value.ToString();
                    }
                    catch {
                        tmp1 = "";
                    }
                   
                    if ("0".Equals(tmp1) || string.IsNullOrEmpty(tmp1))
                    {
                        dataGridView2.Rows[i].Cells["是否处理"].Value = "未处理";
                    }
                    else if ("1".Equals(tmp1))
                    {
                        dataGridView2.Rows[i].Cells["是否处理"].Value = "已处理";
                    }
            }

        }
            button13.Visible = true;

     }

     

        private void button11_Click_1(object sender, EventArgs e)
        {
            label9.Text = "";
            label9.Visible = false;
            button12.Visible = false;
            sql_BCR = "";
            sql_BCR = "SELECT   [BCR_NAME] as '条码阅读器',[Read_Count] as '识读条码数'  ,[Not_Read_count] as '未识读条码数'  ,[PENCENT] as '识读率(%)' ,[insert_date]  as '读取时间'  FROM [BCR_COUNT] where 1=1";
            if (checkBox6.Checked)
            {
                string BCR = "";
                try
                {
                    BCR = comboBox1.Text.ToString();
                    if (string.IsNullOrEmpty(BCR))
                    {
                        label9.Text = "<-请选择需要查询的条码阅读器编号";
                        label9.Visible = true;
                        return;
                    }
                    sql_BCR = sql_BCR + " AND BCR_NAME='" + BCR + "'";
                }
                catch(Exception ex)
                {
                    label9.Text = ex.Message.ToString();
                    label9.Visible = true;
                    return;
                }
               
                
            }

            //if (checkBox5.Checked)
            //{
                sql_BCR = sql_BCR + " and insert_date>='" + dateTimePicker6.Value.ToString() + "' and insert_date<='" + dateTimePicker5.Value.ToString() + "'";
            //}

            sql_BCR = sql_BCR + "  order by BCR_NO ";
            DataSet ds3 = SQLServer.Query(sql_BCR, MainClass.customConnect);

            dataGridView3.DataSource = ds3.Tables["ds"];
           
            // 设定包括Header和所有单元格的列宽自动调整
         //  dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            // 设定包括Header和所有单元格的行高自动调整
            //dataGridView3.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            //dataGridView3.Columns[0].Width = 100;
            //dataGridView3.Columns[1].Width = 100;
            //dataGridView3.Columns[2].Width = 100;
            //dataGridView3.Columns[3].Width = 100;
            //dataGridView3.Columns[4].Width = 200;// dataGridView3.RowHeadersWidth
            //dataGridView3.RowHeadersWidth = 70;
            //添加行号
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
               
                dataGridView3.Rows[i].HeaderCell.Value = (i + 1).ToString();
            }
            button12.Visible = true;
           
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
                
               
                S71500_1.Close();
          
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
               
               
                S71500_1.Close();
          
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
               
              
                S71500_1.Close();
          
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
    }
}
