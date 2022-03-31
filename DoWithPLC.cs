using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using PLC_Adapt;
using Sharetek_Common;
using System.IO;

namespace Sharetek
{
   public class DoWithPLC
    {
    
        Dictionary<String, PLCItem> PLC_Info_itemDict = new Dictionary<string, PLCItem>();
        public FTP_Fun Ftp_Connect;
        PLCManager plcManager = null;
        private string Log_path = "";
        string RemoteGetFilePath1 = "";
        string DownLoadpath1 = "";

        string RemoteGetFilePath2 = "";
        string DownLoadpath2 = "";

        public DoWithPLC(PLCManager PM,string Log_Path_)
        {
            this.Log_path = Log_Path_;
            this.plcManager = PM ;
            this.plcManager.PLCEvent += LoaderPlcManager_PLCEvent;

            string FTP_IP = ConfigurationManager.AppSettings["FTP_IP"];
            string USER = ConfigurationManager.AppSettings["USER"];
            string PASSWORD = ConfigurationManager.AppSettings["PASSWORD"];
             RemoteGetFilePath1 = ConfigurationManager.AppSettings["RemoteGetFilePath1"];
             DownLoadpath1= ConfigurationManager.AppSettings["DownLoadFilePath1"];

             RemoteGetFilePath2 = ConfigurationManager.AppSettings["RemoteGetFilePath2"];
             DownLoadpath2 = ConfigurationManager.AppSettings["DownLoadFilePath2"];

            PLC_Info_itemDict = this.plcManager.plc_InfoItemDict;
            Ftp_Connect = new FTP_Fun(DownLoadpath1,FTP_IP, RemoteGetFilePath1, USER, PASSWORD, Log_Path_);
          

            
       
        }



        

        private void LoaderPlcManager_PLCEvent(PLCEvent Event)
        {
            try
            {
                if (Event.item != null)
                {
                    switch (Event.item.name)
                    {
                        case "Read_Code_Flag_1":
                            {
                             
                               
                                if (Event.newValue.Equals("1"))//Read_Code_Flag由0置1
                                {
                                    string BarCode = this.plcManager.adapter.Read_Single("Bar_Code1", PLC_Info_itemDict, 20);//读条码
                                    string Planeness = this.plcManager.adapter.Read_Single("Planeness", PLC_Info_itemDict);//读平整度
                                    bool a =Ftp_Connect.Down_Load_File(); 
                                    this.plcManager.adapter.Write_Single("Read_Code_Finished1", "1", PLC_Info_itemDict);//写完成信号
                                    MyFuncDll.PublicFunc.WritelnTxt(this.Log_path,"条码："+BarCode+"获取图片成功！", "DownLoadPicSucess1");
                                }
                                break;
                            }
                        case "Read_Code_Flag_2":
                            {

                                if (Event.newValue.Equals ("1"))//Read_Code_Flag由0置1
                                {


                                    DirectoryInfo root = new DirectoryInfo(RemoteGetFilePath2);

                                  
                                    foreach (FileInfo f in root.GetFiles())
                                    {
                                        string name = f.Name;
                                        string fullName = f.FullName;
                                        string str = DownLoadpath2 + "//" + name;
                                        if (name.EndsWith(".jpg"))
                                            File.Copy(f.FullName, str, true);
                                        File.Delete(fullName);
                                    }

                                    string BarCode = this.plcManager.adapter.Read_Single("Bar_Code2", PLC_Info_itemDict, 20);//读条码
                                    

                                    this.plcManager.adapter.Write_Single("Read_Code_Finished2", "TRUE", PLC_Info_itemDict);//写完成信号
                                    MyFuncDll.PublicFunc.WritelnTxt(this.Log_path, "条码：" + BarCode + "获取图片成功！", "GetPicSucess2");
                                }
                                break;
                            }
                        default:break;
                    }
                }
            }
            catch (Exception ex)
            {
                //_Message = CoreHelper.GetExceptionInfo(ex).Message;
                //ShowMessage(_Message);
                //CoreHelper.Logger.Error(_Message);
            }
        }

        




        
      
        
      
        
    }

}
