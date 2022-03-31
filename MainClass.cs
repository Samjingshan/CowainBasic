using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HslCommunication;
using System.Windows.Forms;
using MyFuncDll;
using System.Configuration;
namespace Sharetek
{
    class MainClass
    {
        public static string Path = "E://CCYQ_LOG";
        public static string sFile = string.Format("{0}{1:00}{2:00}", DateTime.Now.Date.Year, DateTime.Now.Date.Month, DateTime.Now.Date.Day);
     
        public static string customConnect = ConfigurationManager.ConnectionStrings["ConnectionString"].ToString();
        /// <summary>
        /// 统一的读取结果的数据解析，显示
        /// </summary>
        /// <typeparam name="T">类型对象</typeparam>
        /// <param name="result">读取的结果值</param>
        /// <param name="address">地址信息</param>
        /// <param name="textBox">输入的控件</param>
        public static void ReadResultRender<T>(OperateResult<T> result, ref string s)
        {
            if (result.IsSuccess)
            {

               s=result.Content.ToString();
            }
            else
            {
                //MessageBox.Show(DateTime.Now.ToString("[HH:mm:ss] ") + address.ToString() + "Read Failed {Environment.NewLine}Reason：{result.ToMessageShowString()}");
            }
        }

        /// <summary>
        /// 统一的数据写入的结果显示
        /// </summary>
        /// <param name="result">写入的结果信息</param>
        /// <param name="address">地址信息</param>
        public static void WriteResultRender(OperateResult result, string address)
        {
            if (result.IsSuccess)
            {
                MessageBox.Show(DateTime.Now.ToString("[HH:mm:ss] ") + address.ToString() + " Write Success");
            }
            else
            {
                MessageBox.Show(DateTime.Now.ToString("[HH:mm:ss] ") + address.ToString() + "Write Failed {Environment.NewLine} Reason：{result.ToMessageShowString()}");
            }
        }

        /// <summary>
        /// 统一的数据写入的结果显示
        /// </summary>
        /// <param name="result">写入的结果信息</param>
        /// <param name="address">地址信息</param>
        public static void WriteResultRender(Func<OperateResult> write, string address)
        {
            try
            {
                OperateResult result = write();
                if (result.IsSuccess)
                {
                    MessageBox.Show(DateTime.Now.ToString("[HH:mm:ss] ") + address.ToString() + " Write Success");
                }
                else
                {
                    MessageBox.Show(DateTime.Now.ToString("[HH:mm:ss] ") + address.ToString() + "Write Failed {Environment.NewLine} Reason：{result.ToMessageShowString()}");
                }
            }
            catch (Exception ex)
            {
                // 主要是为了捕获写入的值不正确的情况
                MessageBox.Show("Data for writting is not corrent: " + ex.Message);
            }
        }

        public static void BulkReadRenderResult(HslCommunication.Core.IReadWriteNet readWrite, TextBox addTextBox, TextBox lengthTextBox, TextBox resultTextBox)
        {
            try
            {
                OperateResult<byte[]> read = readWrite.Read(addTextBox.Text, ushort.Parse(lengthTextBox.Text));
                if (read.IsSuccess)
                {
                    resultTextBox.Text = "Result：" + HslCommunication.BasicFramework.SoftBasic.ByteToHexString(read.Content);
                }
                else
                {
                    MessageBox.Show("Read Failed：" + read.ToMessageShowString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Read Failed：" + ex.Message);
            }
        }

        public static readonly string IpAddressInputWrong = "IpAddress input wrong";
        public static readonly string PortInputWrong = "Port input wrong";
        public static readonly string SlotInputWrong = "Slot input wrong";
        public static readonly string BaudRateInputWrong = "Baud rate input wrong";
        public static readonly string DataBitsInputWrong = "Data bit input wrong";
        public static readonly string StopBitInputWrong = "Stop bit input wrong";
    }
}
