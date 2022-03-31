using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Sharetek
{
    static class Program
    {
        private static System.Threading.Mutex _mutex;
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            _mutex = new System.Threading.Mutex(false, "Sharetek");
            if (_mutex.WaitOne(0, false) == false)
            {
                MessageBox.Show("程序已经启动，不允许重复打开！");
                return;
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
