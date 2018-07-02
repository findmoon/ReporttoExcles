using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Exicel转换1
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            if (DateTime.Compare(DateTime.Now, Convert.ToDateTime("2018-9-30")) > 0)
            {
                return;
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new XmlExcelRtfConvertMain_Form());
        }
    }
}
