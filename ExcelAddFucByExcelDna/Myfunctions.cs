using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddFucByExcelDna
{
    public class Myfunctions
    {
        [ExcelFunction(Description = "Custom function that shows a message box.")]
        public static string SayHello()
        {
            // 弹出消息框
            System.Windows.Forms.MessageBox.Show("Hello!");
           
            return "=SayHello()";
        }
    }
}
