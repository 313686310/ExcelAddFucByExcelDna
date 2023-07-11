using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using ExcelDna.Integration;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelAddFucByExcelDna
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //xcel.Application excelApp = this.Application;
            //excelApp.WorkbookActivate += ExcelApp_WorkbookActivate;

            Excel.Application excelApp = this.Application;
            excelApp.SheetBeforeDoubleClick += ExcelApp_SheetBeforeDoubleClick;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void ExcelApp_WorkbookActivate(Excel.Workbook wb)
        {
            Excel.Worksheet worksheet = wb.ActiveSheet;
            string functionName = "MyCustomFunction";
            string functionDescription = "Test.";
            Excel.Range range = worksheet.get_Range("A1");
            range.Formula = "=MyCustomFunction()";
            range.Value = "你好"; // 计算公式
            var a = MessageBox.Show("你好");
            Excel.Name customFunction = wb.Names.Add(functionName, a);
            customFunction.Comment = functionDescription;
            //
        }
        private void ExcelApp_SheetBeforeDoubleClick(object Sh, Excel.Range Target, ref bool Cancel)
        {
            // 检查目标单元格是否具有注释
            if (Target.Comment != null)
            {
                string commentText = Target.Comment.Text().Trim();

                if (commentText== nameof(Myfunctions.SayHello))
                {
                    MessageBox.Show("你好");
                }
            }
        }
        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
