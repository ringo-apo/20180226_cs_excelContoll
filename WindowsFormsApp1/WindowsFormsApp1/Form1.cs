using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {

        Microsoft.Office.Interop.Excel.Application xlApp_cont = new Microsoft.Office.Interop.Excel.Application();
        Workbooks xlBooks_cont;
        Workbook xlBook_cont = null;
        Sheets xlSheets_cont = null;
        Worksheet xlSheet_cont = null;

        public Form1()
        {
            InitializeComponent();

            xlBooks_cont = xlApp_cont.Workbooks;

        }

        private void button1_Click(object sender, EventArgs e)
        {


            xlBook_cont = xlBooks_cont.Open(Path.GetFullPath(@"D:\data\develop\making\20180226_cs_excelContoll\ExcellController.xlsm"));
            xlSheets_cont = xlBook_cont.Worksheets;
            xlSheet_cont = xlSheets_cont[1] as Worksheet; // 1シート目を操作対象に設定する
            xlSheet_cont.Activate();
            xlApp_cont.Visible = true; // 表示

            xlApp_cont.Run("ExcellController.xlsm!FileOpen"); //目的のファイルのオープン

            xlApp_cont.Run("ExcellController.xlsm!Syori1", textBox1.Text); //目的のファイルのオープン

            xlApp_cont.Run("ExcellController.xlsm!FileClose"); //目的のファイルのクローズ

            //コントロールエクセルファイルのクローズ
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet_cont);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheets_cont);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook_cont);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBooks_cont);
            xlApp_cont.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp_cont);

        }
    }
}
