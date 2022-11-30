using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Core;
using Spire.Xls.Core.Spreadsheet.Shapes;

namespace 干掉Excel多余文本框
{
    public partial class frmMaim : Form
    {
        public frmMaim()
        {
            InitializeComponent();
        }

        private void frmMaim_Load(object sender, EventArgs e)
        {
            Output("程序载入完成，请拖入文件（支持多文件）");        
        }
        private void Output(string text)
        {
            txtOutput.Text += DateTime.Now.ToString() + " " + text + "\r\n";
        }

        private void frmMaim_DragDrop(object sender, DragEventArgs e)
        {
            int i = 0;
            while (((System.Array)e.Data.GetData(DataFormats.FileDrop)).Length > i)
            {
                string filename = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(i).ToString();
                Output("接收到目标文件：" + filename);
                Workbook workbook = new Workbook();
                try
                {
                    workbook.LoadFromFile(filename);
                }
                catch
                {
                    Output("您看看"+filename+"这是Excel文件嘛");
                    return;
                }

                Worksheet sheet = workbook.Worksheets[0];
                int j = 0;
                while (sheet.TextBoxes.Count>j)
                {
                    sheet.TextBoxes[j].Remove();
                    j++;
                }
                sheet.Dispose();
                GC.Collect();
                Output("已经干掉" + j.ToString() + "个文本框");
                string newfilename = System.IO.Path.GetDirectoryName(filename) + @"\" + System.IO.Path.GetFileNameWithoutExtension(filename) + "-去文本框.xlsx";
                workbook.SaveToFile(newfilename, ExcelVersion.Version2007);
                Output("已经保存为新文件：" + newfilename);
                workbook.Dispose();
                GC.Collect();
                i++;
            }
            Output("任务完毕，共处理"+i.ToString()+"个文件");
        }

        private void frmMaim_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.All;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }
    }
}
