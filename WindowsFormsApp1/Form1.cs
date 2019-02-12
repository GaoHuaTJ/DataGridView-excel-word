using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using Microsoft.Office;
using Microsoft.Office.Interop.Word;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private List<Fruit>P_fruit;
        private void Form1_Load(object sender, EventArgs e)
        {
            //添加数据，绑定数据源
             P_fruit= new List<Fruit>() {
                new Fruit { Name = "苹果", Price = 30},
                new Fruit { Name = "香蕉", Price = 40},
                new Fruit { Name = "西瓜", Price = 50},
                new Fruit { Name = "苹果", Price = 60},
            };
            dgv_Message.DataSource = P_fruit;


            //修改列名
            dgv_Message.Columns[0].HeaderText = "水果";
            dgv_Message.Columns[1].HeaderText = "价格";

            float sum = 0;
            //计算平均数和总数
            P_fruit.ForEach(
                (pp) => {
                     sum+= pp.Price;//计算几个的总和
                }
                );

            //隔行换色
            for (int i=0;i<dgv_Message.Rows.Count;i++)
            {
                if (i %2 == 0)
                {
                    dgv_Message.Rows[i].DefaultCellStyle.BackColor = Color.LightYellow;
                }
            }


        //新建状态列
            DataGridViewCheckBoxColumn dgvc = new DataGridViewCheckBoxColumn();
            dgvc.HeaderText = "选中并删除";
            dgv_Message.Columns.Add(dgvc);
            dgv_Message.AutoGenerateColumns = false;//防止datagridview自动变化顺序
        }
        /// <summary>
        /// 删除datagridview中的选定行数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_Remove_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dgv_Message.Rows.Count; i++)//遍历所有的行
            {
                if ((dgv_Message.Rows[i].Cells[0].Value != null) &&
                    (dgv_Message.Rows[i].Cells[1].Value != null) &&
                    (dgv_Message.Rows[i].Cells[2].Value != null))//判断所选择的行的数据是否为空
                {
                    Console.WriteLine(dgv_Message.Rows[i].Cells[0].Value.ToString());
                    if (dgv_Message.Rows[i].Cells[2].EditedFormattedValue.ToString() == "True")//判断第一列是否被选中
                    {
                        P_fruit.RemoveAll(
                        (pp) =>
                        {
                            if ((pp.Name == dgv_Message.Rows[i].Cells[0].Value.ToString())
                            && (pp.Price == int.Parse(dgv_Message.Rows[i].Cells[1].Value.ToString())))
                            pp.ft = true;
                            return false;
                        });
                    }
                }
            }


            var MessageBoxResult = MessageBox.Show("即将删除选定的数据", "警示", MessageBoxButtons.OKCancel);//删除数据再次确定
            if (MessageBoxResult == DialogResult.OK)
            {
                P_fruit.RemoveAll(
                (pp) =>
                {
                    return pp.ft; 
                });
            }

            dgv_Message.DataSource = null;//数据源置空
            dgv_Message.DataSource = P_fruit;//赋值新的数据源
        }


        /// <summary>
        /// 将文件到出为word文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private Microsoft.Office.Interop.Word.Application G_wa;
        private object G_missing = System.Reflection.Missing.Value;
        public string ProcessStr;

        /// <summary>
        /// datagridview中的数据输出至word中间
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog P_SaveFileDialog = new SaveFileDialog();
            P_SaveFileDialog.Filter = "*.doc|*.doc";//该字符串是另存为文件类型文本框中的筛选字符串
            if (P_SaveFileDialog.ShowDialog() == DialogResult.OK)//如果保存文件框的返回结果是ok
            {

                TextBoxProcess.Invoke((MethodInvoker)(() =>
                {
                    ProcessStr = "---正在输出word---\r\n";
                    TextBoxProcess.Text = ProcessStr;
                    TextBoxProcess.Visible = true;
                }));

                ThreadPool.QueueUserWorkItem(
                    (pp) =>
                    {
                        TextBoxProcess.Invoke((MethodInvoker)(() =>
                        {
                            ProcessStr += "正在连接Word\r\n";
                            TextBoxProcess.Text = ProcessStr;
                        }));
                        G_wa = new Microsoft.Office.Interop.Word.Application();//创建应用程序对象

                        TextBoxProcess.Invoke((MethodInvoker)(() =>
                        {
                            ProcessStr += "连接Word完成\r\n";
                            TextBoxProcess.Text = ProcessStr;
                        }));

                        var Path = Environment.CurrentDirectory;//当前的工作路径
                        object G_Templete = Path+"\\resource\\上海铁路局数据处理结果表（模板）.dotx";//获取当前的模板文件地址
                        Microsoft.Office.Interop.Word.Document P_wd = G_wa.Documents.Add(ref G_Templete,ref G_missing, ref G_missing, ref G_missing);//想word程序中添加文档

                        

                        TextBoxProcess.Invoke((MethodInvoker)(()=> 
                        {
                            ProcessStr += "Word模板读取完成\r\n";
                            TextBoxProcess.Text = ProcessStr;
                        }));
                        

                        Microsoft.Office.Interop.Word.Range P_Range = P_wd.Range(ref G_missing, ref G_missing);//得到文档的范围
                        object o1 = Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord8TableBehavior;//规定表格的类型
                        object o2 = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow;//布满窗口
                        Microsoft.Office.Interop.Word.Table P_table = P_Range.Tables.Add(P_Range, dgv_Message.Rows.Count+1, dgv_Message.Columns.Count, ref o1, ref o2);//创建表格

                        
                        TextBoxProcess.Invoke((MethodInvoker)(() =>
                        {
                            ProcessStr += "Word表格创建完成\r\n";
                            TextBoxProcess.Text = ProcessStr;
                        }
                        ));

                        P_table.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                        P_table.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;

                        
                        TextBoxProcess.Invoke((MethodInvoker)(() =>
                        {
                            ProcessStr += "Word表格线条设计完成\r\n";
                            TextBoxProcess.Text = ProcessStr;
                        }
                        ));
                        //输出列标签
                        for (int i = 0; i < dgv_Message.Columns.Count; i++)
                        {
                            P_table.Cell(1, i + 1).Range.InsertAfter(dgv_Message.Columns[i].HeaderText);
                        }
                        //输出数据
                        for (int i = 0; i <dgv_Message.Rows.Count; i++)
                        {
                            for (int j = 0; j < dgv_Message.Columns.Count; j++)
                            {
                                P_table.Cell(i+2, j+1).Range.InsertAfter(dgv_Message[j,i].FormattedValue.ToString());//注意此处的datagridview的索引的顺序，先是列，后是行
                            }
                        }

                        TextBoxProcess.Invoke((MethodInvoker)(() =>
                        {
                            ProcessStr += "Word表格数据输出完成\r\n";
                            TextBoxProcess.Text = ProcessStr;
                        }));


                        object P_path = P_SaveFileDialog.FileName;//文件保存对象
                        P_wd.SaveAs2(ref P_path);
                        

                        TextBoxProcess.Invoke((MethodInvoker)(() =>
                        {
                            ProcessStr += "Word文件保存完成\r\n";
                            TextBoxProcess.Text = ProcessStr;
                        }
                        ));

                        G_wa.Application.Quit();//退出应用程序

                        this.Invoke((MethodInvoker)(()=>
                        {
                            MessageBox.Show("文件保存完毕，地址为"+P_path,"提示");
                            TextBoxProcess.Visible = false;
                        }));
                    }
                    );
            }

        }

        /// <summary>
        /// 将datagridview中的数据导出到excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_ToExcel_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog_Excel = new SaveFileDialog();
            //saveFileDialog_Excel.Filter= "*.xls,*.xlsx|*.xls,*.xlsx";//默认的保存类型
            saveFileDialog_Excel.Filter = "*.xlsx|*.xlsx";//默认的保存类型

            if (saveFileDialog_Excel.ShowDialog() == DialogResult.OK)//如果确认要保存
            {//开启线程池
                ThreadPool.QueueUserWorkItem(
                    (pp)=>
                    {
                        Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();//创建excel的应用对象
                        Microsoft.Office.Interop.Excel.Workbook MyWorkbook = excel.Workbooks.Add();//创建excel文档
                        Microsoft.Office.Interop.Excel.Worksheet MyWorkField = (Microsoft.Office.Interop.Excel.Worksheet)MyWorkbook.Worksheets.Add();//创建工作区域

                        //输出列标签
                        for (int i = 0; i < dgv_Message.Columns.Count; i++)
                        {
                            MyWorkField.Cells[1, i+1] = dgv_Message.Columns[i].HeaderText;//第一行限定列名
                        }

                        //输出数据
                        for (int i = 0; i < dgv_Message.Rows.Count; i++)
                        {
                            for (int j = 0; j < dgv_Message.Columns.Count; j++)
                            {
                                MyWorkField.Cells[i + 2, j + 1] = dgv_Message[j, i].FormattedValue.ToString();//注意此处的datagridview的索引的顺序，先是列，后是行
                            }
                        }

                        object MyWorkBook_Path = saveFileDialog_Excel.FileName;//文件的存储地址
                        MyWorkbook.SaveAs(MyWorkBook_Path);//存储文档

                        this.Invoke((MethodInvoker)(()=>
                            {
                                MessageBox.Show("文件保存完毕，地址为" + MyWorkBook_Path, "提示");
                            }));

                        excel.Application.Quit();//关闭软件
                    } );
            }

        }
    }
}
