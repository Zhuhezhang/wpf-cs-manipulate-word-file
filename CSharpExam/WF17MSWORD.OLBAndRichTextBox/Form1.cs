//C#Windows窗体应用：1、利用MSWORD.OLB组件保存richTextBox里面的图片、文字为word文档
//                   2、加载/关闭word文档
//                   3、删除word文档
//                   4、实现编辑、清空、撤销、恢复操作
//                   5、可以往richtextbox控件拖动添加图片

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

namespace WF17MSWORD.OLBAndRichTextBox
{
    public partial class Form1 : Form
    {
        private string editingFilename;                   //正在编辑的文件的文件名
        public Form1()
        {
            InitializeComponent();
            editingFilename = "";
            richTextBox1.EnableAutoDragDrop = true;       //richTextBox的内容允许拖动
            timer1.Start();                               //timer启动
        }
        ~Form1()
        {
            timer1.Stop();                                //timer停止
        }

        //保存文件
        private void SaveButton_Click(object sender, EventArgs e)
        {
            if (editingFilename != "")                    //如果已经有正在编辑的文件，则不必调用保存文件对话框
                SaveFile(editingFilename);
            else                                          //如果是新建的文件，则调用保存文件对话框
            {
                if ((this.richTextBox1.Text == "") || (this.saveFileDialog1.ShowDialog() == DialogResult.Cancel))
                    return;                               //如果richTextBox1的文本为空或者在对话框中不点击确定，什么都不返回
                editingFilename = this.saveFileDialog1.FileName;
                if (editingFilename == null)              //如果文件名为空，什么都不返回
                    return;
                SaveFile(editingFilename);
            }
        }

        //保存文件函数
        private void SaveFile(string name)
        {
            try
            {
                Microsoft.Office.Interop.Word.Application myApp = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document myDoc = myApp.Documents.Add();
                                                          //添加到打开文档集合中的新空文档
                myDoc.ActiveWindow.Selection.WholeStory();//全选
                richTextBox1.SelectAll();                 //复制Rtf数据到剪贴板
                Clipboard.SetData(DataFormats.Rtf, richTextBox1.SelectedRtf); 
                myDoc.ActiveWindow.Selection.Paste();     //粘贴
                object myFileName = name;                 //string转化为object类型
                myDoc.SaveAs(ref myFileName);
                myDoc.Close();                            //关闭WordDoc文档对象
                myApp.Quit();                             //关闭WordApp组件对象
                MessageBox.Show("WORD文件保存成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }                                             //信息提示内容 对话框标题 显示确定按钮 显示感叹号图标
            catch (Exception ex)
            {
                MessageBox.Show("WORD文件保存失败！\n" + ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //打开word文档并加载到richTextBox
        private void OpenButton_Click(object sender, EventArgs e)
        {
            if (this.openFileDialog2.ShowDialog() == DialogResult.Cancel)
                return;                                   //获取正在打开的文件名
            editingFilename = this.openFileDialog2.FileName;
            Microsoft.Office.Interop.Word.Application myApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document myDoc = null;
            object missing = System.Reflection.Missing.Value;
            object fileName = this.openFileDialog2.FileName;
            try
            {
                myDoc = myApp.Documents.Open(ref fileName, ref missing, ref missing,
                                             ref missing, ref missing, ref missing, ref missing, ref missing,
                                             ref missing, ref missing, ref missing, ref missing, ref missing,
                                             ref missing, ref missing, ref missing);
                myDoc.ActiveWindow.Selection.WholeStory();//全选word文档中的数据
                myDoc.ActiveWindow.Selection.Copy();      //复制数据到剪切板
                richTextBox1.Paste();                     //richTextBox粘贴数据；richTextBox1.Text = doc.Content.Text;//显示无格式数据
                myDoc.Close();
                myApp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("WORD文件打开失败！\n" + ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //添加图片到richTextBox
        private void AddPhoto_Click(object sender, EventArgs e)
        {                                                 //设置对话框的过滤条件
            openFileDialog1.Filter = "img文件（*.img）|*.img|jpg 文件（*.jpg）|*.jpg|ico 文件（*.ico）|*.ico";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {                                             //加载图片到richTextBox
                Clipboard.SetDataObject(Image.FromFile(openFileDialog1.FileName), false);
                richTextBox1.Paste();                     //图片放在剪贴板中；false：退出程序后不将图片保留在剪贴板中
            }
            else
                return;
        }

        //恢复编辑
        private void RedoButton_Click(object sender, EventArgs e)
        {
            richTextBox1.Redo();
        }

        //撤销编辑
        private void UndoButton_Click(object sender, EventArgs e)
        {
            richTextBox1.Undo();
        }

        //清空richTextBox
        private void ClearallButtton_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
        }

        //关闭文档
        private void CloseButton_Click(object sender, EventArgs e)
        {
            editingFilename = "";
            richTextBox1.Text = "";
        }

        //定时刷新lable1的内容，用于显示正在编辑/打开的文件名
        private void timer1_Tick(object sender, EventArgs e)
        {
            this.label1.Text = "正在编辑文件：" + Path.GetFileName(editingFilename);
        }

        //删除文件
        private void DeleteButton_Click(object sender, EventArgs e)
        {
            if (editingFilename == "")
                return;
            else
            {
                File.Delete(editingFilename);
                editingFilename = "";
                richTextBox1.Text = "";
            }
        }
    }
}