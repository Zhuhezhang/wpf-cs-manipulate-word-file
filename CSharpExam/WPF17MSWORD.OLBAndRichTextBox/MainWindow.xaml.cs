//C#WPF应用：1、利用MSWORD.OLB组件保存richTextBox里面的图片、文字为word文档
//                   2、加载/关闭word文档
//                   3、删除word文档
//                   4、实现编辑、清空、撤销、恢复操作
//                   5、可以往richtextbox控件拖动添加图片

using System;
using System.IO;
using System.Windows.Documents;
using System.Windows.Input;
using System.Drawing;
using System.Windows.Forms;
using  Word = Microsoft.Office.Interop.Word;
using MessageBox = System.Windows.Forms.MessageBox;
using Clipboard = System.Windows.Forms.Clipboard;
using System.Windows.Threading;
using Microsoft.Office.Interop.Word;

namespace WPF17MSWORD.OLBAndRichTextBox
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private string editingFilename;                   //正在编辑的文件的文件名
        DispatcherTimer timer1 = new DispatcherTimer();   //刷新editingFilename的值
        OpenFileDialog openPhotoFileDialog = new OpenFileDialog();
        OpenFileDialog openFileDialog = new OpenFileDialog();
        SaveFileDialog saveFileDialog = new SaveFileDialog();

        public MainWindow()
        {
            InitializeComponent();
            saveFileDialog.DefaultExt = "docx";           //默认扩展名
            new TextRange(this.richTextBox1.Document.ContentStart, this.richTextBox1.Document.ContentEnd).Text = "请从此处开始输入信息";
            editingFilename = "";
            timer1.Tick += new EventHandler(timer1_Tick);
            timer1.Start();                               //timer启动
        }
        ~MainWindow()
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
                if (((new TextRange(this.richTextBox1.Document.ContentStart, this.richTextBox1.Document.ContentEnd).Text) == "")
                                                 ||this.saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
                    return;                               //如果文件为空或者在对话框中不点击确定，什么都不返回
                editingFilename = this.saveFileDialog.FileName;
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
                Word.Application myApp = new Word.Application();
                Word.Document myDoc = myApp.Documents.Add();//添加到打开文档集合中的新空文档
                myDoc.ActiveWindow.Selection.WholeStory();  //全选
                richTextBox1.SelectAll();
                richTextBox1.Copy();                        //复制数据到剪贴板

                 
                myDoc.ActiveWindow.Selection.Paste();       //粘贴
                myDoc.ActiveWindow.Selection.WholeStory();
                myApp.Selection.Font.Color = WdColor.wdColorBlack;
                object myFileName = name;                   //string转化为object类型
                myDoc.SaveAs(ref myFileName);
                myDoc.Close();                              //关闭WordDoc文档对象
                myApp.Quit();                               //关闭WordApp组件对象
                MessageBox.Show("WORD文件保存成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }                                               //信息提示内容 对话框标题 显示确定按钮 显示感叹号图标
            catch (Exception ex)
            {
                MessageBox.Show("WORD文件保存失败！\n" + ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //打开word文档并加载到richTextBox
        private void OpenButton_Click(object sender, EventArgs e)
        {
            new TextRange(this.richTextBox1.Document.ContentStart, this.richTextBox1.Document.ContentEnd).Text = "";
            if (this.openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
                return;                                     //获取正在打开的文件名
            editingFilename = this.openFileDialog.FileName;
            Word.Application myApp = new Word.Application();
            Word.Document myDoc = null;
            object missing = System.Reflection.Missing.Value;
            object fileName = this.openFileDialog.FileName;
            try
            {
                myDoc = myApp.Documents.Open(ref fileName);
                myDoc.ActiveWindow.Selection.WholeStory();  //全选word文档中的数据
                myApp.Selection.Font.Color = WdColor.wdColorWhite;
                myDoc.ActiveWindow.Selection.Copy();        //复制数据到剪切板
                richTextBox1.Paste();                       //richTextBox粘贴数据；richTextBox1.Text = doc.Content.Text;//显示无格式数据
                myDoc.Close();
                myApp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("WORD文件打开失败！\n" + ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //添加图片到richTextBox
        private void AddPhotoButton_Click(object sender, EventArgs e)
        {                                                   //设置对话框的过滤条件
            openPhotoFileDialog.Filter = "png 文件（*.png）|*.png|jpg 文件（*.jpg）|*.jpg|bmp 文件（*.bmp）|*.bmp";
            if (openPhotoFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {                                               //加载图片到richTextBox
                Clipboard.SetDataObject(Image.FromFile(openPhotoFileDialog.FileName), false);
                richTextBox1.Paste();                       //图片放在剪贴板中；false：退出程序后不将图片保留在剪贴板中
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
        private void ClearAllButton_Click(object sender, EventArgs e)
        {
            richTextBox1.Document.Blocks.Clear();
        }

        //关闭文档
        private void CloseButton_Click(object sender, EventArgs e)
        {
            editingFilename = "";
            new TextRange(this.richTextBox1.Document.ContentStart, this.richTextBox1.Document.ContentEnd).Text = "";
        }

        //定时刷新lable1的内容，用于显示正在编辑/打开的文件名
        private void timer1_Tick(object sender, EventArgs e)
        {
            this.ShowFileNameLable.Content = "正在编辑文件：" + System.IO.Path.GetFileName(editingFilename);
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
                new TextRange(this.richTextBox1.Document.ContentStart, this.richTextBox1.Document.ContentEnd).Text = "";
            }
        }

        //清除richBoxText提示的输入信息
        private void PromptInputInfoOfRichTextBox(object sender, EventArgs e)
        {
            richTextBox1.Document.Blocks.Clear();           //取消订阅该事件，使得该方法只执一次
            richTextBox1.PreviewMouseDown -= PromptInputInfoOfRichTextBox;
        }

        //鼠标进入按钮控件，字体变黑；鼠标出来，字体恢复
        //撤消按钮
        private void UndoButton_MouseEnter(object sender, EventArgs e)
        {
            UndoButton.Foreground = System.Windows.Media.Brushes.Black;
        }
        private void UndoButton_MouseLeave(object sender, EventArgs e)
        {
            UndoButton.Foreground = System.Windows.Media.Brushes.White;
        }
        //恢复按钮
        private void RedoButton_MouseEnter(object sender, EventArgs e)
        {
            RedoButton.Foreground = System.Windows.Media.Brushes.Black;
        }
        private void RedoButton_MouseLeave(object sender, EventArgs e)
        {
            RedoButton.Foreground = System.Windows.Media.Brushes.White;
        }
        //清除按钮
        private void ClearAllButton_MouseEnter(object sender, EventArgs e)
        {
            ClearAllButton.Foreground = System.Windows.Media.Brushes.Black;
        }
        private void ClearAllButton_MouseLeave(object sender, EventArgs e)
        {
            ClearAllButton.Foreground = System.Windows.Media.Brushes.White;
        }
        //添加图片按钮
        private void AddPhotoButton_MouseEnter(object sender, EventArgs e)
        {
            AddPhotoButton.Foreground = System.Windows.Media.Brushes.Black;
        }
        private void AddPhotoButton_MouseLeave(object sender, EventArgs e)
        {
            AddPhotoButton.Foreground = System.Windows.Media.Brushes.White;
        }
        //删除文件按钮
        private void DeleteButton_MouseEnter(object sender, EventArgs e)
        {
            DeleteButton.Foreground = System.Windows.Media.Brushes.Black;
        }
        private void DeleteButton_MouseLeave(object sender, EventArgs e)
        {
            DeleteButton.Foreground = System.Windows.Media.Brushes.White;
        }
        //关闭文件按钮
        private void CloseButton_MouseEnter(object sender, EventArgs e)
        {
            CloseButton.Foreground = System.Windows.Media.Brushes.Black;
        }
        private void CloseButton_MouseLeave(object sender, EventArgs e)
        {
            CloseButton.Foreground = System.Windows.Media.Brushes.White;
        }
        //打开文件按钮
        private void OpenButton_MouseEnter(object sender, EventArgs e)
        {
            OpenButton.Foreground = System.Windows.Media.Brushes.Black;
        }
        private void OpenButton_MouseLeave(object sender, EventArgs e)
        {
            OpenButton.Foreground = System.Windows.Media.Brushes.White;
        }
        //保存文件按钮
        private void SaveButton_MouseEnter(object sender, EventArgs e)
        {
            SaveButton.Foreground = System.Windows.Media.Brushes.Black;
        }
        private void SaveButton_MouseLeave(object sender, EventArgs e)
        {
            SaveButton.Foreground = System.Windows.Media.Brushes.White;
        }

        //按钮快捷键
        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            //Ctrl+U 撤销
            if ((e.KeyboardDevice.IsKeyDown(Key.LeftCtrl) || e.KeyboardDevice.IsKeyDown(Key.RightCtrl)) && e.KeyboardDevice.IsKeyDown(Key.U))
            {
                UndoButton_Click(null, null);
            }
            //Ctrl+R 恢复
            if ((e.KeyboardDevice.IsKeyDown(Key.LeftCtrl) || e.KeyboardDevice.IsKeyDown(Key.RightCtrl)) && e.KeyboardDevice.IsKeyDown(Key.R))
            {
                RedoButton_Click(null, null);
            }
            //Ctrl+C 清空
            if ((e.KeyboardDevice.IsKeyDown(Key.LeftCtrl) || e.KeyboardDevice.IsKeyDown(Key.RightCtrl)) && e.KeyboardDevice.IsKeyDown(Key.C))
            {
                ClearAllButton_Click(null, null);
            }
            //Ctrl+A 添加图片
            if ((e.KeyboardDevice.IsKeyDown(Key.LeftCtrl) || e.KeyboardDevice.IsKeyDown(Key.RightCtrl)) && e.KeyboardDevice.IsKeyDown(Key.A))
            {
                AddPhotoButton_Click(null, null);
            }
            //Ctrl+D 删除
            if ((e.KeyboardDevice.IsKeyDown(Key.LeftCtrl) || e.KeyboardDevice.IsKeyDown(Key.RightCtrl)) && e.KeyboardDevice.IsKeyDown(Key.D))
            {
                DeleteButton_Click(null, null);
            }
            //Ctrl+C 关闭
            if ((e.KeyboardDevice.IsKeyDown(Key.LeftCtrl) || e.KeyboardDevice.IsKeyDown(Key.RightCtrl)) && e.KeyboardDevice.IsKeyDown(Key.G))
            {
                CloseButton_Click(null, null);
            }
            //Ctrl+O 打开
            if ((e.KeyboardDevice.IsKeyDown(Key.LeftCtrl) || e.KeyboardDevice.IsKeyDown(Key.RightCtrl)) && e.KeyboardDevice.IsKeyDown(Key.O))
            {
                OpenButton_Click(null, null);
            }
            //Ctrl+S 保存
            if ((e.KeyboardDevice.IsKeyDown(Key.LeftCtrl) || e.KeyboardDevice.IsKeyDown(Key.RightCtrl)) && e.KeyboardDevice.IsKeyDown(Key.S))
            {
                SaveButton_Click(null, null);
            }
        }
    }
}