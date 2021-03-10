@[TOC](目录)
# 1.题目
## 1.1基本要求
操作WORD文件：使用MSWORD.OLB组件将RichTextBox中的文本保存为WORD格式文件。
![在这里插入图片描述](https://img-blog.csdnimg.cn/20210311000014972.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQzNzk0NjMz,size_16,color_FFFFFF,t_70#pic_center)

## 1.2 本人新增额外功能
1、RichTextBox的内容实现撤消、恢复、清空、添加图片的操作；
2、关闭、打开WORD格式文件；
3、显示正在打开/编辑的文件名；
4、自定义背景图；
5、快捷键点击按钮；
6、控件背景颜色透明化；
7、鼠标移动到/离开按钮上方，字体颜色变化；
8、分别利用WPF应用和Windows窗体应用实现这些功能（由于两者的设计思想基本一致，所以这里只介绍利用WPF应用实现的，但两者的源码都会和报告一起提交）。 
# 2.窗体运行截图
 ![在这里插入图片描述](https://img-blog.csdnimg.cn/20210311000006841.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQzNzk0NjMz,size_16,color_FFFFFF,t_70#pic_center)

# 3.使用说明
利用Visual Stdio打开源码内的项目文件夹中的后缀为.lsn的文件即可开始使用。
打开窗体，RichTextBox富文本框会提示从此处开始输入信息，点击富文本框提示信息可以清除。输入信息过程中输入错误可以使用窗体上方的撤销、恢复、清空按钮进行操作，同时点击添加图片按钮弹出对话框选择任意图片文件即可导入富文本框。输入完成点击保存，弹出保存文件对话框，输入文件名点击确定即可保存文件，此时窗体会显示该文件名称。
点击关闭可以关闭该文件、清空富文本框的内容，正在编辑文件空。点击打开按钮会显示打开文件对话框，可以选择任意word文件打开并加载到富文本框，窗体显示该文件名。也可以编辑该文件再保存，也可以直接点击删除按钮删除该文件，同时富文本框清空、正在编辑文件空。
# 4.控件/组件
1、标签(System.Windows.Controls.Lable)
2、按钮(System.Windows.Controls.Button）
3、富文本框(System.Windows.Controls..RichTextBox)
4、网格(System.Windows.Controls..Grid)
5、定时器(System.Windows.Threading.DispatcherTimer)
6、打开文件对话框(System.Windows.Forms.OpenFileDialog)
7、保存文件对话框(System.Windows.Forms.SaveFileDialog)
8、消息框(System.Windows.Forms.MessageBox)
9、剪贴板(System.Windows.Forms.Clipboard)
10、文本范围(System.Windows.Documents.TextRange)
11、图像(System.Drawing.Image)
12、MSWORD.OLB(Microsoft.Office.Interop.Word)

# 5.总体设计
窗体利用WPF应用实现，界面设计利用XAML（可扩展应用程序标记语言，Extensible Application Markup Language）实现，程序运行逻辑使用C#语言实现。首先利用Grid网格将窗体分割，接着定义按钮等其他控件并放在网格所规划的位置，并设置它们的各个属性，然后为窗体设计背景图片使得更美观。同时为按钮设置MouseEnter和MouseLeave事件通过改变字体颜色解决鼠标在按钮上方按钮上的文字看不清的问题。
	界面设计完成后设计各个按钮所对应的操作，并将其订阅对应的Click事件，其中保存/打开/添加图片主要利用了savefiledialog/openfiledialog组件打开窗体保存/打开对应的文件，保存文件成功/失败会通过MessageBox显示提示信息。同时还为按钮设计了快捷键，richtetbox显示提示信息，用户点击后会自动清除该提示文本。未来方便用户还在左上角设置Lable标签用于显示正在编辑的文件名，利用定时器定时通过获取类所定义的属性edictingfilename(正在编辑的文件名)刷新Lable的值。
# 6.详细设计
## 6.1界面设计（MainWindow.xaml）
首先设置窗体的标题Title、高度Height、宽度Width以及用来监视键盘输入字符的事件PreviewKeyDown，用于设置按钮快捷键。在对象元素Window.Background里面添加对象元素ImageBrush并设置其附加属性ImageSource设置窗体背景图。接着通过Margin定义一个左上右下边距为(10,0,10,10)、Grid.RowDefinitions和Grid.ColumnDefinitions定义2行17列的Grid网格控件，其中第4行的Height高度设置随着窗口缩小/放大而改变大小，且第1列的宽度为自动使得其根据Lable控件里面的内容设置宽度，并设置最小宽度MinWidth为控件初始长度。
通过按照如下格式设置控件的各个属性：<Button Name="RedoButton" Content="恢复(R)" Grid.Row="0" Grid.Column="4" Height="25" Click="RedoButton_Click" Margin="0,8,0,7" Grid.RowSpan="3" Background="Transparent" BorderBrush="Transparent" Foreground="#FFFDF
DFD" MouseEnter="RedoButton_MouseEnter" MouseLeave="RedoButton_MouseLeave"/>里面的属性分别表示设置一个Button控件，名为RedoButton、内容为恢复(R)、处于网格0行4列的位置、高度25、鼠标点击触发的事件为RedoButton_Click、左上右下边距为(0,8,0,7)、占3行、图标背景颜色为透明、边框颜色为透明、字体颜色为白色、鼠标进入该控件触发的事件为RedoButton_MouseEnter、鼠标移出该控件的范围触发事件RedoButton_MouseLeave，其他的Lable、RichTextBox控件基本也是按照此格式进行定义。
## 6.2 逻辑设计（MainWindow.xaml.cs）
首先定义一个用于保存正在编辑的文件名的变量editingFilename，接着就是用于定时刷新Lable里面的值的DispatcherTimer型的变量timer1，以及用于打开/保存文件的OpenFileDialog 型的openFileDialog、SaveFileDialog型的saveFileDialog。构造函数里面设置保存文件的默认类型为.docx、利用TextRange全选RichTextBox富文本框里面的内容并设置初始的提示信息、设置变量editingFilename为空字符串、将计时器触发的事件的订阅者添加timer1_Tick（用于定时刷新Lable里面的文件名）、timer1.Start()启动该计时器，同时在析构函数里面利用Stop()方法使其停止。

保存按钮SaveButton对应的函数SaveButton_Click：首先判断editingFilename若不为空，则直接调用SaveFile函数保存该文件；若为空，则首先利用TextRange判断RichTextBox里面或者调用保存文件的窗口返回是否是确定，如果不符合条件，则什么不返回；否则设置saveFileDialog返回的文件名设置为editingFilename，若不为空，则调用SaveFile函数保存文件，否则什么不返回。而保存文件函数SaveFile形参为保存的文件名，定义一个Word应用对象以及Word文档对象并将该文档添加到该应用，利用SelectAll以及Copy方法将richTextBox1里面的数据（文本和图片）复制到粘贴板，利用myDoc的Paste方法将其粘贴。由于richTextBox1的文字为白色，保存到word文档后会看不见，所以全选文档里面的文字利用myApp.Selection.Font.Color = WdColor.wdColorBlack将其字体颜色设置成黑色，接着再执行保存、关闭、退出的操作。同时利用MessageBox.Show的方法提示保存成功或者失败，若失败则输出失败的原因。

打开按钮OpenButton对应的函数OpenButton_Click：首先还是利用TextRange将richTextBox1的内容设置为空，还是调用打开文件对话框并且返回确定时接着接下来的操作，否则什么不返回。接着设置editingFilename为savefiledialog返回的内容，利用定义的Word应用对象以及Word文档对象的Open方法打开文件并全选里面的数据，同时由于文档里面的字体为黑色在richTextBox1中显示不清，所以利用上面相同的方法将其设置为白色的字体，并复制到剪贴板上，然后粘贴、关闭、退出，若出现异常则调用MessageBox显示出异常的原因。

添加图片按钮AddPhotoButton对应的函数AddPhotoButton_Click：利用openPhotoFileDialog.Filter设置对话框的过滤条件，同样调用打开文件对话框获取选择的图片，并利用Clipboard.SetDataObject(Image.FromFile(openPhotoFileDialog.FileName), false)将图片复制到剪贴板，参数false表示退出程序后不再将图片保留在剪贴板中。

恢复按钮RedoButton对应的函数RedoButton_Click就是调用富文本框的Redo方法；而撤销按钮UndoButton对应的函数UndoButton_Click则是调用Undo方法；清空按钮ClearAllButto对应的函数ClearAllButton_Click则调用richTextBox1.Document.Blocks.Clear()清空里面的内容；关闭文档按钮对应的函数CloseButton_Click将正在编辑的文件名设置为空字符串，并利用TextRange将富文本框里面的内容也设置为空字符串。

定时刷新lable1的内容，用于显示正在编辑/打开的文件名的函数timer1_Tick则是将字符串“正在编辑文件：”以及editingFilename显示在Lable控件上；删除按钮对应的函数DeleteButton_Click，若正在编辑的文件名为空，则什么都不返回，否则利用File.Delete(editingFilename);删除文件，并将正在编辑的文件设置为空，并将richTextBox1的内容清空；而清除richBoxText提示的输入信息的函数PromptInputInfoOfRichTextBox，也就是富文本框PreviewMouseDown事件对应的函数，是将富文本框的内容清空，并利用richTextBox1.PreviewMouseDown -= PromptInputInfoOfRichTextBox取消订阅该事件，使得该方法只执一次，不至于后面一点击富文本框就将所有的内容清除。

而后面的鼠标进入按钮控件，字体变黑；鼠标出来，字体恢复，其实就是各个按钮对应的MouseEnter和MouseLeave事件分别设置成Button.Foreground = System.Windows.Media.Brushes.Black、Button.Foreground = System.Windows.Media.Brushes.White；而后面的快捷键操作是监视键盘输入，将对应的按键绑定到操作对应的按钮上，具体做法if ((e.KeyboardDevice.IsKeyDown(Key.LeftCtrl) || e.KeyboardDevice.IsKeyDown(Key.RightCtrl)) && e.KeyboardDevice.IsKeyDown(Key.U)) {UndoButton_Click(null, null); }，也就是表示若键盘按Ctrl+U就执行撤销按钮对应的函数。

