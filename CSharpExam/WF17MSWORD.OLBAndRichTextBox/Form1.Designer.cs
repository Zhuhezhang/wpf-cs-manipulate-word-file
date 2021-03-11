namespace WF17MSWORD.OLBAndRichTextBox
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SaveButton = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.AddPhotoButtton = new System.Windows.Forms.Button();
            this.ClearallButtton = new System.Windows.Forms.Button();
            this.OpenButton = new System.Windows.Forms.Button();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.RedoButton = new System.Windows.Forms.Button();
            this.UndoButtton = new System.Windows.Forms.Button();
            this.CloseButton = new System.Windows.Forms.Button();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.DeleteButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // richTextBox1
            // 
            this.richTextBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.richTextBox1.Location = new System.Drawing.Point(12, 41);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(888, 474);
            this.richTextBox1.TabIndex = 0;
            this.richTextBox1.Text = "";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(12, 17);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(105, 14);
            this.label1.TabIndex = 1;
            this.label1.Text = "正在编辑文件：";
            // 
            // SaveButton
            // 
            this.SaveButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.SaveButton.Location = new System.Drawing.Point(825, 12);
            this.SaveButton.Name = "SaveButton";
            this.SaveButton.Size = new System.Drawing.Size(81, 23);
            this.SaveButton.TabIndex = 2;
            this.SaveButton.Text = "保存(&B)";
            this.SaveButton.UseVisualStyleBackColor = true;
            this.SaveButton.Click += new System.EventHandler(this.SaveButton_Click);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.DefaultExt = "docx";
            // 
            // AddPhotoButtton
            // 
            this.AddPhotoButtton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.AddPhotoButtton.Location = new System.Drawing.Point(477, 12);
            this.AddPhotoButtton.Name = "AddPhotoButtton";
            this.AddPhotoButtton.Size = new System.Drawing.Size(81, 23);
            this.AddPhotoButtton.TabIndex = 3;
            this.AddPhotoButtton.Text = "添加图片(&T)";
            this.AddPhotoButtton.UseVisualStyleBackColor = true;
            this.AddPhotoButtton.Click += new System.EventHandler(this.AddPhoto_Click);
            // 
            // ClearallButtton
            // 
            this.ClearallButtton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.ClearallButtton.Location = new System.Drawing.Point(390, 12);
            this.ClearallButtton.Name = "ClearallButtton";
            this.ClearallButtton.Size = new System.Drawing.Size(81, 23);
            this.ClearallButtton.TabIndex = 4;
            this.ClearallButtton.Text = "清空(&Q)";
            this.ClearallButtton.UseVisualStyleBackColor = true;
            this.ClearallButtton.Click += new System.EventHandler(this.ClearallButtton_Click);
            // 
            // OpenButton
            // 
            this.OpenButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.OpenButton.Location = new System.Drawing.Point(738, 12);
            this.OpenButton.Name = "OpenButton";
            this.OpenButton.Size = new System.Drawing.Size(81, 23);
            this.OpenButton.TabIndex = 5;
            this.OpenButton.Text = "打开(&D)";
            this.OpenButton.UseVisualStyleBackColor = true;
            this.OpenButton.Click += new System.EventHandler(this.OpenButton_Click);
            // 
            // RedoButton
            // 
            this.RedoButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.RedoButton.Location = new System.Drawing.Point(303, 12);
            this.RedoButton.Name = "RedoButton";
            this.RedoButton.Size = new System.Drawing.Size(81, 23);
            this.RedoButton.TabIndex = 6;
            this.RedoButton.Text = "恢复(&H)";
            this.RedoButton.UseVisualStyleBackColor = true;
            this.RedoButton.Click += new System.EventHandler(this.RedoButton_Click);
            // 
            // UndoButtton
            // 
            this.UndoButtton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.UndoButtton.Location = new System.Drawing.Point(216, 12);
            this.UndoButtton.Name = "UndoButtton";
            this.UndoButtton.Size = new System.Drawing.Size(81, 23);
            this.UndoButtton.TabIndex = 7;
            this.UndoButtton.Text = "撤销(&C)";
            this.UndoButtton.UseVisualStyleBackColor = true;
            this.UndoButtton.Click += new System.EventHandler(this.UndoButton_Click);
            // 
            // CloseButton
            // 
            this.CloseButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.CloseButton.Location = new System.Drawing.Point(651, 12);
            this.CloseButton.Name = "CloseButton";
            this.CloseButton.Size = new System.Drawing.Size(81, 23);
            this.CloseButton.TabIndex = 8;
            this.CloseButton.Text = "关闭(&G)";
            this.CloseButton.UseVisualStyleBackColor = true;
            this.CloseButton.Click += new System.EventHandler(this.CloseButton_Click);
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // DeleteButton
            // 
            this.DeleteButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.DeleteButton.Location = new System.Drawing.Point(564, 12);
            this.DeleteButton.Name = "DeleteButton";
            this.DeleteButton.Size = new System.Drawing.Size(81, 23);
            this.DeleteButton.TabIndex = 9;
            this.DeleteButton.Text = "删除(&S)";
            this.DeleteButton.UseVisualStyleBackColor = true;
            this.DeleteButton.Click += new System.EventHandler(this.DeleteButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(912, 527);
            this.Controls.Add(this.DeleteButton);
            this.Controls.Add(this.CloseButton);
            this.Controls.Add(this.UndoButtton);
            this.Controls.Add(this.RedoButton);
            this.Controls.Add(this.OpenButton);
            this.Controls.Add(this.ClearallButtton);
            this.Controls.Add(this.AddPhotoButtton);
            this.Controls.Add(this.SaveButton);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.richTextBox1);
            this.ForeColor = System.Drawing.Color.MidnightBlue;
            this.Name = "Form1";
            this.Text = "演示操作WORD文件";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button SaveButton;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button AddPhotoButtton;
        private System.Windows.Forms.Button ClearallButtton;
        private System.Windows.Forms.Button OpenButton;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
        private System.Windows.Forms.Button RedoButton;
        private System.Windows.Forms.Button UndoButtton;
        private System.Windows.Forms.Button CloseButton;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Button DeleteButton;
    }
}

