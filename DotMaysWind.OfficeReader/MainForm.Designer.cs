namespace DotMaysWind.OfficeReader
{
    partial class MainForm
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
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tbInformation = new System.Windows.Forms.TextBox();
            this.mnuMain = new System.Windows.Forms.MenuStrip();
            this.mnuOpen = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuReadDocSummary = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuReadSummary = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuReadContent = new System.Windows.Forms.ToolStripMenuItem();
            this.dlgOpen = new System.Windows.Forms.OpenFileDialog();
            this.mnuClose = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // tbInformation
            // 
            this.tbInformation.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbInformation.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbInformation.Location = new System.Drawing.Point(0, 25);
            this.tbInformation.Multiline = true;
            this.tbInformation.Name = "tbInformation";
            this.tbInformation.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tbInformation.Size = new System.Drawing.Size(784, 416);
            this.tbInformation.TabIndex = 0;
            // 
            // mnuMain
            // 
            this.mnuMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuOpen,
            this.mnuClose,
            this.mnuReadDocSummary,
            this.mnuReadSummary,
            this.mnuReadContent});
            this.mnuMain.Location = new System.Drawing.Point(0, 0);
            this.mnuMain.Name = "mnuMain";
            this.mnuMain.Size = new System.Drawing.Size(784, 25);
            this.mnuMain.TabIndex = 1;
            this.mnuMain.Text = "menuStrip1";
            // 
            // mnuOpen
            // 
            this.mnuOpen.Name = "mnuOpen";
            this.mnuOpen.Size = new System.Drawing.Size(75, 21);
            this.mnuOpen.Text = "&Open File";
            this.mnuOpen.Click += new System.EventHandler(this.mnuOpen_Click);
            // 
            // mnuReadDocSummary
            // 
            this.mnuReadDocSummary.Name = "mnuReadDocSummary";
            this.mnuReadDocSummary.Size = new System.Drawing.Size(172, 21);
            this.mnuReadDocSummary.Text = "Show &Document Summary";
            this.mnuReadDocSummary.Click += new System.EventHandler(this.mnuReadDocSummary_Click);
            // 
            // mnuReadSummary
            // 
            this.mnuReadSummary.Name = "mnuReadSummary";
            this.mnuReadSummary.Size = new System.Drawing.Size(109, 21);
            this.mnuReadSummary.Text = "Show &Summary";
            this.mnuReadSummary.Click += new System.EventHandler(this.mnuReadSummary_Click);
            // 
            // mnuReadContent
            // 
            this.mnuReadContent.Name = "mnuReadContent";
            this.mnuReadContent.Size = new System.Drawing.Size(100, 21);
            this.mnuReadContent.Text = "Show Co&ntent";
            this.mnuReadContent.Click += new System.EventHandler(this.mnuReadContent_Click);
            // 
            // mnuClose
            // 
            this.mnuClose.Name = "mnuClose";
            this.mnuClose.Size = new System.Drawing.Size(75, 21);
            this.mnuClose.Text = "&Close File";
            this.mnuClose.Click += new System.EventHandler(this.mnuClose_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 441);
            this.Controls.Add(this.tbInformation);
            this.Controls.Add(this.mnuMain);
            this.MainMenuStrip = this.mnuMain;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Simple Office Reader";
            this.mnuMain.ResumeLayout(false);
            this.mnuMain.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tbInformation;
        private System.Windows.Forms.MenuStrip mnuMain;
        private System.Windows.Forms.ToolStripMenuItem mnuOpen;
        private System.Windows.Forms.ToolStripMenuItem mnuReadSummary;
        private System.Windows.Forms.ToolStripMenuItem mnuReadDocSummary;
        private System.Windows.Forms.OpenFileDialog dlgOpen;
        private System.Windows.Forms.ToolStripMenuItem mnuReadContent;
        private System.Windows.Forms.ToolStripMenuItem mnuClose;
    }
}

