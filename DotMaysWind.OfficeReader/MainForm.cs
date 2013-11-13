using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

using DotMaysWind.Office;

namespace DotMaysWind.OfficeReader
{
    public partial class MainForm : Form
    {
        #region 常量
        private const String PROGRAM_TITLE = "Simple Office Reader";
        #endregion

        #region 字段
        private IOfficeFile _file;
        #endregion

        #region 构造方法
        public MainForm()
        {
            InitializeComponent();
        }
        #endregion

        #region 菜单事件
        private void mnuOpen_Click(object sender, EventArgs e)
        {
            if (this.dlgOpen.ShowDialog() == DialogResult.OK && !String.IsNullOrEmpty(this.dlgOpen.FileName))
            {
                this.OpenFile(this.dlgOpen.FileName);
            }
        }

        private void mnuClose_Click(object sender, EventArgs e)
        {
            this.CloseFile();
        }

        private void mnuReadDocSummary_Click(object sender, EventArgs e)
        {
            if (this._file == null)
            {
                this.tbInformation.Text = String.Format("Please open the file first.");
                return;
            }

            this.ShowSummary(this._file.DocumentSummaryInformation);
        }

        private void mnuReadSummary_Click(object sender, EventArgs e)
        {
            if (this._file == null)
            {
                this.tbInformation.Text = String.Format("Please open the file first.");
                return;
            }

            this.ShowSummary(this._file.SummaryInformation);
        }

        private void mnuReadContent_Click(object sender, EventArgs e)
        {
            if (this._file == null)
            {
                this.tbInformation.Text = String.Format("Please open the file first.");
            }

            this.ShowContent(this._file);
        }
        #endregion

        #region 私有方法
        private void OpenFile(String filePath)
        {
            try
            {
                this._file = OfficeFileFactory.CreateOfficeFile(this.dlgOpen.FileName);
                this.tbInformation.Text = String.Format((this._file == null ? "Failed to open \"{0}\"." : ""), this.dlgOpen.FileName);
                this.Text = String.Format("{0} - {1}", filePath, PROGRAM_TITLE);
            }
            catch (Exception ex)
            {
                this.CloseFile();
                this.tbInformation.Text = String.Format("Error: {0}", ex.Message);
            }
        }

        private void CloseFile()
        {
            this._file = null;
            this.tbInformation.Text = "";
            this.Text = PROGRAM_TITLE;
        }

        private void ShowSummary(Dictionary<String, String> dictionary)
        {
            if (dictionary == null)
            {
                this.tbInformation.Text = String.Format("This file is not Microsoft Office file.");
                return;
            }

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<String, String> pair in dictionary)
            {
                sb.AppendFormat("[{0}]={1}", pair.Key, pair.Value);
                sb.AppendLine();
            }

            this.tbInformation.Text = sb.ToString();
        }

        private void ShowContent(IOfficeFile file)
        {
            if (file is IWordFile)
            {
                IWordFile wordFile = file as IWordFile;
                this.tbInformation.Text = wordFile.ParagraphText;
            }
            else if (file is IPowerPointFile)
            {
                IPowerPointFile pptFile = file as IPowerPointFile;
                this.tbInformation.Text = pptFile.AllText;
            }
            else
            {
                this.tbInformation.Text = String.Format("Cannot extract content from this file.");
            }
        }
        #endregion
    }
}