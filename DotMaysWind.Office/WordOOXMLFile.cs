using System;
using System.IO.Packaging;
using System.Xml;
using System.Text;

using DotMaysWind.Office.Helper;

namespace DotMaysWind.Office
{
    public class WordOOXMLFile : OfficeOpenXMLFile, IWordFile
    {
        #region 常量
        private const String WordNameSpace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        #endregion

        #region 字段
        private String _paragraphText;
        private String _footnoteText;
        private String _headerText;
        private String _commentText;
        private String _endnoteText;
        #endregion

        #region 属性
        /// <summary>
        /// 获取文档正文内容
        /// </summary>
        public String ParagraphText
        {
            get { return this._paragraphText; }
        }

        /// <summary>
        /// 获取文档页眉和页脚内容
        /// </summary>
        public String HeaderAndFooterText
        {
            get { return this._headerText; }
        }

        /// <summary>
        /// 获取文档批注内容
        /// </summary>
        public String CommentText
        {
            get { return this._commentText; }
        }

        /// <summary>
        /// 获取文档脚注内容
        /// </summary>
        public String FootnoteText
        {
            get { return this._footnoteText; }
        }

        /// <summary>
        /// 获取文档尾注内容
        /// </summary>
        public String EndnoteText
        {
            get { return this._endnoteText; }
        }
        #endregion

        #region 构造函数
        /// <summary>
        /// 初始化DocxFile
        /// </summary>
        /// <param name="filePath">文件路径</param>
        public WordOOXMLFile(String filePath) :
            base(filePath) { }
        #endregion

        #region 读取内容
        protected override void ReadContent()
        {
            if (this._package == null)
            {
                return;
            }

            this._paragraphText = this.ReadXmlText("/word/document.xml", "/w:document/w:body");
            this._commentText = this.ReadXmlText("/word/comments.xml", "/w:comments");
            this._footnoteText = this.ReadXmlText("/word/footnotes.xml", "/w:footnotes");
            this._endnoteText = this.ReadXmlText("/word/endnotes.xml", "/w:endnotes");
            this._headerText = "";

            Int32 index = 1;
            String filePath = "";

            while (this._package.PartExists(new Uri(filePath = String.Format("/word/header{0}.xml", (index++).ToString()), UriKind.Relative)))
            {
                this._headerText += this.ReadXmlText(filePath, "/w:hdr");
            }

            index = 1;
            while (this._package.PartExists(new Uri(filePath = String.Format("/word/footer{0}.xml", (index++).ToString()), UriKind.Relative)))
            {
                this._headerText += this.ReadXmlText(filePath, "/w:ftr");
            }
        }
        #endregion

        #region 读取文本
        private String ReadXmlText(String uriPath, String nodePath)
        {
            if (!this._package.PartExists(new Uri(uriPath, UriKind.Relative)))
            {
                return String.Empty;
            }

            PackagePart part = this._package.GetPart(new Uri(uriPath, UriKind.Relative));
            StringBuilder content = new StringBuilder();
            XmlDocument doc = new XmlDocument();

            doc.Load(part.GetStream());

            XmlNamespaceManager nsManager = new XmlNamespaceManager(doc.NameTable);
            nsManager.AddNamespace("w", WordNameSpace);

            XmlNode node = doc.SelectSingleNode(nodePath, nsManager);

            if (node == null)
            {
                return String.Empty;
            }

            content.Append(NodeHelper.ReadNode(node));

            return content.ToString();
        }
        #endregion
    }
}