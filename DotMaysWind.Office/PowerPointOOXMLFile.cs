using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Text;
using System.Xml;

using DotMaysWind.Office.Helper;

namespace DotMaysWind.Office
{
    public class PowerPointOOXMLFile : OfficeOpenXMLFile, IPowerPointFile
    {
        #region 常量
        private const String PowerPointNameSpace = "http://schemas.openxmlformats.org/presentationml/2006/main";
        #endregion

        #region 字段
        private StringBuilder _allText;
        #endregion

        #region 属性
        /// <summary>
        /// 获取PowerPoint幻灯片中所有文本
        /// </summary>
        public String AllText
        {
            get { return this._allText.ToString(); }
        }
        #endregion

        #region 构造函数
        /// <summary>
        /// 初始化PptxFile
        /// </summary>
        /// <param name="filePath">文件路径</param>
        public PowerPointOOXMLFile(String filePath) :
            base(filePath) { }
        #endregion

        #region 读取内容
        protected override void ReadContent()
        {
            if (this._package == null)
            {
                return;
            }

            this._allText = new StringBuilder();

            XmlDocument doc = null;
            PackagePartCollection col = this._package.GetParts();
            SortedList<Int32, XmlDocument> list = new SortedList<Int32, XmlDocument>();
            
            foreach (PackagePart part in col)
            {
                if (part.Uri.ToString().IndexOf("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) > -1)
                {
                    doc = new XmlDocument();
                    doc.Load(part.GetStream());

                    String pageName = part.Uri.ToString().Replace("/ppt/slides/slide", "").Replace(".xml", "");
                    Int32 index = 0;
                    Int32.TryParse(pageName, out index);

                    list.Add(index, doc);
                }
            }

            foreach (KeyValuePair<Int32, XmlDocument> pair in list)
            {
                XmlNamespaceManager nsManager = new XmlNamespaceManager(doc.NameTable);
                nsManager.AddNamespace("p", PowerPointNameSpace);

                XmlNode node = pair.Value.SelectSingleNode("/p:sld", nsManager);

                if (node == null)
                {
                    continue;
                }

                this._allText.Append(NodeHelper.ReadNode(node));
            }
        }
        #endregion
    }
}