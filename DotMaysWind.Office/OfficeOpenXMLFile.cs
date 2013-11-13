using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Xml;

using DotMaysWind.Office.Summary;

namespace DotMaysWind.Office
{
    public class OfficeOpenXMLFile : IOfficeFile
    {
        #region 常量
        private const String PropertiesNameSpace = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        private const String CorePropertiesNameSpace = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        #endregion

        #region 字段
        protected FileStream _stream;
        protected Package _package;
        protected Dictionary<String, String> _properties;
        protected Dictionary<String, String> _coreProperties;
        #endregion

        #region 属性
        /// <summary>
        /// 获取DocumentSummaryInformation
        /// </summary>
        public Dictionary<String, String> DocumentSummaryInformation
        {
            get { return this._properties; }
        }

        /// <summary>
        /// 获取SummaryInformation
        /// </summary>
        public Dictionary<String, String> SummaryInformation
        {
            get { return this._coreProperties; }
        }
        #endregion

        #region 构造函数
        /// <summary>
        /// 初始化OfficeOpenXMLFile
        /// </summary>
        /// <param name="filePath">文件路径</param>
        public OfficeOpenXMLFile(String filePath)
        {
            try
            {
                this._stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                this._package = Package.Open(this._stream);

                this.ReadProperties();
                this.ReadCoreProperties();
                this.ReadContent();
            }
            finally
            {
                if (this._package != null)
                {
                    this._package.Close();
                }

                if (this._stream != null)
                {
                    this._stream.Close();
                }
            }
        }
        #endregion

        #region 读取Properties
        private void ReadProperties()
        {
            if (this._package == null)
            {
                return;
            }

            PackagePart part = this._package.GetPart(new Uri("/docProps/app.xml", UriKind.Relative));
            if (part == null)
            {
                return;
            }

            XmlDocument doc = new XmlDocument();
            doc.Load(part.GetStream());

            XmlNodeList nodes = doc.GetElementsByTagName("Properties", PropertiesNameSpace);
            if (nodes.Count < 1)
            {
                return;
            }

            this._properties = new Dictionary<String, String>();
            foreach (XmlElement element in nodes[0])
            {
                this._properties.Add(element.LocalName, element.InnerText);
            }
        }
        #endregion

        #region 读取CoreProperties
        private void ReadCoreProperties()
        {
            if (this._package == null)
            {
                return;
            }

            PackagePart part = this._package.GetPart(new Uri("/docProps/core.xml", UriKind.Relative));
            if (part == null)
            {
                return;
            }

            XmlDocument doc = new XmlDocument();
            doc.Load(part.GetStream());

            XmlNodeList nodes = doc.GetElementsByTagName("coreProperties", CorePropertiesNameSpace);
            if (nodes.Count < 1)
            {
                return;
            }
            
            this._coreProperties = new Dictionary<String, String>();
            foreach (XmlElement element in nodes[0])
            {
                this._coreProperties.Add(element.LocalName, element.InnerText);
            }
        }
        #endregion

        #region 读取正文内容
        protected virtual void ReadContent()
        {
            //Do Nothing
        }
        #endregion
    }
}