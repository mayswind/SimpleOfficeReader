using System;
using System.Text;
using System.Xml;

namespace DotMaysWind.Office.Helper
{
    internal static class NodeHelper
    {
        /// <summary>
        /// 抽取Node中的文字
        /// </summary>
        /// <param name="node">XmlNode</param>
        /// <returns>Node中的文字</returns>
        internal static String ReadNode(XmlNode node)
        {
            if ((node == null) || (node.NodeType != XmlNodeType.Element))//如果node为空
            {
                return String.Empty;
            }

            StringBuilder nodeContent = new StringBuilder();

            foreach (XmlNode child in node.ChildNodes)
            {
                if (child.NodeType != XmlNodeType.Element)
                {
                    continue;
                }

                switch (child.LocalName)
                {
                    case "t"://正文
                        nodeContent.Append(child.InnerText.TrimEnd());

                        String space = ((XmlElement)child).GetAttribute("xml:space");
                        if ((!String.IsNullOrEmpty(space)) && (space == "preserve")) nodeContent.Append(' ');
                        break;
                    case "cr"://换行符
                    case "br"://换页符
                        nodeContent.Append(Environment.NewLine);
                        break;
                    case "tab"://Tab
                        nodeContent.Append("\t");
                        break;
                    case "p"://段落
                        nodeContent.Append(ReadNode(child));
                        nodeContent.Append(Environment.NewLine);
                        break;
                    default://其他情况
                        nodeContent.Append(ReadNode(child));
                        break;
                }
            }

            return nodeContent.ToString();
        }
    }
}