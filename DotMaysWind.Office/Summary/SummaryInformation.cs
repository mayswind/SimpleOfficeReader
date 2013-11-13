using System;
using System.Text;

namespace DotMaysWind.Office.Summary
{
    public class SummaryInformation
    {
        #region 字段
        private SummaryInformationType _propertyID;
        private Object _data;
        #endregion

        #region 属性
        /// <summary>
        /// 获取属性类型
        /// </summary>
        public SummaryInformationType Type
        {
            get { return this._propertyID; }
        }

        /// <summary>
        /// 获取属性数据
        /// </summary>
        public Object Data
        {
            get { return this._data; }
        }
        #endregion

        #region 构造函数
        /// <summary>
        /// 初始化新的非字符串型SummaryInformation
        /// </summary>
        /// <param name="propertyID">属性ID</param>
        /// <param name="propertyType">属性数据类型</param>
        /// <param name="data">属性数据</param>
        public SummaryInformation(UInt32 propertyID, UInt32 propertyType, Byte[] data)
        {
            this._propertyID = (SummaryInformationType)propertyID;
            if (propertyType == 0x02) this._data = BitConverter.ToUInt16(data, 0);
            else if (propertyType == 0x03) this._data = BitConverter.ToUInt32(data, 0);
            else if (propertyType == 0x0B) this._data = BitConverter.ToBoolean(data, 0);
            else if (propertyType == 0x40) this._data = DateTime.FromFileTime(BitConverter.ToInt64(data, 0));
        }

        /// <summary>
        /// 初始化新的字符串型SummaryInformation
        /// </summary>
        /// <param name="propertyID">属性ID</param>
        /// <param name="propertyType">属性数据类型</param>
        /// <param name="codePage">代码页标识符</param>
        /// <param name="data">属性数据</param>
        public SummaryInformation(UInt32 propertyID, UInt32 propertyType, Int32 codePage, Byte[] data)
        {
            this._propertyID = (SummaryInformationType)propertyID;
            if (propertyType == 0x1E) this._data = Encoding.GetEncoding(codePage).GetString(data).Replace("\0", "");
        }
        #endregion
    }
}