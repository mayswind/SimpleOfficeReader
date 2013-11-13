using System;

namespace DotMaysWind.Office.Entity
{
    public class PieceElement
    {
        #region 字段
        private UInt16 _info;
        private UInt32 _fc;
        private UInt16 _prm;
        private Boolean _isUnicode;
        #endregion

        #region 属性
        /// <summary>
        /// 获取是否以Unicode形式存储文本
        /// </summary>
        public Boolean IsUnicode
        {
            get { return this._isUnicode; }
        }

        /// <summary>
        /// 获取文本偏移量
        /// </summary>
        public UInt32 Offset
        {
            get { return this._fc; }
        }
        #endregion

        #region 构造函数
        public PieceElement(UInt16 info, UInt32 fcCompressed, UInt16 prm)
        {
            this._info = info;
            this._fc = fcCompressed & 0x3FFFFFFF;//后30位
            this._prm = prm;
            this._isUnicode = (fcCompressed & 0x40000000) == 0;//第31位

            if (!this._isUnicode) this._fc = this._fc / 2;
        }
        #endregion
    }
}