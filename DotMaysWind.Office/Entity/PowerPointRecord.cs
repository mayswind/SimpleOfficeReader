using System;
using System.Collections.Generic;

namespace DotMaysWind.Office.Entity
{
    public enum RecordType : uint
    {
        Unknown = 0,
        DocumentContainer = 0x03E8,
        ListWithTextContainer = 0x0FF0,
        PersistAtom = 0x03F3,
        TextHeaderAtom = 0x0F9F,
        EndDocumentAtom = 0x03EA,
        MainMasterContainer = 0x03F8,
        DrawingContainer = 0x040C,
        SlideContainer = 0x03EE,
        HeadersFootersContainer = 0x0FD9,
        SlideAtom = 0x03EF,
        NotesContainer = 0x03F0,
        TextCharsAtom = 0x0FA0,
        TextBytesAtom = 0x0FA8,
        CString = 0x0FBA
    }

    public class PowerPointRecord
    {
        #region 常量
        private const UInt32 ContainerRecordVersion = 0xF;
        #endregion

        #region 字段
        private UInt16 _recVer;
        private UInt16 _recInstance;
        private RecordType _recType;
        private UInt32 _recLen;
        private Int64 _offset;

        private Int32 _deepth;
        private PowerPointRecord _parent;
        private List<PowerPointRecord> _children;
        #endregion

        #region 属性
        /// <summary>
        /// 获取RecordVersion
        /// </summary>
        public UInt16 RecordVersion
        {
            get { return this._recVer; }
        }

        /// <summary>
        /// 获取RecordInstance
        /// </summary>
        public UInt16 RecordInstance
        {
            get { return this._recInstance; }
        }

        /// <summary>
        /// 获取Record类型
        /// </summary>
        public RecordType RecordType
        {
            get { return this._recType; }
        }

        /// <summary>
        /// 获取Record内容大小
        /// </summary>
        public UInt32 RecordLength
        {
            get { return this._recLen; }
        }
        
        /// <summary>
        /// 获取Record相对PowerPoint Document偏移
        /// </summary>
        public Int64 Offset
        {
            get { return this._offset; }
        }

        /// <summary>
        /// 获取Record深度
        /// </summary>
        public Int32 Deepth
        {
            get { return this._deepth; }
        }

        /// <summary>
        /// 获取Record的父节点
        /// </summary>
        public PowerPointRecord Parent
        {
            get { return this._parent; }
        }

        /// <summary>
        /// 获取Record的子节点
        /// </summary>
        public List<PowerPointRecord> Children
        {
            get { return this._children; }
        }
        #endregion

        #region 构造函数
        /// <summary>
        /// 初始化新的Record
        /// </summary>
        /// <param name="parent">父节点</param>
        /// <param name="version">RecordVersion和Instance</param>
        /// <param name="type">Record类型</param>
        /// <param name="length">Record内容大小</param>
        /// <param name="offset">Record相对PowerPoint Document偏移</param>
        public PowerPointRecord(PowerPointRecord parent, UInt16 version, UInt16 type, UInt32 length, Int64 offset)
        {
            this._recVer = (UInt16)(version & 0xF);
            this._recInstance = (UInt16)((version & 0xFFF0) >> 4);
            this._recType = (RecordType)type;
            this._recLen = length;
            this._offset = offset;
            this._deepth = (parent == null ? 0 : parent._deepth + 1);
            this._parent = parent;

            if (_recVer == ContainerRecordVersion)
            {
                this._children = new List<PowerPointRecord>();
            }
        }
        #endregion

        #region 方法
        public void AddChild(PowerPointRecord entry)
        {
            if (this._children == null)
            {
                this._children = new List<PowerPointRecord>();
            }

            this._children.Add(entry);
        }
        #endregion
    }
}