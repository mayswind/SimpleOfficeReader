using System;
using System.Collections.Generic;

namespace DotMaysWind.Office.Entity
{
    public enum DirectoryEntryType : byte
    {
        Invalid = 0,
        Storage = 1,
        Stream = 2,
        Root = 5
    }

    public class DirectoryEntry
    {
        #region 字段
        private UInt32 _entryID;
        private String _entryName;
        private DirectoryEntryType _entryType;
        private UInt32 _firstSectorID;
        private UInt32 _length;

        private DirectoryEntry _parent;
        private List<DirectoryEntry> _children;
        #endregion

        #region 属性
        /// <summary>
        /// 获取DirectoryEntry的EntryID
        /// </summary>
        public UInt32 EntryID
        {
            get { return this._entryID; }
        }

        /// <summary>
        /// 获取DirectoryEntry名称
        /// </summary>
        public String EntryName
        {
            get { return this._entryName; }
        }

        /// <summary>
        /// 获取DirectoryEntry类型
        /// </summary>
        public DirectoryEntryType EntryType
        {
            get { return this._entryType; }
        }

        /// <summary>
        /// 获取DirectoryEntry的第一个SectorID
        /// </summary>
        public UInt32 FirstSectorID
        {
            get { return this._firstSectorID; }
        }

        /// <summary>
        /// 获取DirectoryEntry的内容大小
        /// </summary>
        public UInt32 Length
        {
            get { return this._length; }
        }

        /// <summary>
        /// 获取DirectoryEntry的父节点
        /// </summary>
        public DirectoryEntry Parent
        {
            get { return this._parent; }
        }

        /// <summary>
        /// 获取DirectoryEntry的子节点
        /// </summary>
        public List<DirectoryEntry> Children
        {
            get { return this._children; }
        }
        #endregion

        #region 构造函数
        /// <summary>
        /// 初始化新的DirectoryEntry
        /// </summary>
        /// <param name="parent">父节点</param>
        /// <param name="entryID">DirectoryEntryID</param>
        /// <param name="entryName">DirectoryEntry名称</param>
        /// <param name="entryType">DirectoryEntry类型</param>
        /// <param name="firstSectorID">第一个SectorID</param>
        /// <param name="length">内容大小</param>
        public DirectoryEntry(DirectoryEntry parent, UInt32 entryID, String entryName, DirectoryEntryType entryType, UInt32 firstSectorID, UInt32 length)
        {
            this._entryID = entryID;
            this._entryName = entryName;
            this._entryType = entryType;
            this._firstSectorID = firstSectorID;
            this._length = length;
            this._parent = parent;

            if (entryType == DirectoryEntryType.Root || entryType == DirectoryEntryType.Storage)
            {
                this._children = new List<DirectoryEntry>();
            }
        }
        #endregion

        #region 方法
        public void AddChild(DirectoryEntry entry)
        {
            if (this._children == null)
            {
                this._children = new List<DirectoryEntry>();
            }

            this._children.Add(entry);
        }

        public DirectoryEntry GetChild(String entryName)
        {
            for (Int32 i = 0; i < this._children.Count; i++)
            {
                if (String.Equals(this._children[i].EntryName, entryName))
                {
                    return this._children[i];
                }
            }

            return null;
        }

        public override String ToString()
        {
            return String.Format("{0}, {1}", base.ToString(), this._entryName);
        }
        #endregion
    }
}