using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

using DotMaysWind.Office.Entity;
using DotMaysWind.Office.Helper;

namespace DotMaysWind.Office
{
    public class PowerPointFile : CompoundBinaryFile, IPowerPointFile
    {
        #region 字段
        private MemoryStream _contentStream;
        private BinaryReader _contentReader;

        private List<PowerPointRecord> _records;
        private StringBuilder _allText;
        #region 测试方法
        private StringBuilder _recordTree;
        #endregion
        #endregion

        #region 属性
        /// <summary>
        /// 获取PowerPoint幻灯片中所有文本
        /// </summary>
        public String AllText
        {
            get { return this._allText.ToString(); }
        }

        #region 测试方法
        /// <summary>
        /// 获取PowerPoint中Record的树形结构
        /// </summary>
        public String RecordTree
        {
            get { return this._recordTree.ToString(); }
        }
        #endregion
        #endregion

        #region 构造函数
        /// <summary>
        /// 初始化PptFile
        /// </summary>
        /// <param name="filePath">文件路径</param>
        public PowerPointFile(String filePath) :
            base(filePath) { }
        #endregion

        #region 读取内容
        protected override void ReadContent()
        {
            DirectoryEntry entry = this._dirRootEntry.GetChild("PowerPoint Document");

            if (entry == null)
            {
                return;
            }

            try
            {
                this._contentStream = new MemoryStream(this._entryData[entry.EntryID]);
                this._contentReader = new BinaryReader(this._contentStream);

                #region 测试方法
                this._recordTree = new StringBuilder();
                #endregion

                this._allText = new StringBuilder();
                this._records = new List<PowerPointRecord>();
                PowerPointRecord record = null;

                while (this._contentStream.Position < this._contentStream.Length)
                {
                    record = this.ReadRecord(null);

                    if (record == null || record.RecordType == 0)
                    {
                        break;
                    }
                }

                this._allText = new StringBuilder(StringHelper.ReplaceString(this._allText.ToString()));
            }
            finally
            {
                if (this._contentReader != null)
                {
                    this._contentReader.Close();
                }

                if (this._contentStream != null)
                {
                    this._contentStream.Close();
                }
            }
        }

        private PowerPointRecord ReadRecord(PowerPointRecord parent)
        {
            PowerPointRecord record = GetRecord(parent);

            if (record == null)
            {
                return null;
            }
            #region 测试方法
            else
            {
                this._recordTree.Append('-', record.Deepth * 2);
                this._recordTree.AppendFormat("[{0}]-[{1}]-[Len:{2}]", record.RecordType, record.Deepth, record.RecordLength);
                this._recordTree.AppendLine();
            }
            #endregion

            if (parent == null)
            {
                this._records.Add(record);
            }
            else
            {
                parent.AddChild(record);
            }

            if (record.RecordVersion == 0xF)
            {
                while (this._contentStream.Position < record.Offset + record.RecordLength)
                {
                    this.ReadRecord(record);
                }
            }
            else
            {
                if (record.Parent != null && (
                    record.Parent.RecordType == RecordType.ListWithTextContainer ||
                    record.Parent.RecordType == RecordType.HeadersFootersContainer ||
                    (UInt32)record.Parent.RecordType == 0xF00D))
                {
                    if (record.RecordType == RecordType.TextCharsAtom || record.RecordType == RecordType.CString)//找到Unicode双字节文字内容
                    {
                        Byte[] data = this._contentReader.ReadBytes((Int32)record.RecordLength);
                        this._allText.Append(StringHelper.GetString(true, data));
                        this._allText.AppendLine();

                    }
                    else if (record.RecordType == RecordType.TextBytesAtom)//找到Unicode<256单字节文字内容
                    {
                        Byte[] data = this._contentReader.ReadBytes((Int32)record.RecordLength);
                        this._allText.Append(StringHelper.GetString(false, data));
                        this._allText.AppendLine();
                    }
                    else
                    {
                        this._contentStream.Seek(record.RecordLength, SeekOrigin.Current);
                    }
                }
                else
                {
                    if (this._contentStream.Position + record.RecordLength < this._contentStream.Length)
                    {
                        this._contentStream.Seek(record.RecordLength, SeekOrigin.Current);
                    }
                    else
                    {
                        this._contentStream.Seek(0, SeekOrigin.End);
                    }
                }
            }

            return record;
        }

        private PowerPointRecord GetRecord(PowerPointRecord parent)
        {
            if (this._contentStream.Position + 8 >= this._contentStream.Length)
            {
                return null;
            }

            UInt16 version = this._contentReader.ReadUInt16();
            UInt16 type = this._contentReader.ReadUInt16();
            UInt32 length = this._contentReader.ReadUInt32();

            return new PowerPointRecord(parent, version, type, length, this._contentStream.Position);
        }
        #endregion
    }
}