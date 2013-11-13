using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

using DotMaysWind.Office.Entity;
using DotMaysWind.Office.Summary;

namespace DotMaysWind.Office
{
    public class CompoundBinaryFile : IOfficeFile
    {
        #region 常量
        private const UInt32 HeaderSize = 0x200;//512字节
        private const UInt32 DirectoryEntrySize = 0x80;//128字节
        protected const UInt32 MaxRegSector = 0xFFFFFFFA;
        protected const UInt32 DifSector = 0xFFFFFFFC;
        protected const UInt32 FatSector = 0xFFFFFFFD;
        protected const UInt32 EndOfChain = 0xFFFFFFFE;
        protected const UInt32 FreeSector = 0xFFFFFFFF;
        #endregion

        #region 字段
        protected FileStream _stream;
        protected BinaryReader _reader;
        protected Int64 _length;
        protected List<UInt32> _fatSectors;
        protected List<UInt32> _minifatSectors;
        protected List<UInt32> _miniSectors;
        protected List<UInt32> _dirSectors;
        protected DirectoryEntry _dirRootEntry;
        protected List<DocumentSummaryInformation> _documentSummaryInformation;
        protected List<SummaryInformation> _summaryInformation;

        #region 头部信息
        private UInt32 _sectorSize;//Sector大小
        private UInt32 _miniSectorSize;//Mini-Sector大小
        private UInt32 _fatCount;//FAT数量
        private UInt32 _dirStartSectorID;//Directory开始的SectorID
        private UInt32 _miniCutoffSize;//Mini-Sector最大的大小
        private UInt32 _miniFatStartSectorID;//Mini-FAT开始的SectorID
        private UInt32 _miniFatCount;//Mini-FAT数量
        private UInt32 _difStartSectorID;//DIF开始的SectorID
        private UInt32 _difCount;//DIF数量
        #endregion
        #endregion

        #region 属性
        /// <summary>
        /// 获取DocumentSummaryInformation
        /// </summary>
        public Dictionary<String, String> DocumentSummaryInformation
        {
            get
            {
                if (this._documentSummaryInformation == null)
                {
                    return null;
                }

                Dictionary<String, String> dict = new Dictionary<String, String>();
                for (Int32 i = 0; i < this._documentSummaryInformation.Count; i++)
                {
                    dict.Add(this._documentSummaryInformation[i].Type.ToString(), this._documentSummaryInformation[i].Data.ToString());
                }

                return dict;
            }
        }

        /// <summary>
        /// 获取SummaryInformation
        /// </summary>
        public Dictionary<String, String> SummaryInformation
        {
            get
            {
                if (this._summaryInformation == null)
                {
                    return null;
                }

                Dictionary<String, String> dict = new Dictionary<String, String>();

                for (Int32 i = 0; i < this._summaryInformation.Count; i++)
                {
                    dict.Add(this._summaryInformation[i].Type.ToString(), this._summaryInformation[i].Data.ToString());
                }

                return dict;
            }
        }
        #endregion

        #region 构造函数
        /// <summary>
        /// 初始化CompoundBinaryFile
        /// </summary>
        /// <param name="filePath">文件路径</param>
        public CompoundBinaryFile(String filePath)
        {
            try
            {
                this._stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                this._reader = new BinaryReader(this._stream);

                this._length = this._stream.Length;

                this.ReadHeader();
                this.ReadFAT();
                this.ReadDirectory();
                this.ReadMiniFAT();
                this.ReadDocumentSummaryInformation();
                this.ReadSummaryInformation();
                this.ReadContent();
            }
            finally
            {
                if (this._reader != null)
                {
                    this._reader.Close();
                }

                if (this._stream != null)
                {
                    this._stream.Close();
                }
            }
        }
        #endregion

        #region 读取头部信息
        private void ReadHeader()
        {
            if (this._reader == null)
            {
                return;
            }

            //先判断是否是Compound Binary文件格式
            Byte[] sig = (this._length > 512 ? this._reader.ReadBytes(8) : null);
            if (sig == null ||
                sig[0] != 0xD0 || sig[1] != 0xCF || sig[2] != 0x11 || sig[3] != 0xE0 ||
                sig[4] != 0xA1 || sig[5] != 0xB1 || sig[6] != 0x1A || sig[7] != 0xE1)
            {
                throw new Exception("This file is not compound binary file.");
            }

            //读取头部信息
            this._stream.Seek(22, SeekOrigin.Current);
            this._sectorSize = (UInt32)Math.Pow(2, this._reader.ReadUInt16());
            this._miniSectorSize = (UInt32)Math.Pow(2, this._reader.ReadUInt16());

            this._stream.Seek(10, SeekOrigin.Current);
            this._fatCount = this._reader.ReadUInt32();
            this._dirStartSectorID = this._reader.ReadUInt32();

            this._stream.Seek(4, SeekOrigin.Current);
            this._miniCutoffSize = this._reader.ReadUInt32();
            this._miniFatStartSectorID = this._reader.ReadUInt32();
            this._miniFatCount = this._reader.ReadUInt32();
            this._difStartSectorID = this._reader.ReadUInt32();
            this._difCount = this._reader.ReadUInt32();
        }
        #endregion

        #region 读取FAT
        private void ReadFAT()
        {
            if (this._fatCount > 0)
            {
                this._fatSectors = new List<UInt32>();
                this.ReadFirst109FatSectors();
            }

            if (this._difCount > 0)
            {
                this.ReadLastFatSectors();
            }

            if (this._fatCount != this._fatSectors.Count)
            {
                throw new Exception("File has been broken (FAT count is INVALID).");
            }
        }

        private void ReadFirst109FatSectors()
        {
            for (Int32 i = 0; i < 109; i++)
            {
                UInt32 nextSector = this._reader.ReadUInt32();

                if (nextSector == CompoundBinaryFile.FreeSector)
                {
                    break;
                }

                this._fatSectors.Add(nextSector);
            }
        }

        private void ReadLastFatSectors()
        {
            UInt32 difSectorID = this._difStartSectorID;

            while (true)
            {
                Int64 entryStart = this.GetSectorOffset(difSectorID);
                this._stream.Seek(entryStart, SeekOrigin.Begin);

                for (Int32 i = 0; i < 127; i++)
                {
                    UInt32 fatSectorID = this._reader.ReadUInt32();

                    if (fatSectorID == CompoundBinaryFile.FreeSector)
                    {
                        return;
                    }

                    this._fatSectors.Add(fatSectorID);
                }

                difSectorID = this._reader.ReadUInt32();
                if (difSectorID == CompoundBinaryFile.EndOfChain)
                {
                    break;
                }
            }
        }
        #endregion

        #region 读取目录信息
        private void ReadDirectory()
        {
            if (this._reader == null)
            {
                return;
            }

            this._dirSectors = new List<UInt32>();
            UInt32 sectorID = this._dirStartSectorID;

            while (true)
            {
                this._dirSectors.Add(sectorID);
                sectorID = this.GetNextSectorID(sectorID);

                if (sectorID == CompoundBinaryFile.EndOfChain)
                {
                    break;
                }
            }

            UInt32 leftSiblingEntryID, rightSiblingEntryID, childEntryID;
            this._dirRootEntry = GetDirectoryEntry(0, null, out leftSiblingEntryID, out rightSiblingEntryID, out childEntryID);
            this.ReadDirectoryEntry(this._dirRootEntry, childEntryID);
        }

        private void ReadDirectoryEntry(DirectoryEntry rootEntry, UInt32 entryID)
        {
            UInt32 leftSiblingEntryID, rightSiblingEntryID, childEntryID;
            DirectoryEntry entry = GetDirectoryEntry(entryID, rootEntry, out leftSiblingEntryID, out rightSiblingEntryID, out childEntryID);

            if (entry == null || entry.EntryType == DirectoryEntryType.Invalid)
            {
                return;
            }
            
            rootEntry.AddChild(entry);

            if (leftSiblingEntryID < UInt32.MaxValue)//有左兄弟节点
            {
                this.ReadDirectoryEntry(rootEntry, leftSiblingEntryID);
            }

            if (rightSiblingEntryID < UInt32.MaxValue)//有右兄弟节点
            {
                this.ReadDirectoryEntry(rootEntry, rightSiblingEntryID);
            }

            if (childEntryID < UInt32.MaxValue)//有孩子节点
            {
                this.ReadDirectoryEntry(entry, childEntryID);
            }
        }

        private DirectoryEntry GetDirectoryEntry(UInt32 entryID, DirectoryEntry parentEntry, out UInt32 leftSiblingEntryID, out UInt32 rightSiblingEntryID, out UInt32 childEntryID)
        {
            leftSiblingEntryID = UInt16.MaxValue;
            rightSiblingEntryID = UInt16.MaxValue;
            childEntryID = UInt16.MaxValue;

            this._stream.Seek(GetDirectoryEntryOffset(entryID), SeekOrigin.Begin);

            if (this._stream.Position >= this._length)
            {
                return null;
            }

            StringBuilder temp = new StringBuilder();
            for (Int32 i = 0; i < 32; i++)
            {
                temp.Append((Char)this._reader.ReadUInt16());
            }

            UInt16 nameLen = this._reader.ReadUInt16();
            String name = (temp.ToString(0, (temp.Length < (nameLen / 2 - 1) ? temp.Length : nameLen / 2 - 1)));
            Byte type = this._reader.ReadByte();

            if (type > 5)
            {
                return null;
            }

            this._stream.Seek(1, SeekOrigin.Current);
            leftSiblingEntryID = this._reader.ReadUInt32();
            rightSiblingEntryID = this._reader.ReadUInt32();
            childEntryID = this._reader.ReadUInt32();

            this._stream.Seek(36, SeekOrigin.Current);
            UInt32 sectorID = this._reader.ReadUInt32();
            UInt32 length = this._reader.ReadUInt32();

            return new DirectoryEntry(parentEntry, entryID, name, (DirectoryEntryType)type, sectorID, length);
        }
        #endregion

        #region 读取MiniFAT
        private void ReadMiniFAT()
        {
            if (this._miniFatCount > 0)
            {
                this._minifatSectors = new List<UInt32>();
                this.ReadMiniFatSectors();
            }

            if (this._minifatSectors != null && this._minifatSectors.Count > 0)
            {
                this._miniSectors = new List<UInt32>();
                this.ReadMiniSectors();
            }

            if (this._miniSectors != null && this._miniSectors.Count != Math.Ceiling((Double)this._dirRootEntry.Length / this._sectorSize))
            {
                throw new Exception("File has been broken (mini-FAT count is INVALID)");
            }
        }

        private void ReadMiniFatSectors()
        {
            UInt32 sectorID = this._miniFatStartSectorID;

            while (true)
            {
                this._minifatSectors.Add(sectorID);
                sectorID = this.GetNextSectorID(sectorID);

                if (sectorID == CompoundBinaryFile.EndOfChain)
                {
                    break;
                }
            }
        }

        private void ReadMiniSectors()
        {
            UInt32 sectorID = this._dirRootEntry.SectorID;

            while (true)
            {
                this._miniSectors.Add(sectorID);
                sectorID = this.GetNextSectorID(sectorID);

                if (sectorID == CompoundBinaryFile.EndOfChain)
                {
                    break;
                }
            }
        }
        #endregion

        #region 读取DocumentSummaryInformation
        private void ReadDocumentSummaryInformation()
        {
            DirectoryEntry entry = this._dirRootEntry.GetChild('\x05' + "DocumentSummaryInformation");

            if (entry == null)
            {
                return;
            }

            Int64 entryStart = this.GetEntryOffset(entry);

            this._stream.Seek(entryStart + 24, SeekOrigin.Begin);
            UInt32 propertysCount = this._reader.ReadUInt32();
            UInt32 docSumamryStart = 0;

            for (Int32 i = 0; i < propertysCount; i++)
            {
                Byte[] clsid = this._reader.ReadBytes(16);
                if (clsid.Length == 16 && 
                    clsid[0] == 0x02 && clsid[1] == 0xD5 && clsid[2] == 0xCD && clsid[3] == 0xD5 &&
                    clsid[4] == 0x9C && clsid[5] == 0x2E && clsid[6] == 0x1B && clsid[7] == 0x10 &&
                    clsid[8] == 0x93 && clsid[9] == 0x97 && clsid[10] == 0x08 && clsid[11] == 0x00 &&
                    clsid[12] == 0x2B && clsid[13] == 0x2C && clsid[14] == 0xF9 && clsid[15] == 0xAE)//如果是DocumentSummaryInformation
                {
                    docSumamryStart = this._reader.ReadUInt32();
                    break;
                }
                else
                {
                    //this._stream.Seek(4, SeekOrigin.Current);
                    return;
                }
            }

            if (docSumamryStart == 0)
            {
                return;
            }

            this._stream.Seek(entryStart + docSumamryStart, SeekOrigin.Begin);
            this._documentSummaryInformation = new List<DocumentSummaryInformation>();
            UInt32 docSummarySize = this._reader.ReadUInt32();
            UInt32 docSummaryCount = this._reader.ReadUInt32();
            Int64 offsetMark = this._stream.Position;
            Int32 codePage = Encoding.Default.CodePage;

            for (Int32 i = 0; i < docSummaryCount; i++)
            {
                if (offsetMark >= this._stream.Length)
                {
                    break;
                }

                this._stream.Seek(offsetMark, SeekOrigin.Begin);
                UInt32 propertyID = this._reader.ReadUInt32();
                UInt32 properyOffset = this._reader.ReadUInt32();

                offsetMark = this._stream.Position;

                this._stream.Seek(entryStart + docSumamryStart + properyOffset, SeekOrigin.Begin);

                if (this._stream.Position > this._stream.Length)
                {
                    continue;
                }

                this._stream.Seek(entryStart + docSumamryStart + properyOffset, SeekOrigin.Begin);
                UInt32 propertyType = this._reader.ReadUInt32();
                DocumentSummaryInformation info = null;
                Byte[] data = null;

                if (propertyType == 0x1E)
                {
                    UInt32 strLen = this._reader.ReadUInt32();
                    data = this._reader.ReadBytes((Int32)strLen);
                    info = new DocumentSummaryInformation(propertyID, propertyType, codePage, data);
                }
                else
                {
                    data = this._reader.ReadBytes(4);
                    info = new DocumentSummaryInformation(propertyID, propertyType, data);

                    if (info.Type == DocumentSummaryInformationType.CodePage && info.Data != null)//如果找到CodePage的属性
                    {
                        codePage = (Int32)(UInt16)info.Data;
                    }
                }

                if (info.Data != null)
                {
                    this._documentSummaryInformation.Add(info);
                }
            }
        }
        #endregion

        #region 读取SummaryInformation
        private void ReadSummaryInformation()
        {
            DirectoryEntry entry = this._dirRootEntry.GetChild('\x05' + "SummaryInformation");

            if (entry == null)
            {
                return;
            }

            Int64 entryStart = this.GetEntryOffset(entry);

            this._stream.Seek(entryStart + 24, SeekOrigin.Begin);
            UInt32 propertysCount = this._reader.ReadUInt32();
            UInt32 docSumamryStart = 0;

            for (Int32 i = 0; i < propertysCount; i++)
            {
                Byte[] clsid = this._reader.ReadBytes(16);
                if (clsid.Length == 16 &&
                    clsid[0] == 0xE0 && clsid[1] == 0x85 && clsid[2] == 0x9F && clsid[3] == 0xF2 &&
                    clsid[4] == 0xF9 && clsid[5] == 0x4F && clsid[6] == 0x68 && clsid[7] == 0x10 &&
                    clsid[8] == 0xAB && clsid[9] == 0x91 && clsid[10] == 0x08 && clsid[11] == 0x00 &&
                    clsid[12] == 0x2B && clsid[13] == 0x27 && clsid[14] == 0xB3 && clsid[15] == 0xD9)//如果是SummaryInformation
                {
                    docSumamryStart = this._reader.ReadUInt32();
                    break;
                }
                else
                {
                    //this._stream.Seek(4, SeekOrigin.Current);
                    return;
                }
            }

            if (docSumamryStart == 0)
            {
                return;
            }

            this._stream.Seek(entryStart + docSumamryStart, SeekOrigin.Begin);
            this._summaryInformation = new List<SummaryInformation>();
            UInt32 docSummarySize = this._reader.ReadUInt32();
            UInt32 docSummaryCount = this._reader.ReadUInt32();
            Int64 offsetMark = this._stream.Position;
            Int32 codePage = Encoding.Default.CodePage;

            for (Int32 i = 0; i < docSummaryCount; i++)
            {
                if (offsetMark >= this._stream.Length)
                {
                    break;
                }

                this._stream.Seek(offsetMark, SeekOrigin.Begin);
                UInt32 propertyID = this._reader.ReadUInt32();
                UInt32 properyOffset = this._reader.ReadUInt32();

                offsetMark = this._stream.Position;

                this._stream.Seek(entryStart + docSumamryStart + properyOffset, SeekOrigin.Begin);

                if (this._stream.Position > this._stream.Length)
                {
                    continue;
                }

                UInt32 propertyType = this._reader.ReadUInt32();
                SummaryInformation info = null;
                Byte[] data = null;

                if (propertyType == 0x1E)
                {
                    UInt32 strLen = this._reader.ReadUInt32();
                    data = this._reader.ReadBytes((Int32)strLen);
                    info = new SummaryInformation(propertyID, propertyType, codePage, data);
                }
                else if (propertyType == 0x40)
                {
                    data = this._reader.ReadBytes(8);
                    info = new SummaryInformation(propertyID, propertyType, data);
                }
                else
                {
                    data = this._reader.ReadBytes(4);
                    info = new SummaryInformation(propertyID, propertyType, data);

                    if (info.Type == SummaryInformationType.CodePage && info.Data != null)//如果找到CodePage的属性
                    {
                        codePage = (Int32)(UInt16)info.Data;
                    }
                }

                if (info.Data != null)
                {
                    this._summaryInformation.Add(info);
                }
            }
        }
        #endregion

        #region 读取正文内容
        protected virtual void ReadContent()
        {
            //Do Nothing
        }
        #endregion

        #region 辅助方法
        protected UInt32 GetNextSectorID(UInt32 sectorID)
        {
            UInt32 sectorInFile = this._fatSectors[(Int32)(sectorID / 128)];
            this._stream.Seek(this.GetSectorOffset(sectorInFile) + 4 * (sectorID % 128), SeekOrigin.Begin);

            return this._reader.ReadUInt32();
        }

        protected UInt32 GetNextMiniSectorID(UInt32 miniSectorID)
        {
            UInt32 sectorInFile = this._minifatSectors[(Int32)(miniSectorID / 128)];
            this._stream.Seek(this.GetSectorOffset(sectorInFile) + 4 * (miniSectorID % 128), SeekOrigin.Begin);

            return this._reader.ReadUInt32();
        }

        protected Int64 GetEntryOffset(DirectoryEntry entry)
        {
            if (entry.Length >= this._miniCutoffSize)
            {
                return GetSectorOffset(entry.SectorID);
            }
            else
            {
                return GetMiniSectorOffset(entry.SectorID);
            }
        }

        protected Int64 GetSectorOffset(UInt32 sectorID)
        {
            return HeaderSize + this._sectorSize * sectorID;
        }

        protected Int64 GetMiniSectorOffset(UInt32 miniSectorID)
        {
            UInt32 sectorID = this._miniSectors[(Int32)((miniSectorID * this._miniSectorSize) / this._sectorSize)];
            UInt32 offset = (UInt32)((miniSectorID * this._miniSectorSize) % this._sectorSize);

            return HeaderSize + this._sectorSize * sectorID + offset;
        }

        protected Int64 GetDirectoryEntryOffset(UInt32 entryID)
        {
            UInt32 sectorID = this._dirSectors[(Int32)(entryID * CompoundBinaryFile.DirectoryEntrySize / this._sectorSize)];
            return this.GetSectorOffset(sectorID) + (entryID * CompoundBinaryFile.DirectoryEntrySize) % this._sectorSize;
        }
        #endregion
    }
}