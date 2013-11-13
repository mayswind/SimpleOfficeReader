using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

using DotMaysWind.Office.Entity;
using DotMaysWind.Office.Helper;

namespace DotMaysWind.Office
{
    public sealed class WordBinaryFile : CompoundBinaryFile, IWordFile
    {
        #region 字段
        private UInt16 _nFib;
        private Boolean _isComplexFile;
        private Boolean _hasPictures;
        private Boolean _isEncrypted;
        private Boolean _is1Table;

        private UInt16 _lidFE;

        private Int32 _cbMac;
        private Int32 _ccpText;
        private Int32 _ccpFtn;
        private Int32 _ccpHdd;
        private Int32 _ccpAtn;
        private Int32 _ccpEdn;
        private Int32 _ccpTxbx;
        private Int32 _ccpHdrTxbx;

        private UInt32 _fcClx;
        private UInt32 _lcbClx;

        private List<UInt32> _lstPieceStartPosition;
        private List<UInt32> _lstPieceEndPosition;
        private List<PieceElement> _lstPieceElement;

        private String _paragraphText;
        private String _footnoteText;
        private String _headerText;
        private String _commentText;
        private String _endnoteText;
        private String _textboxText;
        private String _headerTextboxText;
        #endregion

        #region 属性
        /// <summary>
        /// 获取应用程序版本
        /// </summary>
        public String Version
        {
            get
            {
                switch (this._nFib)
                {
                    case 0x00C1: return "Word 97";
                    case 0x00D9: return "Word 2000";
                    case 0x0101: return "Word 2002";
                    case 0x010C: return "Word 2003";
                    case 0x0112: return "Word 2007";
                    default: return String.Empty;
                }
            }
        }

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

        /// <summary>
        /// 获取文档文本框内容
        /// </summary>
        public String TextboxText
        {
            get { return this._textboxText; }
        }

        /// <summary>
        /// 获取文档页眉文本框内容
        /// </summary>
        public String HeaderTextboxText
        {
            get { return this._headerTextboxText; }
        }
        #endregion

        #region 构造函数
        /// <summary>
        /// 初始化DocFile
        /// </summary>
        /// <param name="filePath">文件路径</param>
        public WordBinaryFile(String filePath) :
            base(filePath) { }
        #endregion

        #region 读取内容
        protected override void ReadContent()
        {
            this.ReadWordDocument();
            this.ReadTableStream();
            this.ReadPieceText();
        }
        #endregion

        #region 读取WordDocument
        private void ReadWordDocument()
        {
            DirectoryEntry entry = this._dirRootEntry.GetChild("WordDocument");

            if (entry == null)
            {
                return;
            }

            Int64 entryStart = this.GetEntryOffset(entry);
            this._stream.Seek(entryStart, SeekOrigin.Begin);

            this.ReadFileInformationBlock();
        }

        #region 读取FileInformationBlock
        private void ReadFileInformationBlock()
        {
            this.ReadFibBase();
            this.ReadFibRgW97();
            this.ReadFibRgLw97();
            this.ReadFibRgFcLcb();
            this.ReadFibRgCswNew();
        }

        #region FibBase
        private void ReadFibBase()
        {
            UInt16 wIdent = this._reader.ReadUInt16();
            if (wIdent != 0xA5EC)
            {
                throw new Exception("This file is not \".doc\" file.");
            }

            this._nFib = this._reader.ReadUInt16();
            this._reader.ReadUInt16();//unused
            this._reader.ReadUInt16();//lid
            this._reader.ReadUInt16();//pnNext

            UInt16 flags = this._reader.ReadUInt16();
            this._isComplexFile = BitHelper.GetBitFromInteger(flags, 2);
            this._hasPictures = BitHelper.GetBitFromInteger(flags, 3);
            this._isEncrypted = BitHelper.GetBitFromInteger(flags, 8);
            this._is1Table = BitHelper.GetBitFromInteger(flags, 9);

            if (this._isComplexFile)
            {
                throw new Exception("Do not support the complex file.");
            }

            if (this._isEncrypted)
            {
                throw new Exception("Do not support the encvypted file.");
            }

            this._stream.Seek(32 - 12, SeekOrigin.Current);
        }
        #endregion

        #region FibRgW97
        private void ReadFibRgW97()
        {
            UInt16 count = this._reader.ReadUInt16();

            if (count != 0x000E)
            {
                throw new Exception("File has been broken (FibRgW97 length is INVALID).");
            }

            this._stream.Seek(26, SeekOrigin.Current);
            this._lidFE = this._reader.ReadUInt16();
        }
        #endregion

        #region FibRgLw97
        private void ReadFibRgLw97()
        {
            UInt16 count = this._reader.ReadUInt16();

            if (count != 0x0016)
            {
                throw new Exception("File has been broken (FibRgLw97 length is INVALID).");
            }

            this._cbMac = this._reader.ReadInt32();
            this._reader.ReadInt32();//reserved1
            this._reader.ReadInt32();//reserved2
            this._ccpText = this._reader.ReadInt32();
            this._ccpFtn = this._reader.ReadInt32();
            this._ccpHdd = this._reader.ReadInt32();
            this._reader.ReadInt32();//reserved3
            this._ccpAtn = this._reader.ReadInt32();
            this._ccpEdn = this._reader.ReadInt32();
            this._ccpTxbx = this._reader.ReadInt32();
            this._ccpHdrTxbx = this._reader.ReadInt32();

            this._stream.Seek(44, SeekOrigin.Current);
        }
        #endregion

        #region FibRgFcLcb
        private void ReadFibRgFcLcb()
        {
            UInt16 count = this._reader.ReadUInt16();
            this._stream.Seek(66 * 4, SeekOrigin.Current);

            this._fcClx = this._reader.ReadUInt32();
            this._lcbClx = this._reader.ReadUInt32();

            this._stream.Seek((count * 2 - 68) * 4, SeekOrigin.Current);
        }
        #endregion

        #region FibRgCswNew
        private void ReadFibRgCswNew()
        {
            UInt16 count = this._reader.ReadUInt16();
            this._nFib = this._reader.ReadUInt16();
            this._stream.Seek((count - 1) * 2, SeekOrigin.Current);
        }
        #endregion
        #endregion
        #endregion

        #region 读取TableStream
        private void ReadTableStream()
        {
            DirectoryEntry entry = this._dirRootEntry.GetChild((this._is1Table ? "1Table" : "0Table"));

            if (entry == null)
            {
                return;
            }

            Int64 pieceTableStart = this.GetEntryOffset(entry) + this._fcClx;
            Int64 pieceTableEnd = pieceTableStart + this._lcbClx;
            this._stream.Seek(pieceTableStart, SeekOrigin.Begin);

            Byte clxt = this._reader.ReadByte();
            Int32 prcLen = 0;

            //判断如果是Prc不是Pcdt
            while (clxt == 0x01 && this._stream.Position < pieceTableEnd)
            {
                this._stream.Seek(prcLen, SeekOrigin.Current);
                clxt = this._reader.ReadByte();
                prcLen = this._reader.ReadInt32();
            }

            if (clxt != 0x02)
            {
                throw new Exception("There's no content in this file.");
            }

            UInt32 size = this._reader.ReadUInt32();
            UInt32 count = (size - 4) / 12;

            this._lstPieceStartPosition = new List<UInt32>();
            this._lstPieceEndPosition = new List<UInt32>();
            this._lstPieceElement = new List<PieceElement>();

            for (Int32 i = 0; i < count; i++)
            {
                this._lstPieceStartPosition.Add(this._reader.ReadUInt32());
                this._lstPieceEndPosition.Add(this._reader.ReadUInt32());
                this._stream.Seek(-4, SeekOrigin.Current);
            }

            this._stream.Seek(4, SeekOrigin.Current);

            for (Int32 i = 0; i < count; i++)
            {
                UInt16 info = this._reader.ReadUInt16();
                UInt32 fcCompressed = this._reader.ReadUInt32();
                UInt16 prm = this._reader.ReadUInt16();

                this._lstPieceElement.Add(new PieceElement(info, fcCompressed, prm));
            }
        }
        #endregion

        #region 读取文本内容
        private void ReadPieceText()
        {
            StringBuilder sb = new StringBuilder();
            DirectoryEntry entry = this._dirRootEntry.GetChild("WordDocument");

            for (Int32 i = 0; i < this._lstPieceElement.Count; i++)
            {
                Int64 pieceStart = this.GetEntryOffset(entry) + this._lstPieceElement[i].Offset;
                this._stream.Seek(pieceStart, SeekOrigin.Begin);

                Int32 length = (Int32)((this._lstPieceElement[i].IsUnicode ? 2 : 1) * (this._lstPieceEndPosition[i] - this._lstPieceStartPosition[i]));
                Byte[] data = this._reader.ReadBytes(length);
                String content = StringHelper.GetString(this._lstPieceElement[i].IsUnicode, data);
                sb.Append(content);
            }

            String allContent = sb.ToString();
            Int32 paragraphEnd = this._ccpText;
            Int32 footnoteEnd = paragraphEnd + this._ccpFtn;
            Int32 headerEnd = footnoteEnd + this._ccpHdd;
            Int32 commentEnd = headerEnd + this._ccpAtn;
            Int32 endnoteEnd = commentEnd + this._ccpEdn;
            Int32 textboxEnd = endnoteEnd + this._ccpTxbx;
            Int32 headerTextboxEnd = textboxEnd + this._ccpHdrTxbx;

            this._paragraphText = StringHelper.ReplaceString(allContent.Substring(0, this._ccpText));
            this._footnoteText = StringHelper.ReplaceString(allContent.Substring(paragraphEnd, this._ccpFtn));
            this._headerText = StringHelper.ReplaceString(allContent.Substring(footnoteEnd, this._ccpHdd));
            this._commentText = StringHelper.ReplaceString(allContent.Substring(headerEnd, this._ccpAtn));
            this._endnoteText = StringHelper.ReplaceString(allContent.Substring(commentEnd, this._ccpEdn));
            this._textboxText = StringHelper.ReplaceString(allContent.Substring(endnoteEnd, this._ccpTxbx));
            this._headerTextboxText = StringHelper.ReplaceString(allContent.Substring(textboxEnd, this._ccpHdrTxbx));
        }
        #endregion
    }
}