using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSWordReport
{
    public class Contains
    {
        public const string START_TAG = "{";
        public const string END_TAG = "}";
        public const string DATETIME_FORMAT = "dd/MM/yyyy";
        public const string GROUP_NODE = "Group:";
        public const char GROUP_CHARACTER = ':';
        public const char EQUAL = '=';
    }
    public class ITemplateWord
    {
        public int IndexTable { set; get; }
        public int IndexRow { set; get; }
        public List<ITempTagWord> ITempTag { get; set; }
        public ITemplateWord(int indexTable, int indexRow, List<ITempTagWord> iTempTag)
        {
            this.IndexTable = indexTable;
            this.IndexRow = indexRow;
            this.ITempTag = iTempTag;
        }
    }
    public class ITempTagWord
    {
        public string TagName { get; set; }
        public string TagType { get; set; }
        public string TagStyle { get; set; }

        public ITempTagWord(string tagName, string tagType, string tagStyle)
        {
            this.TagName = tagName;
            this.TagType = tagType;
            this.TagStyle = tagStyle;
        }
    }
    public class ITableWord
    {
        public int IndexTable { get; set; }
        public int CountRow { get; set; }
        public ITableWord(int indexTable, int countRow)
        {
            this.IndexTable = indexTable;
            this.CountRow = countRow;
        }
    }
    public class IRepeatWord
    {
        public ITemplateWord ITemp { get; set; }
        public object Data { get; set; }
        public IRepeatWord(ITemplateWord iTemp, object data)
        {
            this.ITemp = iTemp;
            this.Data = data;
        }
    }
    public class ITagWord
    {
        public string TagName { get; set; }
        public object Data { get; set; }
        public TagWordType TagType { get; set; }
        public string TagStyle { get; set; }
        public ITagWord(string tagName, object data, TagWordType tagType,string tagStyle)
        {
            this.TagName = tagName;
            this.Data = data;
            this.TagType = tagType;
            this.TagStyle = tagStyle;
        }

        public ITagWord()
        {
            // TODO: Complete member initialization
        }
    }
    #region Enum
    public enum WordFileType
    {
        Doc,
        Dot,
        Docx,
        Word2007,
        Word2010,
        Word2013,
        Word2007Dotx,
        Word2010Dotx,
        Word2013Dotx,
        Word2007Docm,
        Word2010Docm,
        Word2013Docm,
        Word2007Dotm,
        Word2010Dotm,
        Word2013Dotm,
        WordML
    }


    #endregion
}