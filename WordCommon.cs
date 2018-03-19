using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OpenReport.WordCommon
{
    public class ITemplate
    {
        public int IndexTable { set; get; }
        public int IndexRow { set; get; }
        public List<ITempTag> ITempTag { get; set; }
        public ITemplate(int indexTable, int indexRow, List<ITempTag> iTempTag)
        {
            this.IndexTable = indexTable;
            this.IndexRow = indexRow;
            this.ITempTag = iTempTag;
        }
    }
    public class ITempTag
    {
        public string TagName { get; set; }
        public string TagType { get; set; }

        public ITempTag(string tagName, string tagType)
        {
            this.TagName = tagName;
            this.TagType = tagType;
        }
    }
    public class ITable
    {
        public int IndexTable { get; set; }
        public int CountRow { get; set; }
        public ITable(int indexTable, int countRow)
        {
            this.IndexTable = indexTable;
            this.CountRow = countRow;
        }
    }
    public class IRepeat
    {
        public ITemplate ITemp { get; set; }
        public object Data { get; set; }
        public IRepeat(ITemplate iTemp, object data)
        {
            this.ITemp = iTemp;
            this.Data = data;
        }
    }
    public class ITag
    {
        public string TagName { get; set; }
        public object Data { get; set; }
        public string TagType { get; set; }
        public ITag(string tagName, object data, string tagType)
        {
            this.TagName = tagName;
            this.Data = data;
            this.TagType = tagType;
        }
    }
}