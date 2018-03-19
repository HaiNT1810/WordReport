using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.IO;
using System.Reflection;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Drawing;
using System.Net;
using OpenReport.WordCommon;

namespace OpenReport
{
    public class Word
    {
        #region Contractor
        private string fileNameOutput;
        private List<ITemplate> lstTemplate = new List<ITemplate>();
        private List<ITable> lstTable = new List<ITable>();
        private List<IRepeat> lstRepeat = new List<IRepeat>();
        private List<ITag> lstTag = new List<ITag>();
        private int[] countRowTable;
        public WordDocument Document { get; set; }
        public HttpResponse Response { get; set; }
        public WSection Section { get; set; }
        public Word(Stream fileInput, string fileNameOutput)
        {
            Response = HttpContext.Current.Response;
            this.fileNameOutput = fileNameOutput;
            Document = new WordDocument(fileInput, FormatType.Docx);
            Section = Document.LastSection;
            GetAllTemplate();
        }
        public Word(string fileUrl, string fileNameOutput)
        {
            Response = HttpContext.Current.Response;
            this.fileNameOutput = fileNameOutput;
            MemoryStream stream = Common.GetFileStreamFromUrl(fileUrl);
            Document = new WordDocument(stream, FormatType.Docx);

            Section = Document.LastSection;
            GetAllTemplate();
        }
        #endregion

        #region Publish Function
        public void SetRepeat(object obj)
        {
            if (GetTableOfIT(obj, lstTemplate) != null)
                lstRepeat.Add(GetTableOfIT(obj, lstTemplate));
        }
        public void SetTag(string tagName, object obj)
        {
            if (tagName != "" && tagName != null)
            {
                tagName = tagName.Trim();
                if (tagName.IndexOf('{') == -1)
                {
                    tagName = "{" + tagName;
                }
                if (tagName.IndexOf('}') == -1)
                {
                    tagName = tagName + "}";
                }
                lstTag.Add(new ITag(tagName, obj, "Text"));
            }
        }
        public void SetTag(string tagName, object obj, TagWordType type = TagWordType.Text)
        {
            if (tagName != "" && tagName != null)
            {
                tagName = tagName.Trim();
                if (tagName.IndexOf('{') == -1)
                {
                    tagName = "{" + tagName;
                }
                if (tagName.IndexOf('}') == -1)
                {
                    tagName = tagName + "}";
                }
                lstTag.Add(new ITag(tagName, obj, type.ToString()));
            }
        }
        public void DownloadReport(FormatType type = FormatType.Docx)
        {
            try
            {
                SetRepeatValue();
                SetTagValue();
                DeleteItemplate();
                Document.Save(fileNameOutput, type, Response, HttpContentDisposition.Attachment);
            }
            catch (Exception)
            {
                Document.Save(fileNameOutput, type, Response, HttpContentDisposition.Attachment);
            }
        }
        #endregion

        #region Private Function
        private void SetRepeatValue()
        {
            if (lstRepeat.Count > 0)
            {
                foreach (IRepeat iRepeat in lstRepeat)
                {
                    WTableRow r = Section.Tables[iRepeat.ITemp.IndexTable].Rows[iRepeat.ITemp.IndexRow].Clone();
                    Section.Tables[iRepeat.ITemp.IndexTable].Rows.Insert(countRowTable[iRepeat.ITemp.IndexTable], r);
                    countRowTable[iRepeat.ITemp.IndexTable] = countRowTable[iRepeat.ITemp.IndexTable] + 1;
                    string paragraph = string.Empty;
                    WTableCell cell;
                    for (int i = 0, n = Section.Tables[iRepeat.ITemp.IndexTable].LastRow.Cells.Count; i < n; i++)
                    {
                        cell = Section.Tables[iRepeat.ITemp.IndexTable].LastRow.Cells[i];
                        paragraph = ReplateTempInParagraph(cell.LastParagraph.Text, iRepeat);
                        cell.LastParagraph.Text = paragraph;
                    }
                }
            }
        }
        private void SetTagValue()
        {
            if (lstTag.Count > 0)
            {
                foreach (ITag iTag in lstTag)
                {
                    switch (iTag.TagType.ToLower())
                    {
                        case "text":
                            Document.Replace(iTag.TagName, iTag.Data.ToString(), false, false);
                            break;
                        case "image":
                            Image img = ConvertStrBase64ToImage(iTag.Data.ToString());
                            IWParagraph paragraph = Section.AddParagraph();
                            IWPicture picture = paragraph.AppendPicture(img);
                            picture.Height = 100; picture.Width = 200;
                            paragraph.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                            TextBodyPart textBodyPart = new TextBodyPart(Document);
                            textBodyPart.BodyItems.Add(paragraph);
                            Document.Replace(iTag.TagName, textBodyPart, false, false);
                            break;
                        default:
                            break;
                    }

                }
            }
        }
        /// <summary>
        /// Get all template in a table of form
        /// </summary>
        private void GetAllTemplate()
        {
            List<int> CountRowTable = new List<int>();
            for (int i = 0, n = Section.Tables.Count; i < n; i++)
            {
                int indexRow = 0;
                lstTable.Add(new ITable(i, Section.Tables[i].Rows.Count));
                for (int j = 0, m = Section.Tables[i].Rows.Count; j < m; j++)
                {
                    indexRow = j;
                    List<ITempTag> tmpcell = new List<ITempTag>();
                    for (int k = 0, l = Section.Tables[i].Rows[j].Cells.Count; k < l; k++)
                    {
                        GetArrTemp(Section.Tables[i].Rows[j].Cells[k], tmpcell);
                    }
                    if (tmpcell.Count != 0)
                        lstTemplate.Add(new ITemplate(i, indexRow, tmpcell));
                }
                CountRowTable.Add(Section.Tables[i].Rows.Count);
            }
            countRowTable = CountRowTable.ToArray();
        }
        /// <summary>
        /// get array template table in form word
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="tmpcell"></param>
        /// <returns></returns>
        private List<ITempTag> GetArrTemp(WTableCell cell, List<ITempTag> tmpcell)
        {
            string paragraph = cell.LastParagraph.Text;
            if (paragraph.IndexOfAny(new char[] { '{', '}' }) != -1)
            {
                string[] lstTemp = paragraph.Split('{');
                for (int i = 0, n = lstTemp.Length; i < n; i++)
                {
                    if (lstTemp[i].IndexOf('}') != -1)
                    {
                        bool flag = false;
                        string str = lstTemp[i].Split('}')[0];
                        if (str != "" && str != null)
                        {
                            if (str.IndexOf(":") == -1)
                            {
                                foreach (ITempTag itt in tmpcell)
                                {
                                    if (itt.TagName == str)
                                    {
                                        flag = true;
                                    }
                                }
                                if (flag == false)
                                    tmpcell.Add(new ITempTag(str, TagWordType.Text.ToString()));
                            }
                            else
                            {
                                string tagName = str.Split(':')[1];
                                string tagType = str.Split(':')[0];
                                foreach (ITempTag itt in tmpcell)
                                {
                                    if (itt.TagName == tagName)
                                    {
                                        flag = true;
                                    }
                                }
                                if (flag == false)
                                    tmpcell.Add(new ITempTag(tagName, tagType));
                            }
                        }
                    }
                }
            }
            return tmpcell;
        }
        /// <summary>
        /// get template of a row in a table by Itemplate
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="Itemp"></param>
        /// <returns></returns>
        private IRepeat GetTableOfIT(object obj, List<ITemplate> Itemp)
        {
            PropertyInfo[] pptI;
            List<string> str = new List<string>();
            pptI = obj.GetType().GetProperties();
            int length = pptI.Count();
            for (int i = 0; i < length; i++)
            {
                str.Add(pptI[i].Name);
            }
            foreach (ITemplate itemp in Itemp)
            {
                List<string> temp = new List<string>();
                for (int j = 0, m = itemp.ITempTag.Count; j < m; j++)
                {
                    temp.Add(itemp.ITempTag[j].TagName);
                }
                bool flag = CompareArrays(str.ToArray(), temp.ToArray());
                if (flag == true)
                {
                    return new IRepeat(itemp, obj);
                }
            }
            return null;
        }
        private bool CompareArrays(string[] a, string[] b)
        {
            if (a.Length != b.Length) { return false; }
            for (int i = 0; i < a.Length; i++)
            {
                if (a[i] != b[i]) { return false; }

            }
            return true;
        }
        /// <summary>
        /// Replate all tags in a paragraph of a cell
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="obj"></param>
        /// <returns></returns>
        private string ReplateTempInParagraph(string paragraph, IRepeat IRepeat)
        {
            PropertyInfo[] pptI;
            List<string> str = new List<string>();
            pptI = IRepeat.Data.GetType().GetProperties();
            int length = pptI.Count();
            for (int i = 0; i < length; i++)
            {
                str.Add("{" + pptI[i].Name + "}");
            }
            string[] lst = str.ToArray();
            for (int i = 0; i < lst.Length; i++)
            {
                if (paragraph.IndexOf(lst[i]) != -1)
                {
                    paragraph = paragraph.Replace(lst[i], pptI[i].GetValue(IRepeat.Data, null).ToString());
                }
            }
            return paragraph;
        }
        /// <summary>
        /// Delete All Itemplate in form
        /// </summary>
        private void DeleteItemplate()
        {
            foreach (ITable table in lstTable)
            {
                int countR = table.CountRow - 1;
                for (int i = countR; i > 0; i--)
                {
                    foreach (ITemplate iTemp in lstTemplate)
                    {
                        if (iTemp.IndexTable == table.IndexTable && iTemp.IndexRow == i)
                        {
                            Section.Tables[table.IndexTable].Rows.RemoveAt(i);
                            break;
                        }
                    }
                }
            }
        }
        private Image ConvertStrBase64ToImage(string base64String)
        {
            byte[] buffer = Convert.FromBase64String(base64String);

            if (buffer != null)
            {
                ImageConverter ic = new ImageConverter();
                return ic.ConvertFrom(buffer) as Image;
            }
            else
                return null;
        }
        #endregion
    }
    public enum TagWordType
    {
        Text,
        Image
    }
}