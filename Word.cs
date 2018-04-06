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

namespace MSWordReport
{
    public class Word
    {
        #region Contractor
        private string fileNameOutput;
        private List<ITemplateWord> lstTemplate = new List<ITemplateWord>();
        private List<ITableWord> lstTable = new List<ITableWord>();
        private List<IRepeatWord> lstRepeat = new List<IRepeatWord>();
        private List<ITagWord> lstTag = new List<ITagWord>();
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
            Document = new WordDocument(fileUrl, FormatType.Docx);
            Section = Document.LastSection;
            GetAllTemplate();
            //Response = HttpContext.Current.Response;
            //this.fileNameOutput = fileNameOutput;
            //MemoryStream stream = GetFileStreamFromUrl(fileUrl);
            //Document = new WordDocument(stream, FormatType.Docx);
            //Section = Document.LastSection;
            //GetAllTemplate();

        }
        #endregion
        /// <summary>
        /// replate data to template clone
        /// </summary>
        /// <param name="obj"></param>
        /// 
        #region Public Function
        public void SetRepeat(object obj)
        {
            IRepeatWord repeat = GetTableOfIT(obj, lstTemplate);
            if (repeat != null)
                lstRepeat.Add(repeat);
        }
        public void SetTag(string tagName, object obj)
        {
            if (tagName != "" && tagName != null)
            {
                tagName = tagName.Trim();
                GetPropertyOfTag(tagName, obj);
                //GetStype(tagName, obj, TagWordType.Text);
            }
        }
        public void SetTag(string tagName, object obj, TagWordType type = TagWordType.Text)
        {
            if (tagName != "" && tagName != null)
            {
                tagName = tagName.Trim();
                GetPropertyOfTag(tagName, obj);
                //GetStype(tagName, obj, type);
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
                foreach (IRepeatWord iRepeat in lstRepeat)
                {
                    WTableRow r = Section.Tables[iRepeat.ITemp.IndexTable].Rows[iRepeat.ITemp.IndexRow].Clone();
                    Section.Tables[iRepeat.ITemp.IndexTable].Rows.Insert(countRowTable[iRepeat.ITemp.IndexTable], r);
                    countRowTable[iRepeat.ITemp.IndexTable] = countRowTable[iRepeat.ITemp.IndexTable] + 1;
                    string paragraph = string.Empty;
                    WTableCell cell;
                    for (int i = 0, n = Section.Tables[iRepeat.ITemp.IndexTable].LastRow.Cells.Count; i < n; i++)
                    {
                        cell = Section.Tables[iRepeat.ITemp.IndexTable].LastRow.Cells[i];
                        ReplateTempInCell(cell, iRepeat);
                        //paragraph = ReplateTempInParagraph(cell.LastParagraph.Text, iRepeat);
                        //cell.LastParagraph.Text = paragraph;
                    }
                }
            }
        }
        private void SetTagValue()
        {
            if (lstTag.Count > 0)
            {
                foreach (ITagWord iTag in lstTag)
                {
                    string tagName = Contains.START_TAG + iTag.TagName + Contains.END_TAG;
                    switch (iTag.TagType)
                    {
                        case TagWordType.Text:
                            Document.Replace(tagName, iTag.Data.ToString(), false, false);
                            break;
                        case TagWordType.Image:
                            Image img = ConvertStrBase64ToImage(iTag.Data.ToString());
                            IWParagraph paragraph = Section.AddParagraph();
                            IWPicture picture = paragraph.AppendPicture(img);
                            if (iTag.TagStyle != "")
                            {
                                string[] lstStyle = iTag.TagStyle.Split(';');
                                for (int i = 0, n = lstStyle.Length; i < n; i++)
                                    SetStyle(picture, lstStyle[i]);
                            }
                            paragraph.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                            TextBodyPart textBodyPart = new TextBodyPart(Document);
                            textBodyPart.BodyItems.Add(paragraph);
                            Document.Replace(tagName, textBodyPart, false, false);
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
                lstTable.Add(new ITableWord(i, Section.Tables[i].Rows.Count));
                for (int j = 0, m = Section.Tables[i].Rows.Count; j < m; j++)
                {
                    indexRow = j;
                    List<ITempTagWord> tmpcell = new List<ITempTagWord>();
                    for (int k = 0, l = Section.Tables[i].Rows[j].Cells.Count; k < l; k++)
                    {
                        GetArrTemp(Section.Tables[i].Rows[j].Cells[k], tmpcell);
                    }
                    if (tmpcell.Count != 0)
                        lstTemplate.Add(new ITemplateWord(i, indexRow, tmpcell));
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
        private List<ITempTagWord> GetArrTemp(WTableCell cell, List<ITempTagWord> tmpcell)
        {
            for (int i = 0, n = cell.Paragraphs.Count; i < n; i++)
            {
                string paragraph = cell.Paragraphs[i].Text;
                if (paragraph.IndexOfAny(new char[] { '{', '}' }) != -1)
                {
                    string[] lstTemp = paragraph.Split('{');
                    for (int j = 0, m = lstTemp.Length; j < m; j++)
                    {
                        if (lstTemp[j].IndexOf('}') != -1)
                        {
                            bool flag = false;
                            string str = lstTemp[j].Split('}')[0];
                            if (str != "" && str != null)
                            {
                                if (str.IndexOf(":") == -1)
                                {
                                    foreach (ITempTagWord itt in tmpcell)
                                    {
                                        if (itt.TagName == str)
                                        {
                                            flag = true;
                                        }
                                    }
                                    if (!flag)
                                        tmpcell.Add(new ITempTagWord(str, TagWordType.Text.ToString(), ""));
                                }
                                else
                                {
                                    string[] arrProperties = str.Split(':');
                                    string tagStyle = arrProperties[2] == null ? "" : arrProperties[2];
                                    string tagName = arrProperties[1];
                                    string tagType = arrProperties[0];
                                    foreach (ITempTagWord itt in tmpcell)
                                    {
                                        if (itt.TagName == tagName)
                                        {
                                            flag = true;
                                        }
                                    }
                                    if (!flag)
                                        tmpcell.Add(new ITempTagWord(tagName, tagType, tagStyle));
                                    paragraph = paragraph.Replace(str, tagName);
                                }
                            }
                        }
                    }
                    cell.Paragraphs[i].Text = paragraph;
                }
            }
            return tmpcell;
        }
        /// <summary>
        /// get style of a image tag
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        private void GetStype(string tagName, object obj, TagWordType type)
        {
            string str = obj.ToString();
            if (str.IndexOf(':') == -1)
                lstTag.Add(new ITagWord(tagName, obj, type, ""));
            else
            {
                string[] lstStr = str.Split(':');
                if (lstStr.Length == 2)
                {
                    lstTag.Add(new ITagWord(tagName, lstStr[0], type, lstStr[1]));
                }
            }
        }
        private void GetPropertyOfTag(string tagName, object obj)
        {
            string temp = Contains.START_TAG + TagWordType.Image.ToString() + Contains.GROUP_CHARACTER + tagName;
            if (Document.Find(Contains.START_TAG + tagName + Contains.END_TAG, false, true) != null)
                lstTag.Add(new ITagWord(tagName, obj, TagWordType.Text, ""));
            if (Document.Find(Contains.START_TAG + TagWordType.Image.ToString() + tagName + Contains.END_TAG, false, true) != null)
                lstTag.Add(new ITagWord(tagName, obj, TagWordType.Image, ""));
            if (Document.Find(temp, false, true) != null)
            {
                //TextSelection textSelection = Document.Find(Contains.START_TAG + TagWordType.Image.ToString() + tagName + Contains.GROUP_CHARACTER, false, false);
                string str = "";
                for (int a = 0, b = Document.Sections.Count; a < b; a++)
                {
                    for (int i = 0, n = Document.Sections[a].Paragraphs.Count; i < n; i++)
                    {
                        if (Document.Sections[a].Paragraphs[i].Text.IndexOf(temp) != -1)
                        {
                            str += Document.Sections[a].Paragraphs[i].Text;
                            int index = str.IndexOf(temp);
                            int endTag = str.Substring(index + temp.Count()).IndexOf('}');
                            string tag = str.Substring(index, endTag + index + temp.Count() + 1);
                            string style = tag.Substring(temp.Count() + 1).Replace("}", string.Empty);
                            Document.Replace(tag, Contains.START_TAG + tagName + Contains.END_TAG, false, false);
                            lstTag.Add(new ITagWord(tagName, obj, TagWordType.Image, style));
                        }
                    }
                }
            }
        }
        /// <summary>
        /// get template of a row in a table by Itemplate
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="Itemp"></param>
        /// <returns></returns>
        private IRepeatWord GetTableOfIT(object obj, List<ITemplateWord> Itemp)
        {
            PropertyInfo[] pptI;
            List<string> str = new List<string>();
            pptI = obj.GetType().GetProperties();
            int length = pptI.Count();
            for (int i = 0; i < length; i++)
            {
                str.Add(pptI[i].Name);
            }
            foreach (ITemplateWord itemp in Itemp)
            {
                List<string> temp = new List<string>();
                for (int j = 0, m = itemp.ITempTag.Count; j < m; j++)
                {
                    temp.Add(itemp.ITempTag[j].TagName);
                }
                bool flag = CompareArrays(str.ToArray(), temp.ToArray());
                if (flag)
                {
                    return new IRepeatWord(itemp, obj);
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
        private string ReplateTempInParagraph(string paragraph, IRepeatWord iRepeat)
        {
            PropertyInfo[] pptI;
            List<string> str = new List<string>();
            pptI = iRepeat.Data.GetType().GetProperties();
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

                    paragraph = paragraph.Replace(lst[i], pptI[i].GetValue(iRepeat.Data, null).ToString());
                }
            }
            return paragraph;
        }
        private void ReplateTempInCell(WTableCell cell, IRepeatWord iRepeat)
        {
            PropertyInfo[] pptI;
            pptI = iRepeat.Data.GetType().GetProperties();
            int length = pptI.Count();
            for (int i = 0, n = cell.Paragraphs.Count; i < n; i++)
            {
                string paragraph = cell.Paragraphs[i].Text;
                List<string> str = new List<string>();
                for (int j = 0; j < length; j++)
                {
                    str.Add("{" + pptI[j].Name + "}");
                }
                string[] lst = str.ToArray();
                for (int j = 0; j < lst.Length; j++)
                {
                    if (paragraph.IndexOf(lst[j]) != -1)
                    {
                        string tagType = "Text";
                        string tagStyle = string.Empty;
                        for (int k = 0, l = iRepeat.ITemp.ITempTag.Count; k < l; k++)
                        {
                            if (pptI[j].Name == iRepeat.ITemp.ITempTag[k].TagName)
                            {
                                tagType = iRepeat.ITemp.ITempTag[k].TagType;
                                tagStyle = iRepeat.ITemp.ITempTag[k].TagStyle;
                            }
                        }
                        if (tagType.ToLower() == TagWordType.Text.ToString().ToLower())
                        {
                            paragraph = paragraph.Replace(lst[j], pptI[j].GetValue(iRepeat.Data, null).ToString());
                            cell.Paragraphs[i].Text = paragraph;
                        }
                        else
                        {
                            string strImg = pptI[j].GetValue(iRepeat.Data, null).ToString();
                            string s = cell.Paragraphs[i].Text;
                            string[] lsts = s.Split(new string[] { lst[j] }, StringSplitOptions.None);
                            cell.Paragraphs[i].Text = "";
                            cell.Paragraphs[i].AppendText(lsts[0]);
                            if (strImg == "" || strImg == null)
                            {
                                if (lsts.Length != 1)
                                    cell.Paragraphs[i].AppendText(lsts[1]);
                            }
                            else
                            {
                                Image img = ConvertStrBase64ToImage(strImg);
                                IWPicture picture = cell.Paragraphs[i].AppendPicture(img);
                                if (lsts.Length != 1)
                                    cell.Paragraphs[i].AppendText(lsts[1]);
                                if (tagStyle != "")
                                {
                                    string[] lstStyle = tagStyle.Split(';');
                                    for (int a = 0, b = lstStyle.Length; a < b; a++)
                                        SetStyle(picture, lstStyle[a]);
                                }
                            }
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Delete All Itemplate in form
        /// </summary>
        private void DeleteItemplate()
        {
            foreach (ITableWord table in lstTable)
            {
                int countR = table.CountRow - 1;
                for (int i = countR; i > 0; i--)
                {
                    foreach (ITemplateWord iTemp in lstTemplate)
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
        private void SetStyle(IWPicture picture, string style)
        {
            string[] lstStyle = style.Split('-');
            switch (lstStyle[0].ToLower())
            {
                case "w":
                    picture.Width = float.Parse(lstStyle[1]);
                    break;
                case "h":
                    picture.Height = float.Parse(lstStyle[1]);
                    break;
                default:
                    break;
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

        private Stream GetFileStreamFromUrl(string fileUrl)
        {
            try
            {
                byte[] imageData = null;

                using (var wc = new System.Net.WebClient())
                    imageData = wc.DownloadData(fileUrl);

                return new MemoryStream(imageData);
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
        #endregion
    public enum TagWordType
    {
        Text,
        Image
    }
}