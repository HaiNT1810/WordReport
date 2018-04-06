using Microsoft.SharePoint;
using Syncfusion.DocIO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace MSWordReport
{
    public partial class MSWordReport : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            
        }

        protected void Export_Click(object sender, EventArgs e)
        {
            try
            {
                string fileNameOutput = "Download.docx";
                string fileUrl = @"C:/Users/ThanhHai/Desktop/TestReport/OpenReport/BieuMau.docx";
                Word Word = new Word(fileUrl, fileNameOutput);
                //Get base 64 to image
                string path = @"C:/Users/ThanhHai/Desktop/TestReport/OpenReport/barcode.jpg";
                byte[] bytes = System.IO.File.ReadAllBytes(path);
                string baseString = System.Convert.ToBase64String(bytes);

                string path1 = @"C:/Users/ThanhHai/Desktop/TestReport/OpenReport/barcode1.jpg";
                byte[] bytes1 = System.IO.File.ReadAllBytes(path1);
                string baseString1 = System.Convert.ToBase64String(bytes1);
                Word.SetTag("AnhDaiDien", baseString);
                for (int i = 0; i < 10; i++)
                {
                    Word.SetRepeat(new GroupStudent(i + 1, "Nhom so 1"));
                    for (int j = 0; j < 3; j++)
                    {
                        Word.SetRepeat(new InfoStudent("2", "Nhom02", "Nguyễn Thanh Hải1211", "1213123", 8, ""));
                        Word.SetRepeat(new InfoStudent("1", "Nhom01", "Nguyễn Thanh Hải", "1111", 8, baseString));
                    }
                    //Word.SetRepeat(new InfoStudent() { HoTen = "Le Linh", STT = 2, Anh = "abc" });
                    //Word.SetRepeat(new InfoStudent(3, "Nhom01", "Đào Đức Trình", 3));
                    //Word.SetRepeat(new InfoStudent(4, "Nhom01", "Đỗ Ninh Tất Điệp", 4));
                    //Word.SetRepeat(new InfoStudent(5, "Nhom01", , 5));
                    Word.SetRepeat(new SumStudent(i));
                    Word.SetRepeat(new InfoCLass("Lop 10", 5.2));
                }
                Word.SetTag("HoTen", "Nguyen van a");
                Word.SetTag("GioiTinh", "Nam");
                Word.SetTag("Test", DateTime.Now.ToString("dd/MM/yyyy"));
                Word.DownloadReport();
            }
            catch (Exception exx)
            {
                lbMessage.Text = exx.ToString();
            }
        }

        public static Stream GetStreamToSPFile(string folderUrl, string fileName)
        {
            SPWeb oWeb = SPContext.Current.Web;
            SPFolder folder = oWeb.GetFolder(folderUrl);
            SPFile file = folder.Files[fileName];
            return file.OpenBinaryStream();
        }
    }
}