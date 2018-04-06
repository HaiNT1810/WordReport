using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSWordReport
{
    public class Data
    {
    }
    public class GroupStudent
    {
        public int STT { get; set; }
        public string TenNhom { get; set; }
        public GroupStudent(int stt, string tenNhom)
        {
            this.STT = stt;
            this.TenNhom = tenNhom;
        }
    }
    public class InfoStudent
    {
        public string TenLuatPhapLenh { get; set; }
        public string ChuongTrinh { get; set; }
        public string VBTrangThai { get; set; }
        public string GhiChuTinhTrang { get; set; }
        public int Diem { get; set; }
        public string Anh { get; set; }
        public InfoStudent()
        { }
        public InfoStudent(string tenLuatPhapLenh, string chuongTrinh, string vBTrangThai, string ghiChuTinhTrang, int diem, string anh)
        {
            this.TenLuatPhapLenh = tenLuatPhapLenh;
            this.ChuongTrinh = chuongTrinh;
            this.VBTrangThai = vBTrangThai;
            this.GhiChuTinhTrang = ghiChuTinhTrang;
            this.Diem = diem;
            this.Anh = anh;
        }


    }
    public class SumStudent
    {
        public int Tong { get; set; }
        public SumStudent(int tong)
        {
            this.Tong = tong;
        }
    }
    public class InfoCLass
    {
        public string Lop { get; set; }
        public double DiemTB { get; set; }
        public InfoCLass(string lop, double diemTB)
        {
            this.Lop = lop;
            this.DiemTB = diemTB;
        }
    }
    public class Truong
    {
        public string TenTruong { get; set; }
        public string AnhTruong { get; set; }
        public Truong(string tenTruong, string anhTruong)
        {
            this.TenTruong = tenTruong;
            this.AnhTruong = anhTruong;
        }
    }
    public class NhomDuAn
    {
        int STTNhom { get; set; }
        public string DiaBan { get; set; }
        public NhomDuAn(int sttNhom, string diaBan)
        {
            this.STTNhom = sttNhom;
            this.DiaBan = diaBan;
        }
    }
    public class DuAn
    {
        public int STT { get; set; }
        public string TieuDe { get; set; }
        public string ChuTriDuAn { get; set; }
        public string PhuTrachSanPham { get; set; }
        public string PhuTrachTrienKhai { get; set; }
        public DuAn(int stt, string tieuDe, string chuTriDuAn, string phuTrachSanPham, string phuTrachTrienKhai)
        {
            this.STT = stt;
            this.TieuDe = tieuDe;
            this.ChuTriDuAn = chuTriDuAn;
            this.PhuTrachSanPham = phuTrachSanPham;
            this.PhuTrachTrienKhai = phuTrachTrienKhai;
        }
    }
}