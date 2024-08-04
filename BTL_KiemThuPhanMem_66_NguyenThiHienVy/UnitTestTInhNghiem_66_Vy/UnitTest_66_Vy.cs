using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using BTL_KiemThuPhanMem_66_NguyenThiHienVy;
using CsvHelper;//Hỗ trợ đọc ghi dữ liệu file csv
using System.Collections.Generic;//Dùng để biểu diễn một danh sách động
using System.IO;//Được sử dụng để làm các thao tác nhập/xuất dữ liệu từ và đến tệp tin
using CsvHelper.Configuration;//Cấu hình cách thư viện CsvHelper đọc và hiểu dữ liệu từ tập tin CSV
using System.Globalization;//cung cấp các lớp và phương thức để làm việc với
                           //các vấn đề liên quan đến văn hóa, như định dạng số,
                           //ngày tháng, tiền tệ và ngôn ngữ.
using System.Linq;//Cung cấp các phương thức mở rộng và các công cụ hỗ trợ cho việc thao
                  //tác với các cấu trúc dữ liệu kiểu Enumerable trong .NET,
                  //chẳng hạn như danh sách (List), mảng (Array), và các cấu trúc dữ liệu tương tự.
using System.Text;

namespace UnitTestTInhNghiem_66_Vy
{
    [TestClass]
    public class UnitTest_66_Vy
    {
        public TestContext TestContext { get; set; }
        private TimNghiem_66_Vy c_66_Vy;

        //Tạo 1 đối tượng gồm 4 thuộc tính (tương ứng 4 cột trong file csv)
        public class DataRecord_66_Vy
        {
            public int a_66_Vy { get; set; }
            public int b_66_Vy { get; set; }
            public string expected_66_Vy { get; set; }
            public string result_66_Vy { get; set; }
        }
        //--------------------------Unit test phương thức 1 data không ghi kết quả----------------------------------------------------------
       // [TestMethod]
        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV", @".\Data_66_Vy\DataUnitTest_66_Vy.csv", "DataUnitTest_66_Vy#csv", DataAccessMethod.Sequential)]
        public void DataUnitTest_66_Vy()
        {
            int a_66_Vy, b_66_Vy;
            double expected_66_Vy, actual_66_Vy;
            a_66_Vy = int.Parse(TestContext.DataRow[0].ToString());
            b_66_Vy = int.Parse(TestContext.DataRow[1].ToString());
            expected_66_Vy = double.Parse(TestContext.DataRow[2].ToString());
            c_66_Vy = new TimNghiem_66_Vy(a_66_Vy, b_66_Vy);
            actual_66_Vy = c_66_Vy.TimNghiemPhuongTrinh_66_Vy();
            //Ném ngoại lệ kiểm tra pass hay fail
            Assert.AreEqual(expected_66_Vy, actual_66_Vy);
        }
        //--------------------------Unit test phương thức 1 data ghi kết quả----------------------------------------------------------------
        // [TestMethod]
        //[DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV", @".\Data_66_Vy\DataUnitTest_66_Vy.csv", "DataUnitTest_66_Vy#csv", DataAccessMethod.Sequential)]
        //public void DataUnitTest_66_Vy()
        //{
        //    int a_66_Vy, b_66_Vy;
        //    double expected_66_Vy = 0.00001, actual_66_Vy;
        //    a_66_Vy = int.Parse(TestContext.DataRow[0].ToString());
        //    b_66_Vy = int.Parse(TestContext.DataRow[1].ToString());
        //    c_66_Vy = new TimNghiem_66_Vy(a_66_Vy, b_66_Vy);
        //    actual_66_Vy = c_66_Vy.TimNghiemPhuongTrinh_66_Vy();
        //    var records_66_Vy = new List<DataRecord_66_Vy>();//Khai báo 1 danh sách có kiểu dữ liệu là DataRecord_66_Vy đã tạo ở trên
        //    string csvFilePath_66_Vy =
        //        @"D:\Myseft\Ki2_nam3\KiemThu\BTL_Copy\BTL_66_Vy\BTL_KiemThuPhanMem_66_NguyenThiHienVy\UnitTestTInhNghiem_66_Vy\Data_66_Vy\DataUnitTest_66_Vy.csv";
        //    //Đọc dữ liệu từ file csv
        //    using (var reader_66_Vy = new StreamReader(csvFilePath_66_Vy))//Mở tập tin CSV được chỉ định bởi biến csvFilePath để đọc nó
        //    using (var csv_66_Vy = new CsvReader(reader_66_Vy, new CsvConfiguration(CultureInfo.InvariantCulture)))
        //        records_66_Vy = csv_66_Vy.GetRecords<DataRecord_66_Vy>().ToList();//Đọc dữ liệu từ tập tin lưu bào biến records_66_Vy
        //    //Ghi pass hay fail vào cột result
        //    for (int i_66_Vy = 0; i_66_Vy < records_66_Vy.Count; i_66_Vy++)
        //    {   // Kiểm tra đúng hàng mới so sánh
        //        if (records_66_Vy[i_66_Vy].a_66_Vy == a_66_Vy && records_66_Vy[i_66_Vy].b_66_Vy == b_66_Vy)
        //        {//Gáng lại biến expected để ghi vào file csv, vì gặp trường hợp số thực thì định dạng số thực khác nhau giữa c# và csv
        //            expected_66_Vy = double.Parse(records_66_Vy[i_66_Vy].expected_66_Vy);
        //            //So sánh xem giá trị mong đợi và giá trị thực tế có giống nhau hay không
        //            if (double.Parse(records_66_Vy[i_66_Vy].expected_66_Vy) == actual_66_Vy)
        //                records_66_Vy[i_66_Vy].result_66_Vy = "pass";
        //            else
        //                records_66_Vy[i_66_Vy].result_66_Vy = "fail";
        //        }
        //    }
        //    using (var writer_66_Vy = new StreamWriter(csvFilePath_66_Vy, false))// Ghi dữ liệu vào tệp CSV
        //    using (var csvWriter_66_Vy = new CsvWriter(writer_66_Vy, System.Globalization.CultureInfo.InvariantCulture))
        //        csvWriter_66_Vy.WriteRecords(records_66_Vy);//Viết dữ liệu vào file csv
        //    //Ném ngoại lệ kiểm tra pass hay fail
        //    Assert.AreEqual(expected_66_Vy, actual_66_Vy);
        //}
        //--------------------------Unit test phương thức 1 từng test case--------------------------------------------------------
        //[TestMethod]
        //public void TinhNghiem_a_66_Vy()
        //{
        //    double expected_66_Vy, actual_66_Vy;
        //    c_66_Vy = new TimNghiem_66_Vy(2, -10);
        //    expected_66_Vy = 5;
        //    actual_66_Vy = c_66_Vy.TimNghiemPhuongTrinh_66_Vy();
        //    Assert.AreEqual(expected_66_Vy, actual_66_Vy);
        //}
        //[TestMethod]
        //public void TinhNghiem_bc_66_Vy()
        //{
        //    double expected_66_Vy, actual_66_Vy;
        //    c_66_Vy = new TimNghiem_66_Vy(0, 0);
        //    expected_66_Vy = 0.001;
        //    actual_66_Vy = c_66_Vy.TimNghiemPhuongTrinh_66_Vy();
        //    Assert.AreEqual(expected_66_Vy, actual_66_Vy);
        //}
        //[TestMethod]
        //public void TinhNghiem_bd_66_Vy()
        //{
        //    double expected_66_Vy, actual_66_Vy;
        //    c_66_Vy = new TimNghiem_66_Vy(0, 2);
        //    expected_66_Vy = -0.001;
        //    actual_66_Vy = c_66_Vy.TimNghiemPhuongTrinh_66_Vy();
        //    Assert.AreEqual(expected_66_Vy, actual_66_Vy);
        //}

        //--------------------------Unit test phương thức 2 từng test case----------------------------------------------------------
        //[TestMethod]
        //public void KiemTraDiemCoThuocPhuongTrinh_ac_66_Vy()
        //{
        //    c_66_Vy = new TimNghiem_66_Vy(0, 0);
        //    double x_66_Vy = 1;
        //    string expected_66_Vy = "Giá trị 1 thuộc phương trình 0x + 0 = 0";
        //    string actual_66_Vy = c_66_Vy.KiemTraDiemCoThuocPhuongTrinh_66_Vy(x_66_Vy);

        //    Assert.AreEqual(expected_66_Vy, actual_66_Vy);
        //}
        //[TestMethod]
        //public void KiemTraDiemCoThuocPhuongTrinh_ad_66_Vy()
        //{
        //    c_66_Vy = new TimNghiem_66_Vy(0, 2);
        //    double x_66_Vy = 2;
        //    string expected_66_Vy = "Giá trị 2 không thuộc phương trình 0x + 2 = 0";
        //    string actual_66_Vy = c_66_Vy.KiemTraDiemCoThuocPhuongTrinh_66_Vy(x_66_Vy);

        //    Assert.AreEqual(expected_66_Vy, actual_66_Vy);
        //}
        //[TestMethod]
        //public void KiemTraDiemCoThuocPhuongTrinh_be_66_Vy()
        //{
        //    c_66_Vy = new TimNghiem_66_Vy(2, -10);
        //    double x_66_Vy = 5;
        //    string expected_66_Vy = "Giá trị 5 thuộc phương trình 2x + -10 = 0";
        //    string actual_66_Vy = c_66_Vy.KiemTraDiemCoThuocPhuongTrinh_66_Vy(x_66_Vy);

        //    Assert.AreEqual(expected_66_Vy, actual_66_Vy);
        //}
        //[TestMethod]
        //public void KiemTraDiemCoThuocPhuongTrinh_bf_66_Vy()
        //{
        //    c_66_Vy = new TimNghiem_66_Vy(1, 10);
        //    double x_66_Vy = 6;
        //    string expected_66_Vy = "Giá trị 6 không thuộc phương trình 1x + 10 = 0";
        //    string actual_66_Vy = c_66_Vy.KiemTraDiemCoThuocPhuongTrinh_66_Vy(x_66_Vy);

        //    Assert.AreEqual(expected_66_Vy, actual_66_Vy);
        //}

        //Tạo 1 đối tượng gồm 4 thuộc tính(tương ứng 5 cột trong file csv)
        public class DataRecordKiemTraDiem_66_Vy
        {
            public int a_66_Vy { get; set; }
            public int b_66_Vy { get; set; }
            public double x_66_Vy { get; set; }
            public string expected_66_Vy { get; set; }
            public string result_66_Vy { get; set; }
        }
        //--------------------------Unit test phương thức 2 data không ghi kết quả---------------------------------------------------
        //[TestMethod]
        //[DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV",
        //    @".\Data_66_Vy\DataUnitTestKiemTRaDiem_66_Vy.csv",
        //    "DataUnitTestKiemTRaDiem_66_Vy#csv", DataAccessMethod.Sequential)]
        //public void DataUnitTestKiemTraDiem()
        //{
        //    int a_66_Vy, b_66_Vy;
        //    double x_66_Vy;
        //    string actual_66_Vy, expected_66_Vy;
        //    a_66_Vy = int.Parse(TestContext.DataRow[0].ToString());
        //    b_66_Vy = int.Parse(TestContext.DataRow[1].ToString());
        //    x_66_Vy = double.Parse(TestContext.DataRow[2].ToString());
        //    expected_66_Vy = TestContext.DataRow[3].ToString();
        //    expected_66_Vy = Encoding.UTF8.GetString(Encoding.Default.GetBytes(expected_66_Vy));
        //    c_66_Vy = new TimNghiem_66_Vy(a_66_Vy, b_66_Vy);
        //    actual_66_Vy = c_66_Vy.KiemTraDiemCoThuocPhuongTrinh_66_Vy(x_66_Vy);
        //    Assert.AreEqual(expected_66_Vy, actual_66_Vy);//Ném ngoại lệ kiểm tra pass hay fail
        //}

        //--------------------------Unit test phương thức 2 data ghi kết quả---------------------------------------------------
        //[TestMethod]
        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV",
            @".\Data_66_Vy\DataUnitTestKiemTRaDiem_66_Vy.csv",
            "DataUnitTestKiemTRaDiem_66_Vy#csv", DataAccessMethod.Sequential)]
        public void DataUnitTestKiemTraDiem()
        {
            int a_66_Vy, b_66_Vy;
            double x_66_Vy;
            string actual_66_Vy, expected_66_Vy = "";
            a_66_Vy = int.Parse(TestContext.DataRow[0].ToString());
            b_66_Vy = int.Parse(TestContext.DataRow[1].ToString());
            x_66_Vy = double.Parse(TestContext.DataRow[2].ToString());
            c_66_Vy = new TimNghiem_66_Vy(a_66_Vy, b_66_Vy);
            actual_66_Vy = c_66_Vy.KiemTraDiemCoThuocPhuongTrinh_66_Vy(x_66_Vy);
            string csvFilePath_66_Vy =
                @"D:\Myseft\Ki2_nam3\KiemThu\BTL_Copy\BTL_66_Vy\BTL_KiemThuPhanMem_66_NguyenThiHienVy\UnitTestTInhNghiem_66_Vy\Data_66_Vy\DataUnitTestKiemTRaDiem_66_Vy.csv";
            var recordsKiemTraDiem_66_Vy = new List<DataRecordKiemTraDiem_66_Vy>();//Khai báo 1 danh sách có kiểu dữ liệu là DataRecordKiemTraDiem
            //Đọc dữ liệu từ file csv
            using (var reader_66_Vy = new StreamReader(csvFilePath_66_Vy))//mở tập tin CSV được chỉ định bởi biến csvFilePath để đọc nó
            using (var csv_66_Vy = new CsvReader(reader_66_Vy, new CsvConfiguration(CultureInfo.InvariantCulture)))
                recordsKiemTraDiem_66_Vy = csv_66_Vy.GetRecords<DataRecordKiemTraDiem_66_Vy>().ToList();//đọc dữ liệu từ tập tin
            //Ghi pass hay fail vào cột result
            for (int i_66_Vy = 0; i_66_Vy < recordsKiemTraDiem_66_Vy.Count; i_66_Vy++)
            {
                //Kiểm tra đúng hàng mới so sánh
                if (recordsKiemTraDiem_66_Vy[i_66_Vy].a_66_Vy == a_66_Vy && recordsKiemTraDiem_66_Vy[i_66_Vy].b_66_Vy == b_66_Vy && recordsKiemTraDiem_66_Vy[i_66_Vy].x_66_Vy == x_66_Vy)
                {
                    expected_66_Vy = recordsKiemTraDiem_66_Vy[i_66_Vy].expected_66_Vy;//Gáng lại biến expected để ghi vào file csv, gặp trường hợp số thực thì định dạng số thực khác nhau giữa c# và csv
                    if (recordsKiemTraDiem_66_Vy[i_66_Vy].expected_66_Vy == actual_66_Vy)//So sánh xem giá trị mong đợi và giá trị thực tế có giống nhau hay không
                        recordsKiemTraDiem_66_Vy[i_66_Vy].result_66_Vy = "pass";
                    else
                        recordsKiemTraDiem_66_Vy[i_66_Vy].result_66_Vy = "fail";
                }
            }
            // Ghi dữ liệu vào tệp CSV
            using (var writer_66_Vy = new StreamWriter(csvFilePath_66_Vy, false))//false: Xoá file trống và ghi dữ liệu mới, True: Giữ dữ liệu cũ và ghi nối tiếp
            using (var csvWriter_66_Vy = new CsvWriter(writer_66_Vy, System.Globalization.CultureInfo.InvariantCulture))
                csvWriter_66_Vy.WriteRecords(recordsKiemTraDiem_66_Vy);//Viết dữ liệu vào file csv
            Assert.AreEqual(expected_66_Vy, actual_66_Vy);//Ném ngoại lệ kiểm tra pass hay fail
        }
    }
}
