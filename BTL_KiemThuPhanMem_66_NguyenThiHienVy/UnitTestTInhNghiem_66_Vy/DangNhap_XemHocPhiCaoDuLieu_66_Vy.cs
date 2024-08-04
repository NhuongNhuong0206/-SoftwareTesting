using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Threading;
namespace UnitTestTInhNghiem_66_Vy
{
    [TestClass]
    public class DangNhap_XemHocPhiCaoDuLieu_66_Vy
    {
        public void TaoFileExcel_66_Vy(DataTable dataTable_66_Vy, string sheetName_66_Vy, string title_66_Vy, string savePath_66_Vy)
        {
            //T?o các ð?i tý?ng Excel
            //T?o ð?i tý?ng ð?i di?n cho ?ng d?ng excel tên DoiTuongExcel_66_Vy
            Microsoft.Office.Interop.Excel.Application DoiTuongExcel_66_Vy = new Microsoft.Office.Interop.Excel.Application();
            //T?o ð?i tý?ng ð?i di?n cho t?p h?p các workbook trong ?ng d?ng excel tên DuoiTuongBooks_66_Vy
            Microsoft.Office.Interop.Excel.Workbooks DuoiTuongBooks_66_Vy;
            //T?o ð?i tý?ng ð?i di?n cho t?p h?p các worksheet trong 1 workbook tên DuoiTuongSheets_66_Vy
            Microsoft.Office.Interop.Excel.Sheets DuoiTuongSheets_66_Vy;
            //T?o ð?i tý?ng ð?i di?n cho workbook c? th? trong excel tên DuoiTuongBook_66_Vy
            Microsoft.Office.Interop.Excel.Workbook DuoiTuongBook_66_Vy;
            //T?o ð?i tý?ng ð?i di?n cho worksheet c? th? trong excel tên DuoiTuongSheet_66_Vy
            Microsoft.Office.Interop.Excel.Worksheet DuoiTuongSheet_66_Vy;

            //T?o m?i 1 Excel WorkBook
            //Hi?n th? ?ng d?ng trên màn h?nh khi ch?y, n?u b?ng false th? s? không hi?n th? ?ng d?ng lên màn h?nh
            DoiTuongExcel_66_Vy.Visible = true;
            //T?t các thông báo c?nh báo giúp ?ng d?ng ch?y không b? gián ðo?n 
            DoiTuongExcel_66_Vy.DisplayAlerts = false;
            //T?o 1 trang tính cho m?i workbook m?i ðý?c t?o
            DoiTuongExcel_66_Vy.Application.SheetsInNewWorkbook = 1;
            //Gán các workbooks vào ð?i tý?ng excel ð? th?c hi?n các thao tác liên quan
            DuoiTuongBooks_66_Vy = DoiTuongExcel_66_Vy.Workbooks;
            //T?o m?t workbook m?i trong ?ng d?ng Excel và gán nó vào bi?n DuoiTuongBook
            DuoiTuongBook_66_Vy = (Microsoft.Office.Interop.Excel.Workbook)(DoiTuongExcel_66_Vy.Workbooks.Add(Type.Missing));
            //Gán t?p h?p các worksheet trong workbook m?i ðý?c t?o vào bi?n DuoiTuongSheets
            DuoiTuongSheets_66_Vy = DuoiTuongBook_66_Vy.Worksheets;
            //Truy c?p vào sheet ð?u tiên trong workbook m?i t?o và gán nó vào bi?n DuoiTuongSheet
            DuoiTuongSheet_66_Vy = (Microsoft.Office.Interop.Excel.Worksheet)DuoiTuongSheets_66_Vy.get_Item(1);
            //Ð?t tên cho worksheet là tên ðý?c truy?n bào t? tham s? sheetName
            DuoiTuongSheet_66_Vy.Name = sheetName_66_Vy;

            //T?o tiêu ð?
            //Kh?i t?o m?t ð?i tý?ng Range (head) ð?i di?n cho ô ? c?t 1, hàng 1 trong worksheet 
            Microsoft.Office.Interop.Excel.Range head_66_Vy = DuoiTuongSheet_66_Vy.Cells[1, 1];
            head_66_Vy = head_66_Vy.get_Resize(1, 7); // Thay ð?i kích thý?c Range ð? bao g?m 7 c?t
            head_66_Vy.Merge();// Merge các ô trong ph?m vi ð? ch?n thành m?t ô duy nh?t
                               //head.MergeCells = true;//Ð?m b?o r?ng các ô ð? merge không b? split ra sau khi ðý?c merge.
            head_66_Vy.Value2 = title_66_Vy;//Gán tên b?ng là tham s? ðý?c truy?n bào title
            head_66_Vy.Font.Bold = true;//In ð?m tiêu ð?
            head_66_Vy.Font.Size = "20";//Ch?nh size tiêu ð?
            head_66_Vy.Font.Name = "Times New Roman";//Ch?nh font ch? tiêu ð?
                                                     //Cãn gi?a n?i dung c?a Range head theo chi?u ngang
            head_66_Vy.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            //T?o tiêu ð? c?t
            //Kh?i t?o ð?i tý?ng ð?i di?n cho c?t A t?i hàng s? 3 (là 1 ô duy nh?t) trong worksheet
            Microsoft.Office.Interop.Excel.Range col1_66_Vy = DuoiTuongSheet_66_Vy.get_Range("A3", "A3");
            col1_66_Vy.Value2 = "Stt";//Gáng giá tr? cho ô 
            col1_66_Vy.ColumnWidth = 5;//Ð?t chi?u r?ng
            Microsoft.Office.Interop.Excel.Range col2_66_Vy = DuoiTuongSheet_66_Vy.get_Range("B3", "B3");
            col2_66_Vy.Value2 = "Niên h?c h?c k?";
            col2_66_Vy.ColumnWidth = 25.0;
            Microsoft.Office.Interop.Excel.Range col3_66_Vy = DuoiTuongSheet_66_Vy.get_Range("C3", "C3");
            col3_66_Vy.Value2 = "HP chýa gi?m";
            col3_66_Vy.ColumnWidth = 20.0;
            Microsoft.Office.Interop.Excel.Range col4_66_Vy = DuoiTuongSheet_66_Vy.get_Range("D3", "D3");
            col4_66_Vy.Value2 = "Mi?n gi?m";
            col4_66_Vy.ColumnWidth = 20.0;
            Microsoft.Office.Interop.Excel.Range col5_66_Vy = DuoiTuongSheet_66_Vy.get_Range("E3", "E3");
            col5_66_Vy.Value2 = "Ph?i thu";
            col5_66_Vy.ColumnWidth = 20.0;
            Microsoft.Office.Interop.Excel.Range col6_66_Vy = DuoiTuongSheet_66_Vy.get_Range("F3", "F3");
            col6_66_Vy.Value2 = "Ð? thu";
            col6_66_Vy.ColumnWidth = 20.0;
            Microsoft.Office.Interop.Excel.Range col7_66_Vy = DuoiTuongSheet_66_Vy.get_Range("G3", "G3");
            col7_66_Vy.Value2 = "C?n n?";
            col7_66_Vy.ColumnWidth = 15.0;

            //T?o ð?i tý?ng ð?i di?n cho ph?m vi in ð?m
            Microsoft.Office.Interop.Excel.Range rowHead_66_Vy = DuoiTuongSheet_66_Vy.get_Range("A3", "G3");
            //Th?c hi?n in ð?m ph?m vi ? trên
            rowHead_66_Vy.Font.Bold = true;
            //K? vi?n
            rowHead_66_Vy.Borders.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlSolid;
            //Thi?t l?p màu n?n
            rowHead_66_Vy.Interior.ColorIndex = 6;
            //Cãn gi?a n?i dung theo hàng ngang
            rowHead_66_Vy.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            //T?o m?ng theo datatable
            object[,] arr_66_Vy = new object[dataTable_66_Vy.Rows.Count, dataTable_66_Vy.Columns.Count];
            //Chuy?n d? li?u t? dataTable vào m?ng ð?i tý?ng arr
            for (int row_66_Vy = 0; row_66_Vy < dataTable_66_Vy.Rows.Count; row_66_Vy++)
            {
                //Khai báo 1 hàng và gán d? li?u c?a t?ng hàng t? dataTable vào nó
                DataRow dataRow_66_Vy = dataTable_66_Vy.Rows[row_66_Vy];

                for (int col_66_Vy = 0; col_66_Vy < dataTable_66_Vy.Columns.Count; col_66_Vy++)
                {
                    //Gáng d? thi?u t? hàng d? li?u trên vào m?ng
                    arr_66_Vy[row_66_Vy, col_66_Vy] = dataRow_66_Vy[col_66_Vy];
                }
            }
            //Thi?t l?p vùng ði?n d? li?u
            int rowStart_66_Vy = 4, columnStart_66_Vy = 1;
            //Hàng k?t th?c b?ng hàng b?t ð?u + s? lý?ng hàng c?a b?ng - 1: ví d?: 4 + 8 - 1 = 11
            //D? li?u 8 hàng s? ðý?c nh?p t? hàng 4 ð?n hàng 11 g?m 8 d?ng: 11 - 4 + 1 = 8
            int rowEnd_66_Vy = rowStart_66_Vy + dataTable_66_Vy.Rows.Count - 1;
            //C?t b?t ð?u t? l? trái v? v?y c?t k?t th?c c?ng ðúng b?ng s? lý?ng c?t c?a b?ng d? li?u
            int columnEnd_66_Vy = dataTable_66_Vy.Columns.Count;

            //Ô b?t ð?u ði?n d? li?u
            Microsoft.Office.Interop.Excel.Range c1_66_Vy = (Microsoft.Office.Interop.Excel.Range)DuoiTuongSheet_66_Vy.Cells[rowStart_66_Vy, columnStart_66_Vy];
            //Ô k?t thúc ði?n d? li?u
            Microsoft.Office.Interop.Excel.Range c2_66_Vy = (Microsoft.Office.Interop.Excel.Range)DuoiTuongSheet_66_Vy.Cells[rowEnd_66_Vy, columnEnd_66_Vy];
            //L?y v? vùng ði?n d? li?u 
            Microsoft.Office.Interop.Excel.Range range_66_Vy = DuoiTuongSheet_66_Vy.get_Range(c1_66_Vy, c2_66_Vy);
            //Ði?n d? li?u vào cùng ð? thi?t l?p
            range_66_Vy.Value2 = arr_66_Vy;
            //K? vi?n
            range_66_Vy.Borders.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlSolid;
            DuoiTuongBook_66_Vy.SaveAs(savePath_66_Vy); // Lýu workbook t?i ðý?ng d?n savePath
            DuoiTuongBook_66_Vy.Close(); // Ðóng workbook sau khi lýu
            DoiTuongExcel_66_Vy.Quit();//Ðóng ?ng d?ng excel
        }
          
        //[TestMethod]
        public void DangNhap_XemHocPhi_CaoDuLieuHocPhi_66_Vy()
        {   //Kh?i t?o ð?i dý?ng chrome
            IWebDriver driver_66_HienVy = new ChromeDriver();
            driver_66_HienVy.Manage().Window.Maximize();
            //Kh?i ch?y trang web
            driver_66_HienVy.Navigate().GoToUrl("https://tienichsv.ou.edu.vn/#/home");
            WebDriverWait wait = new WebDriverWait(driver_66_HienVy, TimeSpan.FromSeconds(10));
            //Ch? 5s ð? các element ðý?c c?p nh?p ð?
            driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            //L?t nút ðãng nh?p b?ng XPath
            IWebElement btnLogin_66_HienVy = driver_66_HienVy.FindElement(By.XPath("/html/body/app-root/div/" +
                "div/div/div[1]/div/div/div[2]/app-right/app-login/div/div[2]/div/div[2]/button[2]"));
            btnLogin_66_HienVy.Click();//Th?c hi?n cleck bào elemant logon v?a l?y ðý?c
                                        //Ch? 5s ð? các element ðý?c c?p nh?p ð?
            driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            // L?y t?t c? các c?a s? hi?n t?i
            List<string> windowHandles_66_Vy = driver_66_HienVy.WindowHandles.ToList();
            // Chuy?n sang c?a s? m?i nh?t (c?a s? cu?i cùng trong danh sách)
            driver_66_HienVy.SwitchTo().Window(windowHandles_66_Vy[windowHandles_66_Vy.Count - 1]);
            //L?y danh sách các element b?ng TagName có tên là input (là hàng trong b?ng)
            IReadOnlyCollection<IWebElement> elemenstLogonNew_66_Vy = driver_66_HienVy.FindElements(By.TagName("input"));
            int count_66_Vy = 0;
            foreach (IWebElement element_66_Vy in elemenstLogonNew_66_Vy)
            {
                if (count_66_Vy == 1)//input có index th? 1 là ô nh?p mssv
                {
                    element_66_Vy.SendKeys("2151050567");//G?i mssv lên input này
                   
                }
                else if (count_66_Vy == 2)//input có index th? 2 là ô nh?p m?t kh?u
                {
                    element_66_Vy.SendKeys("******");//G?i m?t kh?u lên input này, m?t kh?u công khai không ðúng
                }
                count_66_Vy += 1;
            }
            //Ch? 5s ð? các element ðý?c c?p nh?p ð?
            driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            //L?y 1 element login ? trang xác th?c b?ng ClassName
            IWebElement btnLoginNew_66_HienVy = driver_66_HienVy.FindElement(By.ClassName("m-loginbox-submit-btn"));
            //Nh?n nút vào nút ðãng nh?p v?a l?y ðý?c
            btnLoginNew_66_HienVy.Click();
            //Ð?i ð? các element ðý?c xu?t hi?n ð?y ð?
            driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            //V? nút xem h?c phí và nút xem th?i khoá bi?u không trong khung nh?n v? v?y c?n Scroll xu?ng 1 tí
            //S? d?ng JavaScript ð? th?c hi?n cu?n
            IJavaScriptExecutor js_66_Vy = (IJavaScriptExecutor)driver_66_HienVy;
            //Cu?n xu?ng 200px
            //Ð?i ð? các element ðý?c xu?t hi?n ð?y ð?
            driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(15);
            js_66_Vy.ExecuteScript("window.scrollBy(0, 200);");

            //L?y nút xem h?c phí b?ng ID
            IWebElement btnViewTuition_66_HienVy = driver_66_HienVy.FindElement(By.Id("WEB_HOCPHI"));
            //Nh?n nút xem h?c phí
            btnViewTuition_66_HienVy.Click();
            //Khai báo m?t ð?i tý?ng b?ng
            DataTable dataTable_66_Vy = new DataTable();
            //T?o các ð?i tý?ng c?t và gáng tên
            DataColumn col1_66_Vy = new DataColumn("Stt");
            DataColumn col2_66_Vy = new DataColumn("Niên h?c k?");
            DataColumn col3_66_Vy = new DataColumn("HP chýa gi?m");
            DataColumn col4_66_Vy = new DataColumn("Mi?n gi?m");
            DataColumn col5_66_Vy = new DataColumn("Ph?i thu");
            DataColumn col6_66_Vy = new DataColumn("Ð? thu");
            DataColumn col7_66_Vy = new DataColumn("C?n n?");
            //Thêm các ð?i tý?ng c?t vào ð?i tý?ng b?ng
            dataTable_66_Vy.Columns.Add(col1_66_Vy);
            dataTable_66_Vy.Columns.Add(col2_66_Vy);
            dataTable_66_Vy.Columns.Add(col3_66_Vy);
            dataTable_66_Vy.Columns.Add(col4_66_Vy);
            dataTable_66_Vy.Columns.Add(col5_66_Vy);
            dataTable_66_Vy.Columns.Add(col6_66_Vy);
            dataTable_66_Vy.Columns.Add(col7_66_Vy);

            //Cào d? li?u
            //Ch? 5s ð? các element ðý?c c?p nh?p ð?
            driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            //L?y 1 elemant là thân c?a b?ng b?ng TagName
            IWebElement elemensTableDataP_66_Vy = driver_66_HienVy.FindElement(By.TagName("tbody"));
            //Ch? 5s ð? các element ðý?c c?p nh?p ð?
            driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            //L?y danh sách các element là hàng c?a b?ng b?ng TagName tr
            IReadOnlyCollection<IWebElement> elemensTableDataChilds_66_Vy = elemensTableDataP_66_Vy.FindElements(By.TagName("tr"));
            //Chuy?n danh sác element v?a l?y ðý?c v? d?ng list
            List<IWebElement> elemensTableDataChildsList_66_Vy = new List<IWebElement>(elemensTableDataChilds_66_Vy);
            List<IWebElement> dataList_66_Vy = new List<IWebElement>();//khai báo chu?i r?ng 
                                                                        //datas ðý?c khai báo giá tr? ban ð?u là r?ng
            IReadOnlyCollection<IWebElement> datas_66_Vy = new ReadOnlyCollection<IWebElement>(dataList_66_Vy);
            //Xoá hàng cu?i cào ðý?c ? trên
            elemensTableDataChildsList_66_Vy.RemoveAt(elemensTableDataChildsList_66_Vy.Count - 1);
            //Xoá hàng ð?u cào ðý?c ? trên
            elemensTableDataChildsList_66_Vy.RemoveAt(0);
            foreach (IWebElement elament_66_Vy in elemensTableDataChildsList_66_Vy)
            {
                DataRow dataR_66_Vy = dataTable_66_Vy.NewRow();//Khai báo m?t ð?i tý?ng hàng
                datas_66_Vy = elament_66_Vy.FindElements(By.TagName("td"));//L?y danh sách các element b?ng TagName 
                int i_66_Vy = 0;
                foreach (IWebElement d_66_Vy in datas_66_Vy)
                {
                    dataR_66_Vy[i_66_Vy] = d_66_Vy.Text;//Gáng d? li?u vào t?ng ð?i tý?ng hàng
                    i_66_Vy++;
                }
                //Gáng d? li?u c?a hàng vào ð?i tý?ng b?ng 
                dataTable_66_Vy.Rows.Add(dataR_66_Vy);
            }
            //Khai báo ðý?ng d?n ð? lýu file excel
            string savePath_66_Vy =
                @"D:\Myseft\Ki2_nam3\KiemThu\BTL_Copy\BTL_66_Vy\BTL_KiemThuPhanMem_66_NguyenThiHienVy\UnitTestTInhNghiem_66_Vy\Data_66_Vy\HocPhi.xlsx";

            //G?i phýõng th?c TaoFileExcel_66_Vy vào truy?n các tham s? c?n thi?t
            TaoFileExcel_66_Vy(dataTable_66_Vy, "H?c phí", "B?ng h?c phí các k?", savePath_66_Vy);
            //Ðóng tr?nh duy?t và màn h?nh ðen
            driver_66_HienVy.Quit();
        }
    }
}
