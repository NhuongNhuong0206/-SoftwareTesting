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
            //T?o c�c �?i t�?ng Excel
            //T?o �?i t�?ng �?i di?n cho ?ng d?ng excel t�n DoiTuongExcel_66_Vy
            Microsoft.Office.Interop.Excel.Application DoiTuongExcel_66_Vy = new Microsoft.Office.Interop.Excel.Application();
            //T?o �?i t�?ng �?i di?n cho t?p h?p c�c workbook trong ?ng d?ng excel t�n DuoiTuongBooks_66_Vy
            Microsoft.Office.Interop.Excel.Workbooks DuoiTuongBooks_66_Vy;
            //T?o �?i t�?ng �?i di?n cho t?p h?p c�c worksheet trong 1 workbook t�n DuoiTuongSheets_66_Vy
            Microsoft.Office.Interop.Excel.Sheets DuoiTuongSheets_66_Vy;
            //T?o �?i t�?ng �?i di?n cho workbook c? th? trong excel t�n DuoiTuongBook_66_Vy
            Microsoft.Office.Interop.Excel.Workbook DuoiTuongBook_66_Vy;
            //T?o �?i t�?ng �?i di?n cho worksheet c? th? trong excel t�n DuoiTuongSheet_66_Vy
            Microsoft.Office.Interop.Excel.Worksheet DuoiTuongSheet_66_Vy;

            //T?o m?i 1 Excel WorkBook
            //Hi?n th? ?ng d?ng tr�n m�n h?nh khi ch?y, n?u b?ng false th? s? kh�ng hi?n th? ?ng d?ng l�n m�n h?nh
            DoiTuongExcel_66_Vy.Visible = true;
            //T?t c�c th�ng b�o c?nh b�o gi�p ?ng d?ng ch?y kh�ng b? gi�n �o?n 
            DoiTuongExcel_66_Vy.DisplayAlerts = false;
            //T?o 1 trang t�nh cho m?i workbook m?i ��?c t?o
            DoiTuongExcel_66_Vy.Application.SheetsInNewWorkbook = 1;
            //G�n c�c workbooks v�o �?i t�?ng excel �? th?c hi?n c�c thao t�c li�n quan
            DuoiTuongBooks_66_Vy = DoiTuongExcel_66_Vy.Workbooks;
            //T?o m?t workbook m?i trong ?ng d?ng Excel v� g�n n� v�o bi?n DuoiTuongBook
            DuoiTuongBook_66_Vy = (Microsoft.Office.Interop.Excel.Workbook)(DoiTuongExcel_66_Vy.Workbooks.Add(Type.Missing));
            //G�n t?p h?p c�c worksheet trong workbook m?i ��?c t?o v�o bi?n DuoiTuongSheets
            DuoiTuongSheets_66_Vy = DuoiTuongBook_66_Vy.Worksheets;
            //Truy c?p v�o sheet �?u ti�n trong workbook m?i t?o v� g�n n� v�o bi?n DuoiTuongSheet
            DuoiTuongSheet_66_Vy = (Microsoft.Office.Interop.Excel.Worksheet)DuoiTuongSheets_66_Vy.get_Item(1);
            //�?t t�n cho worksheet l� t�n ��?c truy?n b�o t? tham s? sheetName
            DuoiTuongSheet_66_Vy.Name = sheetName_66_Vy;

            //T?o ti�u �?
            //Kh?i t?o m?t �?i t�?ng Range (head) �?i di?n cho � ? c?t 1, h�ng 1 trong worksheet 
            Microsoft.Office.Interop.Excel.Range head_66_Vy = DuoiTuongSheet_66_Vy.Cells[1, 1];
            head_66_Vy = head_66_Vy.get_Resize(1, 7); // Thay �?i k�ch th�?c Range �? bao g?m 7 c?t
            head_66_Vy.Merge();// Merge c�c � trong ph?m vi �? ch?n th�nh m?t � duy nh?t
                               //head.MergeCells = true;//�?m b?o r?ng c�c � �? merge kh�ng b? split ra sau khi ��?c merge.
            head_66_Vy.Value2 = title_66_Vy;//G�n t�n b?ng l� tham s? ��?c truy?n b�o title
            head_66_Vy.Font.Bold = true;//In �?m ti�u �?
            head_66_Vy.Font.Size = "20";//Ch?nh size ti�u �?
            head_66_Vy.Font.Name = "Times New Roman";//Ch?nh font ch? ti�u �?
                                                     //C�n gi?a n?i dung c?a Range head theo chi?u ngang
            head_66_Vy.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            //T?o ti�u �? c?t
            //Kh?i t?o �?i t�?ng �?i di?n cho c?t A t?i h�ng s? 3 (l� 1 � duy nh?t) trong worksheet
            Microsoft.Office.Interop.Excel.Range col1_66_Vy = DuoiTuongSheet_66_Vy.get_Range("A3", "A3");
            col1_66_Vy.Value2 = "Stt";//G�ng gi� tr? cho � 
            col1_66_Vy.ColumnWidth = 5;//�?t chi?u r?ng
            Microsoft.Office.Interop.Excel.Range col2_66_Vy = DuoiTuongSheet_66_Vy.get_Range("B3", "B3");
            col2_66_Vy.Value2 = "Ni�n h?c h?c k?";
            col2_66_Vy.ColumnWidth = 25.0;
            Microsoft.Office.Interop.Excel.Range col3_66_Vy = DuoiTuongSheet_66_Vy.get_Range("C3", "C3");
            col3_66_Vy.Value2 = "HP ch�a gi?m";
            col3_66_Vy.ColumnWidth = 20.0;
            Microsoft.Office.Interop.Excel.Range col4_66_Vy = DuoiTuongSheet_66_Vy.get_Range("D3", "D3");
            col4_66_Vy.Value2 = "Mi?n gi?m";
            col4_66_Vy.ColumnWidth = 20.0;
            Microsoft.Office.Interop.Excel.Range col5_66_Vy = DuoiTuongSheet_66_Vy.get_Range("E3", "E3");
            col5_66_Vy.Value2 = "Ph?i thu";
            col5_66_Vy.ColumnWidth = 20.0;
            Microsoft.Office.Interop.Excel.Range col6_66_Vy = DuoiTuongSheet_66_Vy.get_Range("F3", "F3");
            col6_66_Vy.Value2 = "�? thu";
            col6_66_Vy.ColumnWidth = 20.0;
            Microsoft.Office.Interop.Excel.Range col7_66_Vy = DuoiTuongSheet_66_Vy.get_Range("G3", "G3");
            col7_66_Vy.Value2 = "C?n n?";
            col7_66_Vy.ColumnWidth = 15.0;

            //T?o �?i t�?ng �?i di?n cho ph?m vi in �?m
            Microsoft.Office.Interop.Excel.Range rowHead_66_Vy = DuoiTuongSheet_66_Vy.get_Range("A3", "G3");
            //Th?c hi?n in �?m ph?m vi ? tr�n
            rowHead_66_Vy.Font.Bold = true;
            //K? vi?n
            rowHead_66_Vy.Borders.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlSolid;
            //Thi?t l?p m�u n?n
            rowHead_66_Vy.Interior.ColorIndex = 6;
            //C�n gi?a n?i dung theo h�ng ngang
            rowHead_66_Vy.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            //T?o m?ng theo datatable
            object[,] arr_66_Vy = new object[dataTable_66_Vy.Rows.Count, dataTable_66_Vy.Columns.Count];
            //Chuy?n d? li?u t? dataTable v�o m?ng �?i t�?ng arr
            for (int row_66_Vy = 0; row_66_Vy < dataTable_66_Vy.Rows.Count; row_66_Vy++)
            {
                //Khai b�o 1 h�ng v� g�n d? li?u c?a t?ng h�ng t? dataTable v�o n�
                DataRow dataRow_66_Vy = dataTable_66_Vy.Rows[row_66_Vy];

                for (int col_66_Vy = 0; col_66_Vy < dataTable_66_Vy.Columns.Count; col_66_Vy++)
                {
                    //G�ng d? thi?u t? h�ng d? li?u tr�n v�o m?ng
                    arr_66_Vy[row_66_Vy, col_66_Vy] = dataRow_66_Vy[col_66_Vy];
                }
            }
            //Thi?t l?p v�ng �i?n d? li?u
            int rowStart_66_Vy = 4, columnStart_66_Vy = 1;
            //H�ng k?t th?c b?ng h�ng b?t �?u + s? l�?ng h�ng c?a b?ng - 1: v� d?: 4 + 8 - 1 = 11
            //D? li?u 8 h�ng s? ��?c nh?p t? h�ng 4 �?n h�ng 11 g?m 8 d?ng: 11 - 4 + 1 = 8
            int rowEnd_66_Vy = rowStart_66_Vy + dataTable_66_Vy.Rows.Count - 1;
            //C?t b?t �?u t? l? tr�i v? v?y c?t k?t th?c c?ng ��ng b?ng s? l�?ng c?t c?a b?ng d? li?u
            int columnEnd_66_Vy = dataTable_66_Vy.Columns.Count;

            //� b?t �?u �i?n d? li?u
            Microsoft.Office.Interop.Excel.Range c1_66_Vy = (Microsoft.Office.Interop.Excel.Range)DuoiTuongSheet_66_Vy.Cells[rowStart_66_Vy, columnStart_66_Vy];
            //� k?t th�c �i?n d? li?u
            Microsoft.Office.Interop.Excel.Range c2_66_Vy = (Microsoft.Office.Interop.Excel.Range)DuoiTuongSheet_66_Vy.Cells[rowEnd_66_Vy, columnEnd_66_Vy];
            //L?y v? v�ng �i?n d? li?u 
            Microsoft.Office.Interop.Excel.Range range_66_Vy = DuoiTuongSheet_66_Vy.get_Range(c1_66_Vy, c2_66_Vy);
            //�i?n d? li?u v�o c�ng �? thi?t l?p
            range_66_Vy.Value2 = arr_66_Vy;
            //K? vi?n
            range_66_Vy.Borders.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlSolid;
            DuoiTuongBook_66_Vy.SaveAs(savePath_66_Vy); // L�u workbook t?i ��?ng d?n savePath
            DuoiTuongBook_66_Vy.Close(); // ��ng workbook sau khi l�u
            DoiTuongExcel_66_Vy.Quit();//��ng ?ng d?ng excel
        }
          
        //[TestMethod]
        public void DangNhap_XemHocPhi_CaoDuLieuHocPhi_66_Vy()
        {   //Kh?i t?o �?i d�?ng chrome
            IWebDriver driver_66_HienVy = new ChromeDriver();
            driver_66_HienVy.Manage().Window.Maximize();
            //Kh?i ch?y trang web
            driver_66_HienVy.Navigate().GoToUrl("https://tienichsv.ou.edu.vn/#/home");
            WebDriverWait wait = new WebDriverWait(driver_66_HienVy, TimeSpan.FromSeconds(10));
            //Ch? 5s �? c�c element ��?c c?p nh?p �?
            driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            //L?t n�t ��ng nh?p b?ng XPath
            IWebElement btnLogin_66_HienVy = driver_66_HienVy.FindElement(By.XPath("/html/body/app-root/div/" +
                "div/div/div[1]/div/div/div[2]/app-right/app-login/div/div[2]/div/div[2]/button[2]"));
            btnLogin_66_HienVy.Click();//Th?c hi?n cleck b�o elemant logon v?a l?y ��?c
                                        //Ch? 5s �? c�c element ��?c c?p nh?p �?
            driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            // L?y t?t c? c�c c?a s? hi?n t?i
            List<string> windowHandles_66_Vy = driver_66_HienVy.WindowHandles.ToList();
            // Chuy?n sang c?a s? m?i nh?t (c?a s? cu?i c�ng trong danh s�ch)
            driver_66_HienVy.SwitchTo().Window(windowHandles_66_Vy[windowHandles_66_Vy.Count - 1]);
            //L?y danh s�ch c�c element b?ng TagName c� t�n l� input (l� h�ng trong b?ng)
            IReadOnlyCollection<IWebElement> elemenstLogonNew_66_Vy = driver_66_HienVy.FindElements(By.TagName("input"));
            int count_66_Vy = 0;
            foreach (IWebElement element_66_Vy in elemenstLogonNew_66_Vy)
            {
                if (count_66_Vy == 1)//input c� index th? 1 l� � nh?p mssv
                {
                    element_66_Vy.SendKeys("2151050567");//G?i mssv l�n input n�y
                   
                }
                else if (count_66_Vy == 2)//input c� index th? 2 l� � nh?p m?t kh?u
                {
                    element_66_Vy.SendKeys("******");//G?i m?t kh?u l�n input n�y, m?t kh?u c�ng khai kh�ng ��ng
                }
                count_66_Vy += 1;
            }
            //Ch? 5s �? c�c element ��?c c?p nh?p �?
            driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            //L?y 1 element login ? trang x�c th?c b?ng ClassName
            IWebElement btnLoginNew_66_HienVy = driver_66_HienVy.FindElement(By.ClassName("m-loginbox-submit-btn"));
            //Nh?n n�t v�o n�t ��ng nh?p v?a l?y ��?c
            btnLoginNew_66_HienVy.Click();
            //�?i �? c�c element ��?c xu?t hi?n �?y �?
            driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            //V? n�t xem h?c ph� v� n�t xem th?i kho� bi?u kh�ng trong khung nh?n v? v?y c?n Scroll xu?ng 1 t�
            //S? d?ng JavaScript �? th?c hi?n cu?n
            IJavaScriptExecutor js_66_Vy = (IJavaScriptExecutor)driver_66_HienVy;
            //Cu?n xu?ng 200px
            //�?i �? c�c element ��?c xu?t hi?n �?y �?
            driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(15);
            js_66_Vy.ExecuteScript("window.scrollBy(0, 200);");

            //L?y n�t xem h?c ph� b?ng ID
            IWebElement btnViewTuition_66_HienVy = driver_66_HienVy.FindElement(By.Id("WEB_HOCPHI"));
            //Nh?n n�t xem h?c ph�
            btnViewTuition_66_HienVy.Click();
            //Khai b�o m?t �?i t�?ng b?ng
            DataTable dataTable_66_Vy = new DataTable();
            //T?o c�c �?i t�?ng c?t v� g�ng t�n
            DataColumn col1_66_Vy = new DataColumn("Stt");
            DataColumn col2_66_Vy = new DataColumn("Ni�n h?c k?");
            DataColumn col3_66_Vy = new DataColumn("HP ch�a gi?m");
            DataColumn col4_66_Vy = new DataColumn("Mi?n gi?m");
            DataColumn col5_66_Vy = new DataColumn("Ph?i thu");
            DataColumn col6_66_Vy = new DataColumn("�? thu");
            DataColumn col7_66_Vy = new DataColumn("C?n n?");
            //Th�m c�c �?i t�?ng c?t v�o �?i t�?ng b?ng
            dataTable_66_Vy.Columns.Add(col1_66_Vy);
            dataTable_66_Vy.Columns.Add(col2_66_Vy);
            dataTable_66_Vy.Columns.Add(col3_66_Vy);
            dataTable_66_Vy.Columns.Add(col4_66_Vy);
            dataTable_66_Vy.Columns.Add(col5_66_Vy);
            dataTable_66_Vy.Columns.Add(col6_66_Vy);
            dataTable_66_Vy.Columns.Add(col7_66_Vy);

            //C�o d? li?u
            //Ch? 5s �? c�c element ��?c c?p nh?p �?
            driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            //L?y 1 elemant l� th�n c?a b?ng b?ng TagName
            IWebElement elemensTableDataP_66_Vy = driver_66_HienVy.FindElement(By.TagName("tbody"));
            //Ch? 5s �? c�c element ��?c c?p nh?p �?
            driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            //L?y danh s�ch c�c element l� h�ng c?a b?ng b?ng TagName tr
            IReadOnlyCollection<IWebElement> elemensTableDataChilds_66_Vy = elemensTableDataP_66_Vy.FindElements(By.TagName("tr"));
            //Chuy?n danh s�c element v?a l?y ��?c v? d?ng list
            List<IWebElement> elemensTableDataChildsList_66_Vy = new List<IWebElement>(elemensTableDataChilds_66_Vy);
            List<IWebElement> dataList_66_Vy = new List<IWebElement>();//khai b�o chu?i r?ng 
                                                                        //datas ��?c khai b�o gi� tr? ban �?u l� r?ng
            IReadOnlyCollection<IWebElement> datas_66_Vy = new ReadOnlyCollection<IWebElement>(dataList_66_Vy);
            //Xo� h�ng cu?i c�o ��?c ? tr�n
            elemensTableDataChildsList_66_Vy.RemoveAt(elemensTableDataChildsList_66_Vy.Count - 1);
            //Xo� h�ng �?u c�o ��?c ? tr�n
            elemensTableDataChildsList_66_Vy.RemoveAt(0);
            foreach (IWebElement elament_66_Vy in elemensTableDataChildsList_66_Vy)
            {
                DataRow dataR_66_Vy = dataTable_66_Vy.NewRow();//Khai b�o m?t �?i t�?ng h�ng
                datas_66_Vy = elament_66_Vy.FindElements(By.TagName("td"));//L?y danh s�ch c�c element b?ng TagName 
                int i_66_Vy = 0;
                foreach (IWebElement d_66_Vy in datas_66_Vy)
                {
                    dataR_66_Vy[i_66_Vy] = d_66_Vy.Text;//G�ng d? li?u v�o t?ng �?i t�?ng h�ng
                    i_66_Vy++;
                }
                //G�ng d? li?u c?a h�ng v�o �?i t�?ng b?ng 
                dataTable_66_Vy.Rows.Add(dataR_66_Vy);
            }
            //Khai b�o ��?ng d?n �? l�u file excel
            string savePath_66_Vy =
                @"D:\Myseft\Ki2_nam3\KiemThu\BTL_Copy\BTL_66_Vy\BTL_KiemThuPhanMem_66_NguyenThiHienVy\UnitTestTInhNghiem_66_Vy\Data_66_Vy\HocPhi.xlsx";

            //G?i ph��ng th?c TaoFileExcel_66_Vy v�o truy?n c�c tham s? c?n thi?t
            TaoFileExcel_66_Vy(dataTable_66_Vy, "H?c ph�", "B?ng h?c ph� c�c k?", savePath_66_Vy);
            //��ng tr?nh duy?t v� m�n h?nh �en
            driver_66_HienVy.Quit();
        }
    }
}
