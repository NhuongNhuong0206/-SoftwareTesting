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
    public class DangNhap_XemThoiKhoaBieuTuan_66_Vy
    {
        //[TestMethod]
        public void DangNhap_XemThoiKhoaBieuTuanTheoMonHoc_Pass_66_Vy()
        {   //Khởi tạo đối dượng chrome
            IWebDriver driver_66_HienVy = new ChromeDriver();
            driver_66_HienVy.Manage().Window.Maximize();
            //Khởi chạy trang web
            driver_66_HienVy.Navigate().GoToUrl("https://tienichsv.ou.edu.vn/#/home");

            //Chờ 5s để các element được cập nhập đủ
            driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(15);
            //Lất nút đăng nhập bằng XPath
            IWebElement btnLogin_66_HienVy = driver_66_HienVy.FindElement(By.XPath("/html/body/app-root/div/" +
                "div/div/div[1]/div/div/div[2]/app-right/app-login/div/div[2]/div/div[2]/button[2]"));
            btnLogin_66_HienVy.Click();//Thực hiện cleck bào elemant logon vừa lấy được
                                       //Chờ 5s để các element được cập nhập đủ
                                       //driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            Thread.Sleep(2);

            // Lấy tất cả các cửa sổ hiện tại
            List<string> windowHandles_66_Vy = driver_66_HienVy.WindowHandles.ToList();
            // Chuyển sang cửa sổ mới nhất (cửa sổ cuối cùng trong danh sách)
            driver_66_HienVy.SwitchTo().Window(windowHandles_66_Vy[windowHandles_66_Vy.Count - 1]);
            //Lấy danh sách các element bằng TagName có tên là input (là hàng trong bảng)
            IReadOnlyCollection<IWebElement> elemenstLogonNew_66_Vy = driver_66_HienVy.FindElements(By.TagName("input"));
            int count_66_Vy = 0;
            foreach (IWebElement element_66_Vy in elemenstLogonNew_66_Vy)
            {
                if (count_66_Vy == 1)//input có index thứ 1 là ô nhập mssv
                {
                    element_66_Vy.SendKeys("2151050567");//Gửi mssv lên input này
                }
                else if (count_66_Vy == 2)//input có index thứ 2 là ô nhập mật khẩu
                {
                    element_66_Vy.SendKeys("*******");//Gửi mật khẩu lên input này, mật khẩu ẩn để công khai mã code
                }
                count_66_Vy += 1;
            }
            //Chờ 5s để các element được cập nhập đủ
            Thread.Sleep(2000);
            //Lấy 1 element login ở trang xác thực bằng ClassName
            IWebElement btnLoginNew_66_HienVy = driver_66_HienVy.FindElement(By.ClassName("m-loginbox-submit-btn"));
            //Nhấn nút vào nút đăng nhập vừa lấy được
            btnLoginNew_66_HienVy.Click();
            //Đợi để các element được xuất hiện đầy đủ
            Thread.Sleep(2000);
            //Vì nút xem học phí và nút xem thời khoá biểu không trong khung nhìn vì vậy cần Scroll xuống 1 tí
            //Sử dụng JavaScript để thực hiện cuộn
            IJavaScriptExecutor js_66_Vy = (IJavaScriptExecutor)driver_66_HienVy;
            //Cuộn xuống 200px
            //Đợi để các element được xuất hiện đầy đủ
            js_66_Vy.ExecuteScript("window.scrollBy(0, 200);");
            Thread.Sleep(3000);
            //Lấy element xem thời khoá biểu tuần
            IWebElement btnViewTimetable_66_HienVy = driver_66_HienVy.FindElement(By.Id("WEB_TKB_1TUAN"));
            btnViewTimetable_66_HienVy.Click();
            //Chờ 5s để các element được cập nhập đủ
            Thread.Sleep(2000);
            //driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(15);
            //Lấy element chọn thời khoá biểu
            IWebElement btnViewTimetableList_66_HienVy = driver_66_HienVy.FindElement(By.XPath("//*[@id='fullScreen']/div[2]/div[1]/div[2]/ng-select"));
            Thread.Sleep(2000);
            //driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(15);
            //Click vào element này để danh sách các thời khoá biểu được hiện ra
            btnViewTimetableList_66_HienVy.Click();
            //Chờ 5s để các element được cập nhập đủ
            Thread.Sleep(2000);
            //driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            //Lấy element cha chứa danh sách các thời khoá biểu
            IWebElement btnViewTimetableSubjectList_66_HienVy = driver_66_HienVy.FindElement(By.ClassName("scrollable-content"));
            //Từ element cha ở trên lấy ra các element con lưu vào danh sách
            IList<IWebElement> btnViewTimetableSubject_66_HienVy = btnViewTimetableSubjectList_66_HienVy.FindElements(By.TagName("div"));
            //Nhấn vào element có index thứ 3 là Thời khoá biểu môn học
            btnViewTimetableSubject_66_HienVy[3].Click();
            //Chờ 2s để các element được cập nhập đủ
            Thread.Sleep(2000);
            //Lấy elemnt để tiềm kiếm môn học
            IWebElement SelectSubject_66_HienVy = driver_66_HienVy.FindElement(By.XPath("/html/body/app-root/div/div/div/div[1]/div/div/div[1]/app-thoikhoabieu-tuan/div/div[2]/div[1]/div[3]"));
            //Nhấn vào element này để các thẻ input nhập từ khoá tìm kiếm môn học hiện ra
            SelectSubject_66_HienVy.Click();
            Thread.Sleep(1000);
            driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(15);
            //Lấy element để gửi từ khoá tìm kiếm
            IWebElement inputSubject_66_HienVy = driver_66_HienVy.FindElement(By.XPath("//*[@id=\"fullScreen\"]/div[2]/div[1]/div[3]/ng-select/div/div/div[3]/input"));
            //Nhấn vào element đó và gửi từ khoá lên
            inputSubject_66_HienVy.SendKeys("Kiểm thử phần mềm");
            //Chờ 5s để các element được cập nhập đủ
            driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            //Lấy element cha chứa danh sách các môn học khớp với từ khoá
            IWebElement KiemThuPhanMem_66_HienVy = driver_66_HienVy.FindElement(By.XPath("/html/body/app-root/div/div/div/div[1]/div/div/div[1]/app-thoikhoabieu-tuan/div/div[2]/div[1]/div[3]/ng-select/ng-dropdown-panel/div/div[2]"));
            //Lưu danh sách các khoá học phù hợp với từ khoá vào danh sách
            IList<IWebElement> divKiemThu_66_Vy = KiemThuPhanMem_66_HienVy.FindElements(By.TagName("div"));
            try
            {
                // Nhấn vào khoá học có index thứ 0 để xem thời khoá biểu
                divKiemThu_66_Vy[1].Click();
                Thread.Sleep(2000);
                //Lấy element chọn thời gian xem thời khoá biểu
                IWebElement dayCha_66_Vy = driver_66_HienVy.FindElement(By.XPath("/html/body/app-root/div/div/div/div[1]/div/div/div[1]/app-thoikhoabieu-tuan/div/div[2]/div[2]"));
                //Click để danh sách thời gian hiện ra
                dayCha_66_Vy.Click();
                Thread.Sleep(2000);
                //Chọn thời gian mong muốn
                IWebElement listDay_66_Vy = driver_66_HienVy.FindElement(By.XPath("/html/body/app-root/div/div/div/div[1]/div/div/div[1]/app-thoikhoabieu-tuan/div/div[2]/div[2]/div[1]/ng-select/ng-dropdown-panel/div/div[2]/div[6]"));
                //Chọn thời gian muốn xem thời khoá biểu
                listDay_66_Vy.Click();
            }
            catch (ArgumentOutOfRangeException ex)
            {
                // Xử lý ngoại lệ ở đây
                Console.WriteLine("Lỗi: " + ex.Message);
            }
            Thread.Sleep(5000);
            driver_66_HienVy.Quit();
        }


        [TestMethod]
        public void DangNhap_XemThoiKhoaBieuTuanTheoMonHoc_Fail_66_Vy()
        {   //Khởi tạo đối dượng chrome
            IWebDriver driver_66_HienVy = new ChromeDriver();
            driver_66_HienVy.Manage().Window.Maximize();
            //Khởi chạy trang web
            driver_66_HienVy.Navigate().GoToUrl("https://tienichsv.ou.edu.vn/#/home");

            //Chờ 5s để các element được cập nhập đủ
            driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(15);
            //Lất nút đăng nhập bằng XPath
            IWebElement btnLogin_66_HienVy = driver_66_HienVy.FindElement(By.XPath("/html/body/app-root/div/" +
                "div/div/div[1]/div/div/div[2]/app-right/app-login/div/div[2]/div/div[2]/button[2]"));
            btnLogin_66_HienVy.Click();//Thực hiện cleck bào elemant logon vừa lấy được
                                       //Chờ 5s để các element được cập nhập đủ
                                       //driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            Thread.Sleep(2);

            // Lấy tất cả các cửa sổ hiện tại
            List<string> windowHandles_66_Vy = driver_66_HienVy.WindowHandles.ToList();
            // Chuyển sang cửa sổ mới nhất (cửa sổ cuối cùng trong danh sách)
            driver_66_HienVy.SwitchTo().Window(windowHandles_66_Vy[windowHandles_66_Vy.Count - 1]);
            //Lấy danh sách các element bằng TagName có tên là input (là hàng trong bảng)
            IReadOnlyCollection<IWebElement> elemenstLogonNew_66_Vy = driver_66_HienVy.FindElements(By.TagName("input"));
            int count_66_Vy = 0;
            foreach (IWebElement element_66_Vy in elemenstLogonNew_66_Vy)
            {
                if (count_66_Vy == 1)//input có index thứ 1 là ô nhập mssv
                {
                    element_66_Vy.SendKeys("2151050567");//Gửi mssv lên input này
                }
                else if (count_66_Vy == 2)//input có index thứ 2 là ô nhập mật khẩu
                {
                    element_66_Vy.SendKeys("75720311");//Gửi mật khẩu lên input này
                }
                count_66_Vy += 1;
            }
            //Chờ 5s để các element được cập nhập đủ
            Thread.Sleep(2000);
            //Lấy 1 element login ở trang xác thực bằng ClassName
            IWebElement btnLoginNew_66_HienVy = driver_66_HienVy.FindElement(By.ClassName("m-loginbox-submit-btn"));
            //Nhấn nút vào nút đăng nhập vừa lấy được
            btnLoginNew_66_HienVy.Click();
            //Đợi để các element được xuất hiện đầy đủ
            Thread.Sleep(2000);
            //Vì nút xem học phí và nút xem thời khoá biểu không trong khung nhìn vì vậy cần Scroll xuống 1 tí
            //Sử dụng JavaScript để thực hiện cuộn
            IJavaScriptExecutor js_66_Vy = (IJavaScriptExecutor)driver_66_HienVy;
            //Cuộn xuống 200px
            //Đợi để các element được xuất hiện đầy đủ
            js_66_Vy.ExecuteScript("window.scrollBy(0, 200);");
            Thread.Sleep(3000);
            //Lấy element xem thời khoá biểu tuần
            IWebElement btnViewTimetable_66_HienVy = driver_66_HienVy.FindElement(By.Id("WEB_TKB_1TUAN"));
            btnViewTimetable_66_HienVy.Click();
            //Chờ 5s để các element được cập nhập đủ
            Thread.Sleep(2000);
            //driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(15);
            //Lấy element chọn thời khoá biểu
            IWebElement btnViewTimetableList_66_HienVy = driver_66_HienVy.FindElement(By.XPath("//*[@id='fullScreen']/div[2]/div[1]/div[2]/ng-select"));
            Thread.Sleep(2000);
            //driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(15);
            //Click vào element này để danh sách các thời khoá biểu được hiện ra
            btnViewTimetableList_66_HienVy.Click();
            //Chờ 5s để các element được cập nhập đủ
            Thread.Sleep(2000);
            //driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            //Lấy element cha chứa danh sách các thời khoá biểu
            IWebElement btnViewTimetableSubjectList_66_HienVy = driver_66_HienVy.FindElement(By.ClassName("scrollable-content"));
            //Từ element cha ở trên lấy ra các element con lưu vào danh sách
            IList<IWebElement> btnViewTimetableSubject_66_HienVy = btnViewTimetableSubjectList_66_HienVy.FindElements(By.TagName("div"));
            //Nhấn vào element có index thứ 3 là Thời khoá biểu môn học
            btnViewTimetableSubject_66_HienVy[3].Click();
            //Chờ 2s để các element được cập nhập đủ
            Thread.Sleep(2000);
            //Lấy elemnt để tiềm kiếm môn học
            IWebElement SelectSubject_66_HienVy = driver_66_HienVy.FindElement(By.XPath("/html/body/app-root/div/div/div/div[1]/div/div/div[1]/app-thoikhoabieu-tuan/div/div[2]/div[1]/div[3]"));
            //Nhấn vào element này để các thẻ input nhập từ khoá tìm kiếm môn học hiện ra
            SelectSubject_66_HienVy.Click();
            Thread.Sleep(1000);
            driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(15);
            //Lấy element để gửi từ khoá tìm kiếm
            IWebElement inputSubject_66_HienVy = driver_66_HienVy.FindElement(By.XPath("//*[@id=\"fullScreen\"]/div[2]/div[1]/div[3]/ng-select/div/div/div[3]/input"));
            //Nhấn vào element đó và gửi từ khoá lên
            inputSubject_66_HienVy.SendKeys("66_HienVy");
            //Chờ 5s để các element được cập nhập đủ
            driver_66_HienVy.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            //Lấy element cha chứa danh sách các môn học khớp với từ khoá
            IWebElement KiemThuPhanMem_66_HienVy = driver_66_HienVy.FindElement(By.XPath("/html/body/app-root/div/div/div/div[1]/div/div/div[1]/app-thoikhoabieu-tuan/div/div[2]/div[1]/div[3]/ng-select/ng-dropdown-panel/div/div[2]"));
            //Lưu danh sách các khoá học phù hợp với từ khoá vào danh sách
            IList<IWebElement> divKiemThu_66_Vy = KiemThuPhanMem_66_HienVy.FindElements(By.TagName("div"));
            try
            {
                // Nhấn vào khoá học có index thứ 0 để xem thời khoá biểu
                divKiemThu_66_Vy[1].Click();
                Thread.Sleep(2000);
                //Lấy element chọn thời gian xem thời khoá biểu
                IWebElement dayCha_66_Vy = driver_66_HienVy.FindElement(By.XPath("/html/body/app-root/div/div/div/div[1]/div/div/div[1]/app-thoikhoabieu-tuan/div/div[2]/div[2]"));
                //Click để danh sách thời gian hiện ra
                dayCha_66_Vy.Click();
                Thread.Sleep(2000);
                //Chọn thời gian mong muốn
                IWebElement listDay_66_Vy = driver_66_HienVy.FindElement(By.XPath("/html/body/app-root/div/div/div/div[1]/div/div/div[1]/app-thoikhoabieu-tuan/div/div[2]/div[2]/div[1]/ng-select/ng-dropdown-panel/div/div[2]/div[6]"));
                //Chọn thời gian muốn xem thời khoá biểu
                listDay_66_Vy.Click();
            }
            catch (ArgumentOutOfRangeException ex)
            {
                // Xử lý ngoại lệ ở đây
                Console.WriteLine("Lỗi: " + ex.Message);
            }
            Thread.Sleep(5000);
            driver_66_HienVy.Quit();
        }

        
    }
}
