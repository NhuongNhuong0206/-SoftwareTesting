using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BTL_KiemThuPhanMem_66_NguyenThiHienVy
{
    public class TimNghiem_66_Vy
    {
        private int a_66_Vy, b_66_Vy;

        public TimNghiem_66_Vy(int a_66_Vy, int b_66_Vy)
        {
            this.a_66_Vy = a_66_Vy;
            this.b_66_Vy = b_66_Vy;
        }
        public double TimNghiemPhuongTrinh_66_Vy()
        {
            double x_66_Vy;//s1
            if (a_66_Vy == 0)//c1
            {   //nếu a = 0 và b = 0 thì phương trình vô số nghiệm, trả về giá trị 0.0001
                if (b_66_Vy == 0)//c2
                    x_66_Vy = 0.001;//s2
                //nếu a = 0 và b != 0 thì phương trình vô nghiệm, trả về giá trị -0.0001
                else//s3
                    x_66_Vy = -0.001;//s4
            }
            //trường hợp a != 0 thì ta tính theo công thức x = -b/a
            else//s5
            {
                x_66_Vy = (double)-b_66_Vy / a_66_Vy;//s6
                //làm tròn kết quả lên 2 số thập phân             
                x_66_Vy = Math.Round(x_66_Vy, 2); //s7
            }
            return x_66_Vy;//s8
        }
        public string KiemTraDiemCoThuocPhuongTrinh_66_Vy(double x_66_Vy)
        {// Tạo một đối tượng của lớp hiện tại
            TimNghiem_66_Vy timNghiem_66_Vy = new TimNghiem_66_Vy(a_66_Vy, b_66_Vy);//s1
            if (a_66_Vy == 0)//c1
            {
                if (b_66_Vy == 0)//c2ư
                    return "Giá trị " + x_66_Vy + " thuộc phương trình " + a_66_Vy + "x + " + b_66_Vy + " = 0";//s2
                else//s3
                    return "Giá trị " + x_66_Vy + " không thuộc phương trình " + a_66_Vy + "x + " + b_66_Vy + " = 0";//s4
            }
            else//s5
            {
                double nghiem_66_Vy = timNghiem_66_Vy.TimNghiemPhuongTrinh_66_Vy();//s6
                if (nghiem_66_Vy == x_66_Vy)//c3
                    return "Giá trị " + x_66_Vy + " thuộc phương trình " + a_66_Vy + "x + " + b_66_Vy + " = 0";//s7
                else//s8
                    return "Giá trị " + x_66_Vy + " không thuộc phương trình " + a_66_Vy + "x + " + b_66_Vy + " = 0";//s9
            }
        }
    }
}
