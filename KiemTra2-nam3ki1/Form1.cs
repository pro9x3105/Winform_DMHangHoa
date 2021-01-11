using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace KiemTra2_nam3ki1
{
    public partial class Form1 : Form
    {
        ProcessDataBase connv1 = new ProcessDataBase();
        DataTable dt;
        DataTable dt1;

        public Form1()
        {
            InitializeComponent();
        }

        public void themcbChatLieu()
        {
            for (int i = 0; i < dt1.Rows.Count; i++)
            {
                cbChatLieu.Items.Add(dt1.Rows[i][1].ToString());
            }
        }
        public void Xoatxb()
        {
            txbMaHang.Clear();
            txbTenHang.Clear();
            cbChatLieu.Text = "";
            txbSoLuong.Clear();
            txbGiaNhap.Clear();
            txbGiaBan.Clear();
        }

        private void BtnThoat_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn thoát không ?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                this.Close();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dt = connv1.DocBang("SELECT * FROM tblHang");
            dgvHienThi.DataSource = dt;
            dt1 = connv1.DocBang("SELECT * FROM tblChatLieu");
            themcbChatLieu();
        }

        private void BtnThem_Click(object sender, EventArgs e)
        {
            string maCL = "";
            for (int i = 0; i < dt1.Rows.Count; i++)
            {
                if (cbChatLieu.Text == dt1.Rows[i][1].ToString())
                {
                    maCL = dt1.Rows[i]["MaChatLieu"].ToString();
                    break;
                }
            }
            for (int i = 0; i < dgvHienThi.Rows.Count - 1; i++)
            {
                if (txbMaHang.Text == dgvHienThi.Rows[i].Cells[0].Value.ToString())
                {
                    MessageBox.Show("Có mã hàng này rồi");
                    return;
                }
            }
            connv1.CapNhapDuLieu("INSERT into tblHang(Mahang,Tenhang,Machatlieu,Soluong,Dongianhap,Dongiaban) VALUES ('" + txbMaHang.Text + "',N'" + txbTenHang.Text + "','" + maCL + "','" + txbSoLuong.Text + "','" + txbGiaNhap.Text + "','" + txbGiaBan.Text + "')");
            dgvHienThi.DataSource = connv1.DocBang("SELECT * FROM tblHang");
            Xoatxb();
        }

        private void TxbSoLuong_TextChanged(object sender, EventArgs e)
        {

        }

        private void TxbSoLuong_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void TxbGiaNhap_TextChanged(object sender, EventArgs e)
        {

        }

        private void TxbGiaBan_TextChanged(object sender, EventArgs e)
        {

        }

        private void TxbGiaNhap_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void TxbGiaBan_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void BtnSua_Click(object sender, EventArgs e)
        {
            string maCL = "";
            for (int i = 0; i < dt1.Rows.Count; i++)
            {
                if (cbChatLieu.Text == dt1.Rows[i][1].ToString())
                {
                    maCL = dt1.Rows[i]["MaChatLieu"].ToString();
                    break;
                }
            }
            connv1.CapNhapDuLieu("UPDATE tblHang SET Tenhang='" + txbTenHang.Text + "',Machatlieu='" + maCL + "',Soluong='" + txbSoLuong.Text + "',Dongianhap='" + txbGiaNhap.Text + "',Dongiaban='" + txbGiaBan.Text + "' WHERE Mahang='" + txbMaHang.Text + "'");
            dgvHienThi.DataSource = connv1.DocBang("SELECT * from tblHang");
            Xoatxb();

        }

        private void DgvHienThi_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txbMaHang.Text = dgvHienThi.CurrentRow.Cells[0].Value.ToString();
            txbTenHang.Text = dgvHienThi.CurrentRow.Cells[1].Value.ToString();
            for (int i = 0; i < dt1.Rows.Count; i++)
            {
                if (dgvHienThi.CurrentRow.Cells[2].Value.ToString() == dt1.Rows[i][0].ToString())
                {
                    cbChatLieu.Text = dt1.Rows[i]["TenChatLieu"].ToString();
                    break;
                }
            }
            txbSoLuong.Text = dgvHienThi.CurrentRow.Cells[3].Value.ToString();
            txbGiaNhap.Text = dgvHienThi.CurrentRow.Cells[4].Value.ToString();
            txbGiaBan.Text = dgvHienThi.CurrentRow.Cells[5].Value.ToString();
        }
        
        private void BtnXoa_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                connv1.CapNhapDuLieu("DELETE from tblHang WHERE Mahang='" + txbMaHang.Text + "'");
                dgvHienThi.DataSource = connv1.DocBang("SELECT * from tblHang");
                Xoatxb();
            }
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            if (dgvHienThi.Rows.Count > 0) //TH có dữ liệu được ghi
            {
                Excel.Application exApp = new Excel.Application();
                Excel.Workbook exBook = exApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet exSheet = (Excel.Worksheet)exBook.Worksheets[1];
                //Ðịnh dạng chung
                Excel.Range tenCuaHang = (Excel.Range)exSheet.Cells[1, 1];
                tenCuaHang.Font.Size = 12;
                tenCuaHang.Font.Bold = true; tenCuaHang.Font.Color = Color.Blue;
                tenCuaHang.Value = "CỬA HÀNG BÁN GA";
                Excel.Range dcCuaHang = (Excel.Range)exSheet.Cells[2, 1];
                dcCuaHang.Font.Size = 12;
                dcCuaHang.Font.Bold = true;
                dcCuaHang.Font.Color = Color.Blue;
                dcCuaHang.Value = "Ðịa chỉ: Số 3 - Láng Thuợng - Cầu Giấy - Hà Nội";
                Excel.Range dtCuaHang = (Excel.Range)exSheet.Cells[3, 1];
                dtCuaHang.Font.Size = 12;
                dtCuaHang.Font.Bold = true;
                dtCuaHang.Font.Color = Color.Blue;
                dtCuaHang.Value = "Ðiện thoại: 0961658137";
                Excel.Range header = (Excel.Range)exSheet.Cells[5, 2];
                exSheet.get_Range("B5:G5").Merge(true);
                header.Font.Size = 13;
                header.Font.Bold = true;
                header.Font.Color = Color.Red;
                header.Value = "DANH SÁCH CÁC LO?I GA";
                //Định dạng tiêu đề bảng
                exSheet.get_Range("A7:G7").Font.Bold = true;
                exSheet.get_Range("A7:G7").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                exSheet.get_Range("A7").Value = "STT";
                exSheet.get_Range("B7").Value = "Mã bình";
                exSheet.get_Range("C7").Value = "Tên bình";
                exSheet.get_Range("C7").ColumnWidth = 20;
                exSheet.get_Range("D7").Value = "Mã loại";
                exSheet.get_Range("E7").Value = "Mã màu";
                exSheet.get_Range("F7").Value = "Mã khối luợng";
                exSheet.get_Range("F7").ColumnWidth = 15;
                exSheet.get_Range("G7").Value = "Mã nuớc sản xuất";
                exSheet.get_Range("G7").ColumnWidth = 15;
                
                //In dữ liệu
                for (int i = 0; i < dgvHienThi.Rows.Count; i++)
                {
                    exSheet.get_Range("A" + (i + 8).ToString() + ":G" + (i + 8).ToString()).Font.Bold = false;
                    exSheet.get_Range("A" + (i + 8).ToString()).Value = (i + 1).ToString();
                    exSheet.get_Range("B" + (i + 8).ToString()).Value = dgvHienThi.Rows[i].Cells[0].Value;
                    exSheet.get_Range("C" + (i + 8).ToString()).Value = dgvHienThi.Rows[i].Cells[1].Value;
                    exSheet.get_Range("D" + (i + 8).ToString()).Value = dgvHienThi.Rows[i].Cells[2].Value;
                    exSheet.get_Range("E" + (i + 8).ToString()).Value = dgvHienThi.Rows[i].Cells[3].Value;
                    exSheet.get_Range("F" + (i + 8).ToString()).Value = dgvHienThi.Rows[i].Cells[4].Value;
                    exSheet.get_Range("G" + (i + 8).ToString()).Value = dgvHienThi.Rows[i].Cells[5].Value;
                }
                exSheet.Name = "Hang";
                exBook.Activate(); //Kích hoạt file Excel
                                   //Thiết lập các thuộc tính của SaveFileDialog
                saveExcel.Filter = "Excel Document(*.xls)|*.xls |Word Document(*.doc)| *.doc | All files(*.*) | *.* ";
                saveExcel.FilterIndex = 1;
                saveExcel.AddExtension = true;
                saveExcel.DefaultExt = ".xls";
                if (saveExcel.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    exBook.SaveAs(saveExcel.FileName.ToString());//Luu file Excel
                exApp.Visible = true;
                //exApp.Quit();//Thoát khỏi ứng dụng
            }
            else
            {
                MessageBox.Show("Không có danh sách hàng d? in");
            }
                
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            connv1.XuatExcelsqlcode("SELECT * FROM tblHang");
        }
    }
}
