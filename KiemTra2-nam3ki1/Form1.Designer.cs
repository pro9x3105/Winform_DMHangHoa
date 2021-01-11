namespace KiemTra2_nam3ki1
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnThoat = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.btnThem = new System.Windows.Forms.Button();
            this.dgvHienThi = new System.Windows.Forms.DataGridView();
            this.btnSua = new System.Windows.Forms.Button();
            this.btnXoa = new System.Windows.Forms.Button();
            this.btnExcel = new System.Windows.Forms.Button();
            this.txbMaHang = new System.Windows.Forms.TextBox();
            this.txbTenHang = new System.Windows.Forms.TextBox();
            this.txbSoLuong = new System.Windows.Forms.TextBox();
            this.txbGiaNhap = new System.Windows.Forms.TextBox();
            this.txbGiaBan = new System.Windows.Forms.TextBox();
            this.cbChatLieu = new System.Windows.Forms.ComboBox();
            this.saveExcel = new System.Windows.Forms.SaveFileDialog();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvHienThi)).BeginInit();
            this.SuspendLayout();
            // 
            // btnThoat
            // 
            this.btnThoat.Location = new System.Drawing.Point(514, 373);
            this.btnThoat.Name = "btnThoat";
            this.btnThoat.Size = new System.Drawing.Size(75, 23);
            this.btnThoat.TabIndex = 0;
            this.btnThoat.Text = "Thoát";
            this.btnThoat.UseVisualStyleBackColor = true;
            this.btnThoat.Click += new System.EventHandler(this.BtnThoat_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(324, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(208, 26);
            this.label1.TabIndex = 1;
            this.label1.Text = "Danh mục hàng hóa";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(56, 57);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(49, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Mã hàng";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(56, 98);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "Tên hàng";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(280, 57);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(48, 13);
            this.label4.TabIndex = 4;
            this.label4.Text = "Chất liệu";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(280, 98);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(49, 13);
            this.label5.TabIndex = 5;
            this.label5.Text = "Số lượng";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(511, 57);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(50, 13);
            this.label6.TabIndex = 6;
            this.label6.Text = "Giá nhập";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(511, 98);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(44, 13);
            this.label7.TabIndex = 7;
            this.label7.Text = "Giá bán";
            // 
            // btnThem
            // 
            this.btnThem.Location = new System.Drawing.Point(47, 373);
            this.btnThem.Name = "btnThem";
            this.btnThem.Size = new System.Drawing.Size(75, 23);
            this.btnThem.TabIndex = 8;
            this.btnThem.Text = "Thêm";
            this.btnThem.UseVisualStyleBackColor = true;
            this.btnThem.Click += new System.EventHandler(this.BtnThem_Click);
            // 
            // dgvHienThi
            // 
            this.dgvHienThi.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvHienThi.Location = new System.Drawing.Point(12, 151);
            this.dgvHienThi.Name = "dgvHienThi";
            this.dgvHienThi.Size = new System.Drawing.Size(733, 198);
            this.dgvHienThi.TabIndex = 9;
            this.dgvHienThi.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DgvHienThi_CellClick);
            // 
            // btnSua
            // 
            this.btnSua.Location = new System.Drawing.Point(172, 373);
            this.btnSua.Name = "btnSua";
            this.btnSua.Size = new System.Drawing.Size(75, 23);
            this.btnSua.TabIndex = 10;
            this.btnSua.Text = "Sửa";
            this.btnSua.UseVisualStyleBackColor = true;
            this.btnSua.Click += new System.EventHandler(this.BtnSua_Click);
            // 
            // btnXoa
            // 
            this.btnXoa.Location = new System.Drawing.Point(283, 373);
            this.btnXoa.Name = "btnXoa";
            this.btnXoa.Size = new System.Drawing.Size(75, 23);
            this.btnXoa.TabIndex = 11;
            this.btnXoa.Text = "Xóa";
            this.btnXoa.UseVisualStyleBackColor = true;
            this.btnXoa.Click += new System.EventHandler(this.BtnXoa_Click);
            // 
            // btnExcel
            // 
            this.btnExcel.Location = new System.Drawing.Point(394, 373);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(75, 23);
            this.btnExcel.TabIndex = 12;
            this.btnExcel.Text = "Xuất Excel";
            this.btnExcel.UseVisualStyleBackColor = true;
            this.btnExcel.Click += new System.EventHandler(this.BtnExcel_Click);
            // 
            // txbMaHang
            // 
            this.txbMaHang.Location = new System.Drawing.Point(110, 54);
            this.txbMaHang.Name = "txbMaHang";
            this.txbMaHang.Size = new System.Drawing.Size(100, 20);
            this.txbMaHang.TabIndex = 13;
            // 
            // txbTenHang
            // 
            this.txbTenHang.Location = new System.Drawing.Point(110, 91);
            this.txbTenHang.Multiline = true;
            this.txbTenHang.Name = "txbTenHang";
            this.txbTenHang.Size = new System.Drawing.Size(100, 45);
            this.txbTenHang.TabIndex = 14;
            // 
            // txbSoLuong
            // 
            this.txbSoLuong.Location = new System.Drawing.Point(344, 91);
            this.txbSoLuong.Name = "txbSoLuong";
            this.txbSoLuong.Size = new System.Drawing.Size(100, 20);
            this.txbSoLuong.TabIndex = 16;
            this.txbSoLuong.TextChanged += new System.EventHandler(this.TxbSoLuong_TextChanged);
            this.txbSoLuong.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxbSoLuong_KeyPress);
            // 
            // txbGiaNhap
            // 
            this.txbGiaNhap.Location = new System.Drawing.Point(576, 54);
            this.txbGiaNhap.Name = "txbGiaNhap";
            this.txbGiaNhap.Size = new System.Drawing.Size(100, 20);
            this.txbGiaNhap.TabIndex = 17;
            this.txbGiaNhap.TextChanged += new System.EventHandler(this.TxbGiaNhap_TextChanged);
            this.txbGiaNhap.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxbGiaNhap_KeyPress);
            // 
            // txbGiaBan
            // 
            this.txbGiaBan.Location = new System.Drawing.Point(576, 91);
            this.txbGiaBan.Name = "txbGiaBan";
            this.txbGiaBan.Size = new System.Drawing.Size(100, 20);
            this.txbGiaBan.TabIndex = 18;
            this.txbGiaBan.TextChanged += new System.EventHandler(this.TxbGiaBan_TextChanged);
            this.txbGiaBan.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxbGiaBan_KeyPress);
            // 
            // cbChatLieu
            // 
            this.cbChatLieu.FormattingEnabled = true;
            this.cbChatLieu.Location = new System.Drawing.Point(344, 53);
            this.cbChatLieu.Name = "cbChatLieu";
            this.cbChatLieu.Size = new System.Drawing.Size(100, 21);
            this.cbChatLieu.TabIndex = 19;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(650, 373);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 20;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.cbChatLieu);
            this.Controls.Add(this.txbGiaBan);
            this.Controls.Add(this.txbGiaNhap);
            this.Controls.Add(this.txbSoLuong);
            this.Controls.Add(this.txbTenHang);
            this.Controls.Add(this.txbMaHang);
            this.Controls.Add(this.btnExcel);
            this.Controls.Add(this.btnXoa);
            this.Controls.Add(this.btnSua);
            this.Controls.Add(this.dgvHienThi);
            this.Controls.Add(this.btnThem);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnThoat);
            this.Name = "Form1";
            this.Text = "Danh mục hàng hóa";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvHienThi)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnThoat;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button btnThem;
        private System.Windows.Forms.DataGridView dgvHienThi;
        private System.Windows.Forms.Button btnSua;
        private System.Windows.Forms.Button btnXoa;
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.TextBox txbMaHang;
        private System.Windows.Forms.TextBox txbTenHang;
        private System.Windows.Forms.TextBox txbSoLuong;
        private System.Windows.Forms.TextBox txbGiaNhap;
        private System.Windows.Forms.TextBox txbGiaBan;
        private System.Windows.Forms.ComboBox cbChatLieu;
        private System.Windows.Forms.SaveFileDialog saveExcel;
        private System.Windows.Forms.Button button1;
    }
}

