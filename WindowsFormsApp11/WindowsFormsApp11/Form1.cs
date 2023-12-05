using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp11
{
    public partial class Form1 : Form
    {
        ProccessDatabase pd = new ProccessDatabase();
        string thuMucDich;
        string tenTepTin;
        public Form1()
        {
            InitializeComponent();
            this.KeyPreview = true;
        }

        // Dung tích
        private void comboBox1_DropDown(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            comboBox1.Items.Add("50");
            comboBox1.Items.Add("70");
            comboBox1.Items.Add("100");
            comboBox1.Items.Add("110");
            comboBox1.Items.Add("150");
        }

        // chọn ảnh xe
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.gif;*.bmp|All Files|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string duongDanAnh = openFileDialog.FileName;

                thuMucDich = @"D:\anhxe\";

                if (!Directory.Exists(thuMucDich))
                {
                    Directory.CreateDirectory(thuMucDich);
                }

                tenTepTin = Path.GetFileName(duongDanAnh);

                string duongDanDenThuMucDich = Path.Combine(thuMucDich, tenTepTin);

                if (File.Exists(duongDanDenThuMucDich))
                {
                    HienThiHinhAnh(duongDanDenThuMucDich);
                }
                else
                {
                    File.Copy(duongDanAnh, duongDanDenThuMucDich, true);
                }
            }
        }

        // thêm
        private bool KiemTraDL(string value)
        {
            return float.TryParse(value, out _);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if(textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "" ||
                comboBox1.Text == "" || textBox5.Text == "" || textBox6.Text == "")
            {
                MessageBox.Show("Hãy nhập đủ thông tin", "Thông báo", MessageBoxButtons.OKCancel, 
                    MessageBoxIcon.Error);
                textBox1.Focus();
                textBox2.Focus();
                textBox3.Focus();
                comboBox1.Focus();
                textBox5.Focus();
                textBox6.Focus();
            }
            else
            {
                DataTable dt = pd.DocBang($"Select * From tblXe Where SoKhung = '{textBox1.Text}'");
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Số khung này đã tồn tại, hãy nhập số khung khác", "Thông báo", MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Warning);
                    textBox1.Focus();
                }
                else
                {
                    if (!KiemTraDL(textBox3.Text))
                    {
                        MessageBox.Show("Mã màu phải là số", "Thông báo", MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Warning);
                        textBox3.Focus();
                    }

                    else
                    {
                        string imagePath = string.Empty; // Khởi tạo biến imagePath với giá trị chuỗi trống

                        if (!string.IsNullOrEmpty(tenTepTin))
                        {
                            imagePath = Path.Combine(thuMucDich, tenTepTin);
                        }
                        string sql = "Insert into tblXe Values(@soKhung, @soMay, @maMau, @dtxilanh, @hangXe, @tenXe, @anh)";
                        SqlParameter[] parameter = new SqlParameter[]
                        {
                            new SqlParameter("@soKhung", textBox1.Text),
                            new SqlParameter("@soMay", textBox2.Text),
                            new SqlParameter("@maMau", textBox3.Text),
                            new SqlParameter("@dtxilanh", comboBox1.Text),
                            new SqlParameter("@hangXe", textBox5.Text),
                            new SqlParameter("@tenXe", textBox6.Text),
                            new SqlParameter("@anh", string.IsNullOrEmpty(imagePath) ? (object)DBNull.Value : imagePath)
                        };
                        pd.CapNhatTS(sql, parameter);
                        MessageBox.Show("Thêm dữ liệu thành công", "Thông báo", MessageBoxButtons.OK);
                        dataGridView1.DataSource = pd.DocBang("Select * From tblXe");
                    }
                }
            }
        }

        // sửa
        private void button3_Click(object sender, EventArgs e)
        {
            string imagePath = string.Empty; 

            if (!string.IsNullOrEmpty(tenTepTin))
            {
                imagePath = Path.Combine(thuMucDich, tenTepTin);
            }
            string sql = "Update tblXe Set SoMay = @soMay, MaMau = @maMau, DungTichXiLanh = @dtxilanh, HangXe = @hangXe, " +
                "TenXe = @tenXe, Anh = @anh Where SoKhung = @soKhung";
            SqlParameter[] parameter = new SqlParameter[]
            {
                            new SqlParameter("@soKhung", textBox1.Text),
                            new SqlParameter("@soMay", textBox2.Text),
                            new SqlParameter("@maMau", textBox3.Text),
                            new SqlParameter("@dtxilanh", comboBox1.Text),
                            new SqlParameter("@hangXe", textBox5.Text),
                            new SqlParameter("@tenXe", textBox6.Text),
                            new SqlParameter("@anh", string.IsNullOrEmpty(imagePath) ? (object)DBNull.Value : imagePath)
            };
            pd.CapNhatTS(sql, parameter);
            MessageBox.Show("Cập nhật dữ liệu thành công", "Thông báo", MessageBoxButtons.OK);
            dataGridView1.DataSource = pd.DocBang("Select * From tblXe");
        }

        // xóa
        private void button4_Click(object sender, EventArgs e)
        {
            string sql = "Delete tblXe Where SoKhung = @soKhung";
            SqlParameter[] parameter = new SqlParameter[]
            {
                new SqlParameter("@soKhung", textBox1.Text)
            };
            pd.CapNhatTS(sql, parameter);
            MessageBox.Show("Xóa dữ liệu thành công", "Thông báo", MessageBoxButtons.OK);
            dataGridView1.DataSource = pd.DocBang("Select * From tblXe");
        }

        // làm mới
        private void reset()
        {
            textBox1.Text = textBox2.Text = textBox3.Text = comboBox1.Text = textBox5.Text = textBox6.Text = "";
            pictureBox1.Image = null;
            textBox1.Enabled = true;
            dataGridView1.DataSource = pd.DocBang("Select * From tblXe");
        }
        private void button5_Click(object sender, EventArgs e)
        {
           reset();
        }

        // xuất theo hãng
        private void button6_Click(object sender, EventArgs e)
        {
            if(textBox5.Text == "")
            {
                MessageBox.Show("Hãy nhập hãng xe", "Thông báo", MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Warning);
                textBox5.Focus();
            }
            else
            {
                var filteredRows = dataGridView1.Rows.Cast<DataGridViewRow>()
                             .Where(row => row.Cells.Count > 4 && row.Cells[4].Value != null && row.Cells[4].Value.ToString().Equals(textBox5.Text, StringComparison.OrdinalIgnoreCase))
                             .ToList();

                if (filteredRows.Count > 0)
                {
                    Excel.Application exApp = new Excel.Application();
                    Excel.Workbook exBook =
                    exApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    Excel.Worksheet exSheet =
                        (Excel.Worksheet)exBook.Worksheets[1];
                    Excel.Range tenvung = (Excel.Range)exSheet.Cells[1, 1];
                    tenvung.Font.Name = "Arial"; tenvung.Font.Size = 16;
                    tenvung.Font.Color = Color.Blue;
                    tenvung.Value = "Danh sách xe";
                    tenvung.HorizontalAlignment = HorizontalAlignment.Center;

                    exSheet.get_Range("A1: H1").Merge(true);
                    exSheet.get_Range("A2:H2").Font.Size = 14;
                    exSheet.get_Range("A2:H2").Font.Bold = true;
                    exSheet.get_Range("A2:H2").ColumnWidth = 20;
                    exSheet.get_Range("A2:H2").HorizontalAlignment = HorizontalAlignment.Center;
                    exSheet.get_Range("A2").Value = "STT";
                    exSheet.get_Range("B2").Value = "Số khung";
                    exSheet.get_Range("C2").Value = "Số máy";
                    exSheet.get_Range("D2").Value = "Màu";
                    exSheet.get_Range("E2").Value = "Dung tích xi-lanh";
                    exSheet.get_Range("F2").Value = "Hãng xe";
                    exSheet.get_Range("G2").Value = "Tên xe";
                    exSheet.get_Range("H2").Value = "Ảnh";

                    int k = filteredRows.Count;

                    exSheet.get_Range("A2:H" + (k + 2).ToString()).
                        Borders.LineStyle
                        = Excel.XlLineStyle.xlContinuous;
                    exSheet.get_Range("A2:H" + (k + 2).ToString()).
                        Borders.LineStyle
                        = Excel.XlBorderWeight.xlThin;
                    for (int i = 0; i < k; i++)
                    {
                        exSheet.get_Range("A" + (3 + i).ToString()).Value =
                            (i + 1).ToString();
                        exSheet.get_Range("B" + (3 + i).ToString()).Value =
                            dataGridView1.Rows[i].Cells[0].Value.ToString();
                        exSheet.get_Range("C" + (3 + i).ToString()).Value =
                            dataGridView1.Rows[i].Cells[1].Value.ToString();
                        exSheet.get_Range("D" + (3 + i).ToString()).Value =
                            dataGridView1.Rows[i].Cells[2].Value.ToString();
                        exSheet.get_Range("E" + (3 + i).ToString()).Value =
                            dataGridView1.Rows[i].Cells[3].Value.ToString();
                        exSheet.get_Range("F" + (3 + i).ToString()).Value =
                            dataGridView1.Rows[i].Cells[4].Value.ToString();
                        exSheet.get_Range("G" + (3 + i).ToString()).Value =
                            dataGridView1.Rows[i].Cells[5].Value.ToString();
                        exSheet.get_Range("H" + (3 + i).ToString()).Value =
                            dataGridView1.Rows[i].Cells[6].Value.ToString();
                    }
                    exBook.Activate();
                    SaveFileDialog svf = new SaveFileDialog();
                    svf.Title = "Chọn đường dẫn để lưu";
                    svf.ShowDialog();
                    string filename = svf.FileName;
                    if (filename == "")
                    {
                        MessageBox.Show("Bạn chưa đặt tên file");
                    }
                    exBook.SaveAs(filename);
                    exApp.Quit();
                }
                else
                {
                    MessageBox.Show("Hãng xe này không tồn tại!", "Thông báo",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        // tìm
        private void button7_Click(object sender, EventArgs e)
        {
            string tukhoa = textBox5.Text.Trim();
            string tukhoa_1 = textBox3.Text.Trim();

            if (string.IsNullOrEmpty(tukhoa) || string.IsNullOrEmpty(tukhoa_1))
            {
                MessageBox.Show("Vui lòng nhập thông tin cần tìm !", "Thông báo",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                return;
            }

            string sql = "Select * From tblXe Where HangXe Like @tukhoa And MaMau Like @tukhoa_1";

            DataTable dt = pd.LayDuLieu(sql,
                 string.IsNullOrEmpty(tukhoa) ? null : new SqlParameter("@tukhoa", "%" + tukhoa + "%"),
                 string.IsNullOrEmpty(tukhoa_1) ? null : new SqlParameter("@tukhoa_1", "%" + tukhoa_1 + "%"));

            //SqlParameter[] parameters = new SqlParameter[2];
            //parameters[0] = new SqlParameter("@tukhoa", SqlDbType.NVarChar);
            //parameters[0].Value = string.IsNullOrEmpty(tukhoa) ? (object)DBNull.Value : "%" + tukhoa + "%";

            //parameters[1] = new SqlParameter("@tukhoa_1", SqlDbType.NVarChar);
            //parameters[1].Value = string.IsNullOrEmpty(tukhoa_1) ? (object)DBNull.Value : "%" + tukhoa_1 + "%";

            //DataTable dt = pd.LayDuLieu(sql, parameters);

            if (dt.Rows.Count > 0)
            {
                dataGridView1.DataSource = dt;
            }
            else
            {
                MessageBox.Show("Không tồn tại thông tin cằn tìm !", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // thoát
        private void button8_Click(object sender, EventArgs e)
        {
            if(MessageBox.Show("Bạn có muốn thoát không ?", "Thông báo", 
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Application.Exit();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = pd.DocBang("Select * From tblXe");
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Enabled = false;

            textBox1.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            textBox2.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            textBox3.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            comboBox1.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            textBox5.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            textBox6.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            string imagePath = dataGridView1.CurrentRow.Cells[6].Value.ToString();

            HienThiHinhAnh(imagePath);
        }

        private void HienThiHinhAnh(string imagePath)
        {
            if (!string.IsNullOrEmpty(imagePath))
            {
                byte[] imageBytes = File.ReadAllBytes(imagePath);

                using (MemoryStream ms = new MemoryStream(imageBytes))
                {
                    pictureBox1.Image = Image.FromStream(ms);
                }
            }
            else
            {
                pictureBox1.Image = null;
            }
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Alt && e.KeyCode == Keys.M)
            {
                reset();
            }
        }
    }
}
