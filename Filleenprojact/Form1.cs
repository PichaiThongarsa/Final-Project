using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Filleenprojact
{
    public partial class Form1 : Form
    {
        public string name;
        public string author;
        public string category;
        public string page; 
        public string day;
        public string duration;
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                name = name1.Text;
                author = author1.Text;
                category = category1.Text;
                page = page1.Text;
                day = textBox1.Text;
                Book book = new Book(name, author, category, page, day);
                var message = $"เช่าสำเร็จ\nรายการของคุณคือ: {book.getName()}\nผู้เขียน: {book.getAuthor()}\nหมวดหมู่: {book.getCategory()}\nจำนวนหน้า: {book.getPage()} หน้า\nจำนวนวันในการเช่า: {book.getDay()} วัน";
                using (var package = new ExcelPackage(new FileInfo(@"C:\Users\chacr\Desktop\Filleenprojact\Filleenprojact\bin\Debug\data.xlsx")))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                    if (worksheet == null)
                    {
                        throw new Exception("Worksheet does not exist");
                    }
                    int lastUsedRow = worksheet.Dimension.End.Row;

                    ExcelRange cell = worksheet.Cells[lastUsedRow + 1, 1];
                    cell.Value = message;

                    package.Save();
                }
                MessageBox.Show(message);
                List list = new List();
                list.ShowDialog();
            }
            else
            {
                MessageBox.Show("กรุณากรอกวันจำนวนวันที่จะยืม!");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox3.Text != "")
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                name = name3.Text;
                author = author3.Text;
                category = category3.Text;
                duration = duration1.Text;
                day = textBox3 .Text;
                Movie movie = new Movie(name, author, category, duration, day);
                var message = $"เช่าสำเร็จ\nรายการของคุณคือ: {movie.getName()}\nผู้เขียน: {movie.getAuthor()}\nหมวดหมู่: {movie.getCategory()}\nความยาว: {movie.getDuration()}\nจำนวนวันในการเช่า: {movie.getDay()} วัน";
                using (var package = new ExcelPackage(new FileInfo(@"C:\Users\chacr\Desktop\Filleenprojact\Filleenprojact\bin\Debug\data.xlsx")))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                    if (worksheet == null)
                    {
                        throw new Exception("Worksheet does not exist");
                    }
                    int lastUsedRow = worksheet.Dimension.End.Row;

                    ExcelRange cell = worksheet.Cells[lastUsedRow + 1, 1];
                    cell.Value = message;

                    package.Save();
                }
                MessageBox.Show(message);
                List list = new List();
                list.ShowDialog();
            }
            else
            {
                MessageBox.Show("กรุณากรอกวันจำนวนวันที่จะยืม!");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox2.Text != "")
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                name = name2.Text;
                author = author2.Text;
                category = category2.Text;
                page = page2.Text;
                day = textBox2.Text;
                Book book = new Book(name, author, category, page, day);
                var message = $"เช่าสำเร็จ\nรายการของคุณคือ: {book.getName()}\nผู้เขียน: {book.getAuthor()}\nหมวดหมู่: {book.getCategory()}\nจำนวนหน้า: {book.getPage()} หน้า\nจำนวนวันในการเช่า: {book.getDay()} วัน";
                using (var package = new ExcelPackage(new FileInfo(@"C:\Users\chacr\Desktop\Filleenprojact\Filleenprojact\bin\Debug\data.xlsx")))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                    if (worksheet == null)
                    {
                        throw new Exception("Worksheet does not exist");
                    }
                    int lastUsedRow = worksheet.Dimension.End.Row;

                    ExcelRange cell = worksheet.Cells[lastUsedRow + 1, 1];
                    cell.Value = message;

                    package.Save();
                }
                MessageBox.Show(message);
                List list = new List();
                list.ShowDialog();
            }
            else
            { 
                MessageBox.Show("กรุณากรอกวันจำนวนวันที่จะยืม!");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox4.Text != "")
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                name = name4.Text;
                author = author4.Text;
                category = category4.Text;
                duration = duration2.Text;
                day = textBox4.Text;
                Movie movie = new Movie(name, author, category, duration, day);
                var message = $"เช่าสำเร็จ\nรายการของคุณคือ: {movie.getName()}\nผู้เขียน: {movie.getAuthor()}\nหมวดหมู่: {movie.getCategory()}\nความยาว: {movie.getDuration()}\nจำนวนวันในการเช่า: {movie.getDay()} วัน";
                using (var package = new ExcelPackage(new FileInfo(@"C:\Users\chacr\Desktop\Filleenprojact\Filleenprojact\bin\Debug\data.xlsx")))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                    if (worksheet == null)
                    {
                        throw new Exception("Worksheet does not exist");
                    }
                    int lastUsedRow = worksheet.Dimension.End.Row;

                    ExcelRange cell = worksheet.Cells[lastUsedRow + 1, 1];
                    cell.Value = message;

                    package.Save();
                }
                MessageBox.Show(message);
                List list = new List();
                list.ShowDialog();
            }
            else
            {
                MessageBox.Show("กรุณากรอกวันจำนวนวันที่จะยืม!");
            }
        }

        private void ประวตการเชาToolStripMenuItem_Click(object sender, EventArgs e)
        {
            List list = new List();
            list.ShowDialog();
        }
    }
}
