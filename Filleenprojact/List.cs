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
    public partial class List : Form
    {
        public List()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(@"C:\Users\chacr\Desktop\Filleenprojact\Filleenprojact\bin\Debug\data.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                int lastUsedRow = worksheet.Dimension.End.Row;
                for (int row = 1; row <= lastUsedRow; row++)
                {
                    string value = worksheet.Cells[row, 1].Value?.ToString();
                    if (!string.IsNullOrEmpty(value))
                    {
                        this.dataGridView1.Rows.Add(value);
                    }
                }

            }
        }
    }
}
