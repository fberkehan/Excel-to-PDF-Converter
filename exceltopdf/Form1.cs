using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using ExcelDataReader;
using iTextSharp.text;
using iTextSharp.text.pdf;



namespace exceltopdf
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void dosyaSecButton_Click_1(object sender, EventArgs e)
        {System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx", ValidateNames = true })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    FileStream fs = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read);

                    IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(fs);

                    DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                    reader.Close();

                    DataTable dt = result.Tables[0];

                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string folderPath = Path.Combine(desktopPath, "EXCEL2PDF");
                    Directory.CreateDirectory(folderPath);

                    string pdfFileName = Path.ChangeExtension(ofd.SafeFileName, ".pdf");
                    string pdfPath = Path.Combine(folderPath, pdfFileName);


                    using (FileStream stream = new FileStream(pdfPath, FileMode.Create))
                    {
                        Document pdfDoc = new Document(PageSize.A4, 10f, 10f, 10f, 0f);
                        PdfWriter.GetInstance(pdfDoc, stream);

                        pdfDoc.Open();
                        PdfPTable pdfTable = new PdfPTable(dt.Columns.Count);
                        pdfTable.DefaultCell.Padding = 3;
                        pdfTable.WidthPercentage = 100;
                        pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;

                        foreach (DataColumn column in dt.Columns)
                        {
                            PdfPCell cell = new PdfPCell(new Phrase(column.ColumnName));
                            pdfTable.AddCell(cell);
                        }

                        foreach (DataRow row in dt.Rows)
                        {
                            foreach (object cell in row.ItemArray)
                            {
                                pdfTable.AddCell(new Phrase(cell.ToString()));
                            }
                        }

                        pdfDoc.Add(pdfTable);
                        pdfDoc.Close();
                        stream.Close();
                        MessageBox.Show("Ýþlem baþarýyla tamamlandý!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }
                }
            }
        }

   
    }
}
