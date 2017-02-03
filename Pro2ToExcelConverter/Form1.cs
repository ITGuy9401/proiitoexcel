using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XWPF.UserModel;
using NPOI.SS.UserModel;
using NPOI.OpenXml4Net.OPC;
using NPOI.Util;


namespace Pro2ToExcelConverter
{
    public partial class Form1 : Form
    {

        private Dictionary<String, List<Point>> points = new Dictionary<string, List<Point>>();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void tableLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create an instance of the open file dialog box.
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            // Set filter options and filter index.
            openFileDialog1.Filter = "Pro II PLT Files|*.plt";
            openFileDialog1.FilterIndex = 1;

            openFileDialog1.Multiselect = false;

            // Call the ShowDialog method to show the dialog box.
            DialogResult userClickedOK = openFileDialog1.ShowDialog();

            // Process input if the user clicked OK.
            if (userClickedOK == DialogResult.OK)
            {
                // Open the selected file to read.
                System.IO.Stream fileStream = openFileDialog1.OpenFile();

                //Resets the lists
                points = new Dictionary<string, List<Point>>();

                fileNameTextBox.Text = openFileDialog1.FileName;

                using (System.IO.StreamReader reader = new System.IO.StreamReader(fileStream))
                {
                    String[] plt = reader.ReadToEnd().Split('\n');
                    for (int i = 0; i < plt.Length; i++)
                    {
                        String line = plt[i];
                        if (line.IndexOf("CURVE") != -1)
                        {
                            String title = line.Split('"')[1];
                            line = plt[++i];
                            line = stripAndTrim(plt[++i]);
                            List<Point> buffer = new List<Point>();
                            do
                            {
                                String[] split = line.Split(' ');
                                buffer.Add(new Point(XmlConvert.ToDouble(split[0]), XmlConvert.ToDouble(split[1])));
                                line = stripAndTrim(plt[++i]);
                            } while (line.IndexOf("MARKER") == -1);
                            points.Add(title, buffer);
                        }
                    }

                }
                fileStream.Close();
            }
        }

        private void convertButton_Click(object sender, EventArgs e)
        {
            if (points.Count > 0)
            {
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "EXCEL 2007+ File |*.xlsx";
                saveFileDialog1.Title = "Save the generated Excel File";
                saveFileDialog1.ShowDialog();

                // If the file name is not an empty string open it for saving.
                if (saveFileDialog1.FileName != "")
                {

                    // Converts the data to Excel format
                    IWorkbook wb = new XSSFWorkbook();
                    ISheet sheet = wb.CreateSheet();
                    IRow heading = sheet.CreateRow(0);
                    heading.CreateCell(0).SetCellValue("Pressure/Temp");

                    for (int i = 0; i < points.Count; i++)
                    {
                        points.ElementAt(i).Value.Sort(delegate (Point a, Point b)
                        {
                            return a.TemperatureOrPressure.CompareTo(b.TemperatureOrPressure);
                        });
                        heading.CreateCell(i + 1).SetCellValue(points.ElementAt(i).Key);
                    }


                    for (int i = 0; i < points.ElementAt(0).Value.Count; i++)
                    {
                        IRow row = sheet.CreateRow(i + 1);

                        ICell temp = row.CreateCell(0);
                        temp.SetCellType(CellType.Numeric);
                        temp.SetCellValue(points.ElementAt(0).Value.ElementAt(0).TemperatureOrPressure);

                        for (int j = 0; j < points.Count; j++)
                        {
                            ICell x = row.CreateCell(i+1);
                            x.SetCellType(CellType.Numeric);
                            x.SetCellValue(points.ElementAt(j).Value.ElementAt(i).Fraction);
                        }
                        
                    }

                    // Saves the file via a FileStream created by the OpenFile method.
                    System.IO.FileStream fs =
                       (System.IO.FileStream)saveFileDialog1.OpenFile();
                    wb.Write(fs);
                    try
                    {
                        fs.Close();
                    }
                    catch (Exception ex) { }
                    MessageBox.Show("File has been saved", "User Support");
                }
            } else
            {
                MessageBox.Show("You should first select a file", "User Support");
            }
        }

        private String stripAndTrim(String data)
        {
            return data.Trim();
        }
    }
}
