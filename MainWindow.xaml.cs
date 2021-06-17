using System;
using System.Collections.Generic;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;


namespace Customer_integ
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>

    public partial class MainWindow : System.Windows.Window
    {
        string[] mainstrings = new string[200];
        int[] isprinting = new int[200];
        int[] excludearr = new int[200];
        public MainWindow()
        {
            InitializeComponent();
            for (int i = 0; i < excludearr.Length; i++)
            {
                excludearr[i] = 1;
            }
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open("C:/Users/VIJAY SIDHARTH/Desktop/Copy of appa-cust-editable.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            range = xlWorkSheet.UsedRange;
            string str;
            for (int i = 2; i < 98; i++)
            {
                str=(string)(range.Cells[i, 2] as Excel.Range).Value2 + '\n' +
                    (string)(range.Cells[i, 6] as Excel.Range).Value2 + '\n' +
                    (string)(range.Cells[i, 7] as Excel.Range).Value2 + '\n' +
                    (string)(range.Cells[i, 8] as Excel.Range).Value2 + '-' +
                    (string)Convert.ToString((range.Cells[i, 9] as Excel.Range).Value2) + '\n' +
                    "Ph No : " + "+" + (string)Convert.ToString((range.Cells[i, 4] as Excel.Range).Value2);
                mainstrings[i] = str;
            }

        }
        int cc = 2;
        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            Word.Application Wapp;
            Word.Document WDoc;
            object misValue = System.Reflection.Missing.Value;
            
            //Excel Arean
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open("C:/Users/VIJAY SIDHARTH/Desktop/Copy of appa-cust-editable.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            //Word Arena
            Wapp = new Word.Application();
            WDoc = new Word.Document();
            Wapp.ShowAnimation = true;
            Wapp.Visible = true;
            object missing = System.Reflection.Missing.Value;
            WDoc = Wapp.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            foreach(Microsoft.Office.Interop.Word.Section section in WDoc.Sections)
            {
                Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                headerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                headerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlue;
                
                //headerRange.Font.Size = 30;
               // headerRange.Font.Bold = 1;
                //headerRange.Text = "Customers Address";
            }
            foreach (Microsoft.Office.Interop.Word.Section section in WDoc.Sections)
            {
                Microsoft.Office.Interop.Word.Range headerRange = section.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                headerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
                headerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlack;
                headerRange.Font.Bold = 1;
                headerRange.Font.Size = 10;

                headerRange.Text = "FROM:  SRINIVASAN R, MUTUAL FUNDS AND INSURANCE CONSULTANT," + '\n' +
"4 - 1 - 395, ARASU NAGAR (BEHIND NO: 1 RETREADING CO)," + '\n' + "P.C.PATTY, THENI – 625531.        Ph : 9894883455, 8124999281";

            }
            range = xlWorkSheet.UsedRange;
            
            WDoc.Content.SetRange(100,100);
            string str =".";
            Word.WdPaperSize HJ = (Word.WdPaperSize)22;
            WDoc.PageSetup.PaperSize = HJ;
            WDoc.PageSetup.TopMargin = 5;
            WDoc.ConvertNumbersToText();
            WDoc.Content.PageSetup.PaperSize = HJ;
            List<Microsoft.Office.Interop.Word.Paragraph> paralist = new List<Microsoft.Office.Interop.Word.Paragraph>(20);
            Word.WdTextOrientation a = (Word.WdTextOrientation)0;
            //WDoc.PageSetup.PageHeight = (float)11.5;
            //WDoc.PageSetup.PageWidth = (float)27.5;
            
            for (int i = 2; i < 98;i++)
            {
                if (excludearr[i] == 0)
                {

                }
                else
                {
                    str = str + '\n' + '\n' + '\n' + '\n' + "To" + '\n';
                    str = str + (string)(range.Cells[i, 2] as Excel.Range).Value2 + '\n' +
                        (string)(range.Cells[i, 6] as Excel.Range).Value2 + '\n' +
                        (string)(range.Cells[i, 7] as Excel.Range).Value2 + '\n' +
                        (string)(range.Cells[i, 8] as Excel.Range).Value2 + '-' +
                        (string)Convert.ToString((range.Cells[i, 9] as Excel.Range).Value2) + '\n' +
                        "Ph No : " + (string)Convert.ToString((range.Cells[i, 4] as Excel.Range).Value2) + '\n' + '\n';
                    WDoc.Content.Orientation = a;
                    WDoc.Content.Font.Bold = 1;
                    WDoc.Content.Font.Size = 12;
                    /*paralist[i] = WDoc.Content.Paragraphs.Add(ref missing);
                    object styleHeading1 = "Heading 1";
                    paralist[i].Range.set_Style(ref styleHeading1);
                    paralist[i].Range.Text = str;
                    str = (string)(range.Cells[i, 3] as Excel.Range).Value2;
                    paralist[i].Range.Text = str;
                    WDoc.Content.Text = str;
                    paralist[i].Range.InsertParagraphAfter();*/
                }
            }
            WDoc.Content.Text = str;
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        System.Windows.Media.Brush gh = System.Windows.Media.Brushes.Blue;
        System.Windows.Media.Brush bh = System.Windows.Media.Brushes.Aqua;
        private void next_Click_1(object sender, RoutedEventArgs e)
        {
            cc++;
            if (excludearr[cc] == 0)
            {
                
                textBlock.Background = gh;
            }
            else
            {
                textBlock.Background = bh;
            }
            if(cc >= mainstrings.Length)
            {
                cc = mainstrings.Length-1;
            }
            textBlock.Text = mainstrings[cc];
        }

        private void prev_Click(object sender, RoutedEventArgs e)
        {
            cc--;
            if (excludearr[cc] == 0)
            {

                textBlock.Background = gh;
            }
            else
            {
                textBlock.Background = bh;
            }
            
            if (cc < 2)
            {
                cc = 2;
            }
            textBlock.Text = mainstrings[cc];
        }

        private void button1_Click_1(object sender, RoutedEventArgs e)
        {
            excludearr[cc] = 0;
            textBlock.Background = gh;
        }

        private void button2_Click(object sender, RoutedEventArgs e)
        {
            excludearr[cc] = 1;
            textBlock.Background = bh;
        }
    }

}
