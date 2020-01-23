using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ProgmaticallyGenerateWordFiles
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnGenerateCV_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();

            word.Documents.Add();
            word.Visible = false;
            word.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
            Microsoft.Office.Interop.Word.Document doc = word.ActiveDocument;


            Microsoft.Office.Interop.Word.Paragraph paragraph = doc.Content.Paragraphs.Add(System.Reflection.Missing.Value);
            paragraph.Range.Text = "Kamal Ashraf";
            paragraph.Range.Font.Size = 22;
            paragraph.Range.Font.Bold = 1;
            paragraph.Range.Font.Name = "Times New Roman (Headings CS)";
            paragraph.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            paragraph.Range.InsertParagraphAfter();


            paragraph = doc.Content.Paragraphs.Add(System.Reflection.Missing.Value);
            paragraph.Range.Text = "Gujrat, Pakistan";
            paragraph.Range.Font.Size = 12;
            paragraph.Range.Font.Bold = 0;
            paragraph.Range.Font.Name = "Times New Roman (Headings CS)";
            paragraph.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            paragraph.Range.InsertParagraphAfter();





            paragraph = doc.Content.Paragraphs.Add(System.Reflection.Missing.Value);
            paragraph.Range.Text = "CNIC:\t\t\t\t31001-1234567-1";
            paragraph.Range.Font.Size = 12;
            paragraph.Range.Font.Bold = 0;
            paragraph.Range.Font.Name = "Times New Roman (Headings CS)";
            paragraph.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            paragraph.Range.InsertParagraphAfter();


            string textToFind = "CNIC:";
            Microsoft.Office.Interop.Word.Range range;
            Microsoft.Office.Interop.Word.Range textFormat;

            range = doc.Range();
            range.Find.Execute(textToFind);
            object start = range.Start;
            object end = range.Start + textToFind.Length + 1;
            textFormat = doc.Range(ref start, ref end);
            textFormat.Font.Bold = 1;


            paragraph = doc.Content.Paragraphs.Add(System.Reflection.Missing.Value);
            paragraph.Range.Text = "Date of birth:\t\t\t12th November 1992";
            paragraph.Range.Font.Size = 12;
            paragraph.Range.Font.Bold = 0;
            paragraph.Range.Font.Name = "Times New Roman (Headings CS)";
            paragraph.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            paragraph.Range.InsertParagraphAfter();

            textToFind = "Date of birth:";
            range = doc.Range();
            range.Find.Execute(textToFind);
            start = range.Start;
            end = range.Start + textToFind.Length + 1;
            textFormat = doc.Range(ref start, ref end);
            textFormat.Font.Bold = 1;

            paragraph = doc.Content.Paragraphs.Add(System.Reflection.Missing.Value);
            paragraph.Range.Text = "Email address:\t\t" + "myemail@gmail.com";
            paragraph.Range.Font.Size = 12;
            paragraph.Range.Font.Bold = 0;
            paragraph.Range.Font.Name = "Times New Roman (Headings CS)";
            paragraph.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            paragraph.Range.InsertParagraphAfter();

            textToFind = "Email address:";
            range = doc.Range();
            range.Find.Execute(textToFind);
            start = range.Start;
            end = range.Start + textToFind.Length + 1;
            textFormat = doc.Range(ref start, ref end);
            textFormat = doc.Range(ref start, ref end);
            textFormat.Font.Bold = 1;

            paragraph = doc.Content.Paragraphs.Add(System.Reflection.Missing.Value);
            paragraph.Range.Text = "Mobile number:\t\t+92-300-1234567";
            paragraph.Range.Font.Size = 12;
            paragraph.Range.Font.Bold = 0;
            paragraph.Range.Font.Name = "Times New Roman (Headings CS)";
            paragraph.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            paragraph.Range.InsertParagraphAfter();

            textToFind = "Mobile number:";
            range = doc.Range();
            range.Find.Execute(textToFind);
            start = range.Start;
            end = range.Start + textToFind.Length + 1;
            textFormat = doc.Range(ref start, ref end);
            textFormat = doc.Range(ref start, ref end);
            textFormat.Font.Bold = 1;

            paragraph = doc.Content.Paragraphs.Add(System.Reflection.Missing.Value);
            paragraph.Range.Text = "Landline number:\t\t+92-55-66777";
            paragraph.Range.Font.Size = 12;
            paragraph.Range.Font.Bold = 0;
            paragraph.Range.Font.Name = "Times New Roman (Headings CS)";
            paragraph.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            paragraph.Range.InsertParagraphAfter();

            textToFind = "Landline number:";
            range = doc.Range();
            range.Find.Execute(textToFind);
            start = range.Start;
            end = range.Start + textToFind.Length + 1;
            textFormat = doc.Range(ref start, ref end);
            textFormat = doc.Range(ref start, ref end);
            textFormat.Font.Bold = 1;


            paragraph = doc.Content.Paragraphs.Add(System.Reflection.Missing.Value);
            paragraph.Range.Text = "About me";
            paragraph.Range.Font.Size = 18;
            paragraph.Range.Font.Bold = 1;
            paragraph.Range.Font.Name = "Calibri";
            paragraph.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            paragraph.Shading.ForegroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorLightBlue;


            paragraph.Range.InsertParagraphAfter();

            paragraph = doc.Content.Paragraphs.Add(System.Reflection.Missing.Value);
            paragraph.Range.Text = "I am able to handle multiple tasks on a daily basis. I use a creative approach to problem solve. I am a dependable person who is great at time management. I am always energetic and eager to learn new skills. I am flexible in my working hours, being able to work evenings and weekends. I am hardworking and always the last to leave the office in the evening.";
            paragraph.Range.Font.Size = 12;
            paragraph.Range.Font.Bold = 0;
            paragraph.Range.Font.Name = "Times New Roman (Headings CS)";
            paragraph.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            paragraph.Shading.ForegroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorWhite;
            paragraph.Range.InsertParagraphAfter();


            paragraph = doc.Content.Paragraphs.Add(System.Reflection.Missing.Value);
            paragraph.Range.Text = "Qualification";
            paragraph.Range.Font.Size = 18;
            paragraph.Range.Font.Bold = 1;
            paragraph.Range.Font.Name = "Calibri";
            paragraph.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            paragraph.Shading.ForegroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorLightBlue;
            paragraph.Range.InsertParagraphAfter();


            List<string> qualificationList = new List<string>();

            qualificationList.Add("MPhill: 3.0/4.0 CGPA");
            qualificationList.Add("MSc: 70% marks");
            qualificationList.Add("BSc: 65% marks");
            qualificationList.Add("Inter: 70% marks");
            qualificationList.Add("Matriculation: 75% marks");
            
            object oEndOfDoc = "\\endofdoc";

            Microsoft.Office.Interop.Word.Range wrdRng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            paragraph.Shading.ForegroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorWhite;


            paragraph = doc.Content.Paragraphs.Add(System.Reflection.Missing.Value);
            paragraph.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            paragraph.Range.Font.Name = "Times New Roman (Headings CS)";

            paragraph.Range.ListFormat.ApplyBulletDefault();

            for (int i = 0; i < qualificationList.Count; i++)
            {
                string bulletItem = qualificationList[i];
                if (i < qualificationList.Count - 1)
                    bulletItem = bulletItem + "\n";
                paragraph.Range.Font.Bold = 0;
                paragraph.Range.InsertBefore(bulletItem);
            }


            paragraph = doc.Content.Paragraphs.Add(System.Reflection.Missing.Value);
            paragraph.Range.Text = "Experience";
            paragraph.Range.Font.Size = 18;
            paragraph.Range.Font.Bold = 1;
            paragraph.Range.Font.Name = "Calibri";
            paragraph.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            paragraph.Shading.ForegroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorLightBlue;
            paragraph.Range.InsertParagraphAfter();

            List<string> experienceList = new List<string>();

            experienceList.Add("Worked in Company A: 2014 to 2015");
            experienceList.Add("Worked in Company B: 2015 to 2016");
            experienceList.Add("Worked in Company C: 2016 to 2017");
            experienceList.Add("Worked in Company D: 2017 to 2018");
            experienceList.Add("Worked in Company E: 2018 to 2019");

            oEndOfDoc = "\\endofdoc";



            paragraph = doc.Content.Paragraphs.Add(System.Reflection.Missing.Value);
            paragraph.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            paragraph.Range.Font.Name = "Times New Roman (Headings CS)";
            paragraph.Shading.ForegroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorWhite;

            paragraph.Range.ListFormat.ApplyBulletDefault();

            for (int i = 0; i < experienceList.Count; i++)
            {
                string bulletItem = experienceList[i];
                if (i < experienceList.Count - 1)
                    bulletItem = bulletItem + "\n";
                paragraph.Range.Font.Bold = 0;
                paragraph.Range.InsertBefore(bulletItem);
            }

            paragraph = doc.Content.Paragraphs.Add(System.Reflection.Missing.Value);
            paragraph.Range.Text = "Interests";
            paragraph.Range.Font.Size = 18;
            paragraph.Range.Font.Bold = 1;
            paragraph.Range.Font.Name = "Calibri";
            paragraph.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            paragraph.Shading.ForegroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorLightBlue;
            paragraph.Range.InsertParagraphAfter();

            paragraph = doc.Content.Paragraphs.Add(System.Reflection.Missing.Value);
            paragraph.Range.Text = "Playing video games and cricket";
            paragraph.Range.Font.Size = 12;
            paragraph.Range.Font.Bold = 0;
            paragraph.Range.Font.Name = "Calibri";
            paragraph.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            paragraph.Shading.ForegroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorWhite;
            paragraph.Range.InsertParagraphAfter();


            string imagePath = System.AppDomain.CurrentDomain.BaseDirectory + "../../../person.png";

            Microsoft.Office.Interop.Word.InlineShape inlineShape = doc.InlineShapes.AddPicture(imagePath, Type.Missing, Type.Missing, Type.Missing);
            Microsoft.Office.Interop.Word.Shape shape = inlineShape.ConvertToShape();
            shape.HeightRelative = 10f;
            shape.WidthRelative = 18f;
            shape.Left = (float)Microsoft.Office.Interop.Word.WdShapePosition.wdShapeRight;
            shape.Top = 40F;
            shape.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapSquare;


            foreach (Microsoft.Office.Interop.Word.Section wordSection in doc.Sections)
            {
                Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);

                footerRange.Fields.Add(footerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages);
                Microsoft.Office.Interop.Word.Paragraph p4 = footerRange.Paragraphs.Add();
                p4.Range.Text = " of ";
                footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;

                footerRange.Fields.Add(footerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                Microsoft.Office.Interop.Word.Paragraph p1 = footerRange.Paragraphs.Add();
                p1.Range.Text = "Page ";
                footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;

                footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;
            }



            Microsoft.Win32.SaveFileDialog dialogBox = new Microsoft.Win32.SaveFileDialog();
            dialogBox.Title = "Choose destination to save file";
            dialogBox.DefaultExt = ".pdf";
            dialogBox.Filter = "Word documents (.docx)|*.docx|PDF documents (.pdf)|*.pdf";
            bool? result = dialogBox.ShowDialog();
            if (result == true)
            {
                string fileName = dialogBox.FileName;
                if (fileName.EndsWith(".docx"))
                {
                    doc.SaveAs(fileName);
                }
                else if (fileName.EndsWith(".pdf"))
                {
                    doc.ExportAsFixedFormat(fileName, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
                }
                MessageBox.Show("File \"" + fileName + "\"" + " saved.", "Success");
            }

            doc.Close(0);
            word.Quit();
            Marshal.ReleaseComObject(doc);
        }

    }
}
