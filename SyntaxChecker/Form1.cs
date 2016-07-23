using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using word = DocumentFormat.OpenXml.Wordprocessing;
namespace SyntaxChecker
{
    public partial class Form1 : Form
    {
        //private Microsoft.Office.Interop.Word.Application oWord;
        public Form1()
        {
            InitializeComponent();
            //    oWord = new Microsoft.Office.Interop.Word.Application();
            //    oWord.DocumentBeforeClose += 
            //new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(
            //oWord_DocumentBeforeClose ); 

            string filepath = "d:\\venous flow.docx";
            //"d:\\Demo File (PDF to Word).docx";
            // using (WordprocessingDocument doc = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            // {

            //     MainDocumentPart mainDocumentPart = doc.AddMainDocumentPart();
            //     mainDocumentPart.Document = new Document();
            //     Body body = mainDocumentPart.Document.AppendChild(new Body());
            //     Paragraph para = body.AppendChild(new Paragraph());
            //     Run run = para.AppendChild(new Run());
            //     RunProperties runProperties = run.AppendChild(new RunProperties());
            //     FontSize fontSize = new FontSize();
            //     fontSize.Val = "40";
            //     RunFonts runFont = new RunFonts();
            //     runFont.Ascii = "Times New Roman";
            //     runProperties.Append(runFont);
            //     runProperties.AppendChild(fontSize);

            //     RunProperties rPr = new RunProperties(
            //new RunFonts()
            //{
            //    Ascii = "Arial"
            //});

            //     Run r = mainDocumentPart.Document.Descendants<Run>().First();
            //     r.PrependChild<RunProperties>(rPr);


            //     mainDocumentPart.Document.Save();
            // }







            //WordprocessingDocument wordprocessingDocument =
            //    WordprocessingDocument.Open(filepath, true);

            //// Assign a reference to the existing document body.
            //Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
            //SectionProperties sectionProps = new SectionProperties();
            //PageMargin pageMargin = new PageMargin() { Top = 1440, Right = (UInt32Value)1800U, Bottom = 1440, Left = (UInt32Value)1800U, Header = (UInt32Value)0U, Footer = (UInt32Value)0U, Gutter = (UInt32Value)0U };
            //sectionProps.Append(pageMargin);

            //// setting PageSize A4 and Portrait
            //PageSize pgsize = new PageSize();
            //pgsize.Width = 11909;
            //pgsize.Height = 16834;
            //pgsize.Orient = PageOrientationValues.Portrait;
            //sectionProps.Append(pgsize);

            //Run run = new Run();
            //RunProperties runProp = new RunProperties(); // Create run properties.
            //RunFonts runFont = new RunFonts();           // Create font
            //runFont.Ascii = "Times New Roman";                     // Specify font family

            //FontSize size = new FontSize();
            //size.Val = new StringValue("48");  // 48 half-point font size
            //runProp.Append(runFont);
            //runProp.Append(size);


            //run.PrependChild<RunProperties>(runProp);

            //wordprocessingDocument.MainDocumentPart.Document.Body.Append(sectionProps);
            //wordprocessingDocument.Close();





            const string documentRelationshipType =
         "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
            const string stylesRelationshipType =
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
            const string wordmlNamespace =
              "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace w = wordmlNamespace;

            XDocument xDoc = null;
            XDocument styleDoc = null;



            using (Package wdPackage = Package.Open(filepath, FileMode.Open, FileAccess.ReadWrite))
            {



                PackageRelationship docPackageRelationship =
                  wdPackage
                  .GetRelationshipsByType(documentRelationshipType)
                  .FirstOrDefault();
                if (docPackageRelationship != null)
                {
                    Uri documentUri =
                        PackUriHelper
                        .ResolvePartUri(
                           new Uri("/", UriKind.Relative),
                                 docPackageRelationship.TargetUri);
                    PackagePart documentPart =
                        wdPackage.GetPart(documentUri);

                    //  Load the document XML in the part into an XDocument instance.
                    xDoc = XDocument.Load(XmlReader.Create(documentPart.GetStream()));

                    //  Find the styles part. There will only be one.
                    PackageRelationship styleRelation =
                      documentPart.GetRelationshipsByType(stylesRelationshipType)
                      .FirstOrDefault();
                    if (styleRelation != null)
                    {
                        Uri styleUri = PackUriHelper.ResolvePartUri(documentUri, styleRelation.TargetUri);
                        PackagePart stylePart = wdPackage.GetPart(styleUri);

                        //  Load the style XML in the part into an XDocument instance.
                        styleDoc = XDocument.Load(XmlReader.Create(stylePart.GetStream()));
                    }
                }


                string defaultStyle =
           (string)(
               from style in styleDoc.Root.Elements(w + "style")
               where (string)style.Attribute(w + "type") == "paragraph" &&
                     (string)style.Attribute(w + "default") == "1"
               select style
           ).First().Attribute(w + "styleId");

                // Find all paragraphs in the document.
                var paragraphs1 =
                    from para in xDoc
                                 .Root
                                 .Element(w + "body")
                                 .Descendants(w + "p")

                    let styleNode = para
                                    .Elements(w + "pPr")
                                    .Elements(w + "pStyle")
                                    .FirstOrDefault()
                    select new
                    {
                        ParagraphNode = para,
                        StyleName = styleNode != null ?
                            (string)styleNode.Attribute(w + "val") :
                            defaultStyle
                    };


                // Retrieve the text of each paragraph.
                var paraWithText =
                    from para in paragraphs1
                    select new
                    {
                        ParagraphNode = para.ParagraphNode,
                        StyleName = para.StyleName,
                        Text = ParagraphText(para.ParagraphNode)

                    };

                foreach (var p in paraWithText)
                {
                    //Console.WriteLine("StyleName:{0} >{1}<", p.StyleName, p.Text);
                    // label1.Text = label1.Text + p.StyleName + ":" + p.Text + Environment.NewLine + Environment.NewLine;
                    label1.Text = label1.Text + p.Text + Environment.NewLine + Environment.NewLine;

                    textBox1.Text = textBox1.Text + GetDoubleSpace(p.Text);
                }

                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(wdPackage))
                {



                    //foreach (var item in wordDocument.MainDocumentPart.Document.Elements<Run>().ToList())
                    //{
                    //    Run r = item;
                    //    r.Remove();
                    //    r.RunProperties.Remove();

                    //}





                    foreach (Paragraph para in wordDocument.MainDocumentPart.Document.Body.Elements<Paragraph>().ToList())
                    {

                        foreach (var run in para.Elements<Run>())
                        {
                            foreach (var item in run.Elements<RunProperties>().ToList())
                            {
                                item.Remove();
                            }
                        }

                        foreach (var paraProp in para.Elements<ParagraphProperties>().ToList())
                        {
                            // remove existing Paragraph Properties
                            paraProp.Remove();
                        }

                    }


                    foreach (Paragraph para in wordDocument.MainDocumentPart.Document.Body.Elements<Paragraph>().ToList())
                    {

                        foreach (var run in para.Elements<Run>())
                        {
                            RunProperties runProp;
                            if (run.RunProperties != null)
                            {
                                runProp = (RunProperties)run.RunProperties.CloneNode(true); // Create run properties.
                            }
                            runProp = new RunProperties();
                            RunFonts runFont = new RunFonts();           // Create font
                            runFont.Ascii = "Times New Roman";                     // Specify font family

                            FontSize size = new FontSize();
                            size.Val = new StringValue("22");  // 48 half-point font size
                            runProp.Append(runFont);
                            runProp.Append(size);

                            runProp.Append(new EmbedRegularFont());

                            NoProof np = new NoProof();
                            np.Val = OnOffValue.FromBoolean(true);
                            runProp.Append(np);


                            Zoom zoom = new Zoom();
                            zoom.Percent = "100";
                            runProp.Append(zoom);

                            run.PrependChild<RunProperties>(runProp);
                        }
                    }

                    //MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                    //StyleDefinitionsPart stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    ////creation of a style
                    //Style style = new Style();
                    //style.StyleId = "MyHeading1"; //this is the ID of the style
                    //style.Append(new Name() { Val = "My Heading 1" }); //this is name
                    //// our style based on Normal style
                    //style.Append(new BasedOn() { Val = "Heading1" });
                    //// the next paragraph is Normal type
                    //style.Append(new NextParagraphStyle() { Val = "Normal" });

                    //stylePart.Styles = new Styles();
                    //stylePart.Styles.Append(style);
                    //stylePart.Styles.Save(); // we save the style part

                    foreach (Paragraph p in wordDocument.MainDocumentPart.Document.Body.Elements<Paragraph>())
                    {
                        // apply Paragraph Properties as  alignments
                        word.ParagraphProperties ParaProperties = new word.ParagraphProperties();
                        Justification justification1 = new Justification() { Val = JustificationValues.Left };

                        ParaProperties.ParagraphStyleId = new ParagraphStyleId() { Val = "MyHeading1" };

                        ParaProperties.PrependChild(justification1);
                        p.PrependChild(ParaProperties);

                    }




                    wordDocument.MainDocumentPart.Document.Save();





                }



            }



            //Environment.Exit(0);


        }



        public static string GetDoubleSpace(string plainText)
        {

            StringBuilder Result = new StringBuilder();
            char[] charArray = plainText.ToCharArray();
            char[] charPredef = { ';', ',', ':', '\'', '"', '?', '/', '\\', '!', '~', '%', '-', '-', '(', ')', '{', '}', '[', ']' };
            //  ;  ,   :   ‘   “   ?   /   \  !  ~  %  -  --  (   )  {   }  [   ]  

            char[] charDigits = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };
            char[] charRoman = { 'I', 'V', 'X', 'L', 'C', 'D', 'M' };
            //char[] charRoman = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };


            int index = 0;
            foreach (char c in plainText)
            {

                if (c == '.')
                {

                    // check last word /character of sentence before dot
                    if (Array.IndexOf(charPredef, charArray[index - 1]) > -1)
                    {

                    }
                    if (Array.IndexOf(charDigits, charArray[index - 1]) > -1)
                    {
                        Result.Append(c.ToString());
                    }
                    else
                    {
                        Result.Append(c.ToString() + ' ');

                    }


                }
                else
                {
                    Result.Append(c);
                }


                index++;
            }


            string res = FindDashInWords(Result.ToString());

            return res.ToString();
        }

        public static string FindDashInWords(string plainText)
        {
            int status = 0;

            //Data – Vision 		=> 		Data– Vision	(Remove first space)    1
            //Data–Vision			=>		Data–Vision                             2
            //Data –Vision   		=>		Data– Vision                            3

            string result = "";
            char[] charDash = { '-' };
            int index = 0;
            string strLeft = "";
            string strRight = "";
            string WholeString = "";

            foreach (char item in plainText)
            {
                if (item == charDash[0])
                {
                    plainText = GetWordLeftOrRight(plainText, index, 1, out status);
                    
                    if (status == 4)
                    {
                        //skip dash word with no space on both side of dash
                    }
                }

                index++;
            }

            result = plainText;
            return result;
        }
        public static List<KeyValuePair<string, string>> lst = new List<KeyValuePair<string, string>>();
        public static List<string> strLst = new List<string>();
        public static string GetWordLeftOrRight(string plainText, int indexOfDash, int isLeft, out int status)
        {
            char[] charArray = plainText.ToCharArray();

            bool isContainsLSpace = false;
            bool isContainsRSpace = false;


            StringBuilder sbResultLeft = new StringBuilder();
            StringBuilder sbResultRight = new StringBuilder();
            string strResultLeft = "";
            string strResultRight = "";
            string strResult="";

            //check if left and right side of dash doestn't have single space then neglect
            if (charArray[indexOfDash + 1] != ' ' && charArray[indexOfDash - 1] != ' ')
            {
                status = 4;
                return plainText;
            }

            //check if left and right side of dash  have single space 
            if (charArray[indexOfDash + 1] == ' ' && charArray[indexOfDash - 1] == ' ')
            {


                {
                    //for left traversing
                    for (int i = indexOfDash - 1; i != 0; i--)
                    {
                        if (charArray[i] != ' ')
                        {
                            sbResultLeft.Append(charArray[i]);
                        }
                        else if (sbResultLeft.Length > 1)
                        {
                            break;

                        }
                        else
                        {
                            isContainsLSpace = true;
                        }

                    }

                    strResultLeft = Reverse(sbResultLeft.ToString());
                }




                {
                    //for right traversing
                    for (int i = indexOfDash + 1; i != 999; i++)
                    {
                        if (charArray[i] != ' ')
                        {
                            sbResultRight.Append(charArray[i]);
                        }
                        else if (sbResultRight.Length > 1)
                        {

                            break;

                        }
                        else
                        {
                            isContainsRSpace = true;
                        }
                    }

                    strResultRight = (sbResultRight.ToString());
                }


                // strResult = strResultLeft + '-'+ ' ' + strResultRight;

                var regex = new Regex(Regex.Escape(strResultLeft + ' ' + '-' + ' ' + strResultRight));
                strResult = regex.Replace(plainText, strResultLeft + '-' + ' ' + strResultRight, 1);
                strLst.Add(strResult);
            }
            else
            {
                strResult = plainText;
            }

          

           

            if (isContainsLSpace && isContainsRSpace)
            {
                status = 3;
            }
            else if (isContainsLSpace)
            {
                status = 1;
            }
            else if (isContainsRSpace)
            {
                status = 2;
            }
            else
            {
                status = 0;
            }


            return strResult; 
        }

        public static string Reverse(string s)
        {
            char[] charArray = s.ToCharArray();
            Array.Reverse(charArray);
            return new string(charArray);
        }

        public static string ParagraphText(XElement e)
        {
            XNamespace w = e.Name.Namespace;
            return e
                   .Elements(w + "r")
                   .Elements(w + "t")
                   .StringConcatenate(element => (string)element);
        }


    }


    public static class LocalExtensions
    {
        public static string StringConcatenate(this IEnumerable<string> source)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string s in source)
                sb.Append(s);
            return sb.ToString();
        }

        public static string StringConcatenate<T>(this IEnumerable<T> source,
            Func<T, string> func)
        {
            StringBuilder sb = new StringBuilder();
            foreach (T item in source)
                sb.Append(func(item));
            return sb.ToString();
        }

        public static string StringConcatenate(this IEnumerable<string> source, string separator)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string s in source)
                sb.Append(s).Append(separator);
            return sb.ToString();
        }

        public static string StringConcatenate<T>(this IEnumerable<T> source,
            Func<T, string> func, string separator)
        {
            StringBuilder sb = new StringBuilder();
            foreach (T item in source)
                sb.Append(func(item)).Append(separator);
            return sb.ToString();
        }
    }
}
