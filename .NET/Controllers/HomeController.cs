using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using TOC.Models;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;

namespace TOC.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult CreateTOC()
        {
            using(WordDocument document = new WordDocument())
            {
                document.EnsureMinimal();
                document.LastSection.PageSetup.Margins.All = 72;
                WParagraph paragraph = document.LastParagraph;
                paragraph.AppendText("Essential DocIO - Table of Contents");
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                paragraph.ApplyStyle(BuiltinStyle.Heading4);
                paragraph = document.LastSection.AddParagraph() as WParagraph;
                paragraph = document.LastSection.AddParagraph() as WParagraph;

                TableOfContent TOC = paragraph.AppendTOC(1, 3);
                WSection section = document.LastSection;
                WParagraph newPara = section.AddParagraph() as WParagraph;
                newPara.AppendBreak(BreakType.PageBreak);
                AddHeading(section, BuiltinStyle.Heading1, "Document with built-in styles", "This is the built-in heading 1 style. This sample demonstrates the TOC insertion in a word document. Note that DocIO can insert TOC field in a word document. It can refresh or update TOC field by using UpdateTableOfContents method. MS Word refreshes the TOC field after insertion. Please update the field or press F9 key to refresh the TOC.");
                AddHeading(section, BuiltinStyle.Heading2, "Section 1", "This is the built-in heading 2 style. A document can contain any number of sections. Sections are used to apply same formatting for a group of paragraphs. You can insert sections by inserting section breaks.");
                AddHeading(section, BuiltinStyle.Heading3, "Paragraph 1", "This is the built-in heading 3 style. Each section contains any number of paragraphs. A paragraph is a set of statements that gives a meaning for the text.");
                AddHeading(section, BuiltinStyle.Heading3, "Paragraph 2", "This is the built-in heading 3 style. This demonstrates the paragraphs at the same level and style as that of the previous one. A paragraph can have any number formatting. This can be attained by formatting each text range in the paragraph.");

                section = document.AddSection() as WSection;
                section.PageSetup.Margins.All = 72;
                section.BreakCode = SectionBreakCode.NewPage;
                AddHeading(section, BuiltinStyle.Heading2, "Section 2", "This is the built-in heading 2 style. A document can contain any number of sections. Sections are used to apply same formatting for a group of paragraphs. You can insert sections by inserting section breaks.");
                AddHeading(section, BuiltinStyle.Heading3, "Paragraph 1", "This is the built-in heading 3 style. Each section contains any number of paragraphs. A paragraph is a set of statements that gives a meaning for the text.");
                AddHeading(section, BuiltinStyle.Heading3, "Paragraph 2", "This is the built-in heading 3 style. This demonstrates the paragraphs at the same level and style as that of the previous one. A paragraph can have any number formatting. This can be attained by formatting each text range in the paragraph.");
                document.UpdateTableOfContents();

                MemoryStream stream = new MemoryStream();
                document.Save(stream, FormatType.Docx);
                document.Close();
                return File(stream, "application/docx", "TOC.docx");
            }
        }

        private void AddHeading(WSection section, BuiltinStyle builtinStyle, string headingText, string paragraphText)
        {
            WParagraph paragraph = section.AddParagraph() as WParagraph;
            WTextRange text = paragraph.AppendText(headingText) as WTextRange;
            paragraph.ApplyStyle(builtinStyle);
            paragraph = section.AddParagraph() as WParagraph;
            paragraph.AppendText(paragraphText);
            section.AddParagraph();
        }

        private void AddHeading(WSection section, string styleName, string headingText, string paragraphText)
        {
            WParagraph paragraph = section.AddParagraph() as WParagraph;
            WTextRange text = paragraph.AppendText(headingText) as WTextRange;
            paragraph.ApplyStyle(styleName);
            paragraph = section.AddParagraph() as WParagraph;
            paragraph.AppendText(paragraphText);
            section.AddParagraph();
        }

        public IActionResult EditTOC()
        {
            using(WordDocument document = new WordDocument())
            {
                FileStream stream = new FileStream(Path.GetFullPath("Data/TOC.docx"), FileMode.Open, FileAccess.Read);
                document.Open(stream, FormatType.Docx);
                stream.Dispose();
                TableOfContent TOC = document.Sections[0].Body.Paragraphs[2].Items[0] as TableOfContent;
                TOC.LowerHeadingLevel = 1;
                TOC.UpperHeadingLevel= 2;
                TOC.IncludePageNumbers = false;
                
                document.UpdateTableOfContents();
                MemoryStream ms = new MemoryStream();
                document.Save(ms, FormatType.Docx);
                document.Close();
                return File(ms, "application/docx", "TOC-edited.docx");
            }
        }

        public IActionResult CustomStyleTOC()
        {
            using(WordDocument document = new WordDocument())
            {
                document.EnsureMinimal();
                document.LastSection.PageSetup.Margins.All = 72;
                WParagraph para = document.LastParagraph;
                para.AppendText("Essential DocIO - Table of Contents");
                para.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                para.ApplyStyle(BuiltinStyle.Heading4);
                para = document.LastSection.AddParagraph() as WParagraph;
                WParagraphStyle pStyle1 = (WParagraphStyle)document.AddParagraphStyle("MyStyle1");
                pStyle1.CharacterFormat.FontSize = 18f;
                WParagraphStyle pStyle2 = (WParagraphStyle)document.AddParagraphStyle("MyStyle2");
                pStyle2.CharacterFormat.FontSize = 16f;
                WParagraphStyle pStyle3 = (WParagraphStyle)document.AddParagraphStyle("MyStyle3");
                pStyle3.CharacterFormat.FontSize = 14f;
                para = document.LastSection.AddParagraph() as WParagraph;

                TableOfContent TOC = para.AppendTOC(1, 3);
                TOC.UseHeadingStyles = false;
                TOC.SetTOCLevelStyle(1, "MyStyle1");
                TOC.SetTOCLevelStyle(2, "MyStyle2");
                TOC.SetTOCLevelStyle(3, "MyStyle3");
                WSection section = document.LastSection;
                WParagraph newPara = section.AddParagraph() as WParagraph;
                newPara.AppendBreak(BreakType.PageBreak);

                AddHeading(section, "MyStyle1", "Document with custom styles", "This is the 1st custom style. This sample demonstrates the TOC insertion in a Word document. Note that DocIO can insert TOC fields in a Word document. It can refresh or update the TOC field by using the UpdateTableOfContents method. MS Word refreshes the TOC field after insertion. Please update the field or press F9 to refresh the TOC.");
                AddHeading(section, "MyStyle2", "Section 1", "This is the 2nd custom style. A document can contain any number of sections. Sections are used to apply the same formatting to a group of paragraphs. You can insert sections by inserting section breaks.");
                AddHeading(section, "MyStyle3", "Paragraph 1", "This is the 3rd custom style. Each section contains any number of paragraphs. A paragraph is a set of statements that gives meaning to the text.");
                AddHeading(section, "MyStyle3", "Paragraph 2", "This is the 3rd custom style. This demonstrates the paragraphs at the same level and with the same style as the previous one. A paragraph can have any kind of formatting. This can be attained by formatting each text range in the paragraph.");
                AddHeading(section, "Normal", "Paragraph with normal", "This is the paragraph with the Normal style. This demonstrates the paragraph with outline level 4 and the Normal style. This can be attained by formatting the outline level of the paragraph.");

                section = document.AddSection() as WSection;
                section.PageSetup.Margins.All = 72;
                section.BreakCode = SectionBreakCode.NewPage;
                AddHeading(section, "MyStyle2", "Section 2", "This is the 2nd custom style. A document can contain any number of sections. Sections are used to apply the same formatting to a group of paragraphs. You can insert sections by inserting section breaks.");
                AddHeading(section, "MyStyle3", "Paragraph 1", "This is the 3rd custom style. Each section contains any number of paragraphs. A paragraph is a set of statements that gives meaning to the text.");
                AddHeading(section, "MyStyle3", "Paragraph 2", "This is the 3rd custom style. This demonstrates the paragraphs at the same level and with the same style as the previous one. A paragraph can have any kind of formatting. This can be attained by formatting each text range in the paragraph.");

                document.UpdateTableOfContents();
                MemoryStream ms = new MemoryStream();
                document.Save(ms, FormatType.Docx);
                document.Close();
                return File(ms, "application/docx", "TOC-customstyle.docx");
            }
        }

        public IActionResult CustomTOCEntries()
        {
            using (WordDocument document = new WordDocument())
            {
                document.EnsureMinimal();
                document.LastSection.PageSetup.Margins.All = 72;
                WParagraph paragraph = document.LastParagraph;
                paragraph.AppendText("Essential DocIO - Table of Contents");
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                paragraph.ApplyStyle(BuiltinStyle.Heading4);
                paragraph = document.LastSection.AddParagraph() as WParagraph;
                paragraph = document.LastSection.AddParagraph() as WParagraph;

                TableOfContent TOC = paragraph.AppendTOC(1, 3);
                WSection section = document.LastSection;
                WParagraph newPara = section.AddParagraph() as WParagraph;
                newPara.AppendBreak(BreakType.PageBreak);
                AddHeading(section, BuiltinStyle.Heading1, "Document with built-in styles", "This is the built-in heading 1 style. This sample demonstrates the TOC insertion in a word document. Note that DocIO can insert TOC field in a word document. It can refresh or update TOC field by using UpdateTableOfContents method. MS Word refreshes the TOC field after insertion. Please update the field or press F9 key to refresh the TOC.");
                AddHeading(section, BuiltinStyle.Heading2, "Section 1", "This is the built-in heading 2 style. A document can contain any number of sections. Sections are used to apply same formatting for a group of paragraphs. You can insert sections by inserting section breaks.");
                AddHeading(section, BuiltinStyle.Heading3, "Paragraph 1", "This is the built-in heading 3 style. Each section contains any number of paragraphs. A paragraph is a set of statements that gives a meaning for the text.");
                AddHeading(section, BuiltinStyle.Heading3, "Paragraph 2", "This is the built-in heading 3 style. This demonstrates the paragraphs at the same level and style as that of the previous one. A paragraph can have any number formatting. This can be attained by formatting each text range in the paragraph.");

                section = document.AddSection() as WSection;
                section.PageSetup.Margins.All = 72;
                section.BreakCode = SectionBreakCode.NewPage;
                AddHeading(section, BuiltinStyle.Heading2, "Section 2", "This is the built-in heading 2 style. A document can contain any number of sections. Sections are used to apply same formatting for a group of paragraphs. You can insert sections by inserting section breaks.");
                AddHeading(section, BuiltinStyle.Heading3, "Paragraph 1", "This is the built-in heading 3 style. Each section contains any number of paragraphs. A paragraph is a set of statements that gives a meaning for the text.");
                AddHeading(section, BuiltinStyle.Heading3, "Paragraph 2", "This is the built-in heading 3 style. This demonstrates the paragraphs at the same level and style as that of the previous one. A paragraph can have any number formatting. This can be attained by formatting each text range in the paragraph.");

                IWParagraphStyle TOC1Style = document.AddParagraphStyle("TOC 1");
                TOC1Style.CharacterFormat.FontName = "Calibri";
                TOC1Style.CharacterFormat.FontSize = 14;
                TOC1Style.CharacterFormat.Bold = true;
                TOC1Style.CharacterFormat.Italic = true;
                TOC1Style.ParagraphFormat.AfterSpacing = 8;

                IWParagraphStyle TOC2style = document.AddParagraphStyle("TOC 2");
                TOC2style.CharacterFormat.FontName = "Calibri";
                TOC2style.CharacterFormat.FontSize = 12;
                TOC2style.ParagraphFormat.AfterSpacing = 5;
                TOC2style.CharacterFormat.Bold = true;
                TOC2style.ParagraphFormat.LeftIndent = 11;

                IWParagraphStyle TOC3style = document.AddParagraphStyle("TOC 3"); ;
                TOC3style.CharacterFormat.FontName = "Calibri";
                TOC3style.CharacterFormat.FontSize = 12;
                TOC3style.ParagraphFormat.AfterSpacing = 3;
                TOC3style.CharacterFormat.Italic = true;
                TOC3style.ParagraphFormat.LeftIndent = 22;

                document.UpdateTableOfContents();

                MemoryStream stream = new MemoryStream();
                document.Save(stream, FormatType.Docx);
                document.Close();
                return File(stream, "application/docx", "TOC-customentries.docx");
            }
        }
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}