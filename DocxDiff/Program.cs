using System.Collections;
using System.IO.Compression;
using System.Xml;

namespace DocxDiff
{
    internal class Program
    {
        static XmlDocument? getDocumentInArchive(string path, string search)
        {
            using (var file = File.OpenRead(path))
            {
                using (var zip = new ZipArchive(file, ZipArchiveMode.Read))
                {
                    foreach (var entry in zip.Entries)
                    {
                        if (entry.Name == search)
                        {
                            XmlDocument? document = new XmlDocument();
                            document.Load(entry.Open());
                            return document;
                        }
                    }
                }
            }

            return null;
        }

        static void Main(string[] args)
        {
            string originalPath = "C:\\Users\\User\\Desktop\\temp\\diff\\original.docx";
            string diffPath = "C:\\Users\\User\\Desktop\\temp\\diff\\modified.docx";

            XmlDocument? originalDocument = getDocumentInArchive(originalPath, "document.xml");
            XmlDocument? diffDocument = getDocumentInArchive(diffPath, "document.xml");

            if (originalDocument == null || diffDocument == null)
            {
                Console.WriteLine("Unable to open original or diff document.");
                return;
            }

            WordDocument doc = new WordDocument(diffDocument);
            foreach (var paragraph in doc.Paragraphs)
            {
                Console.WriteLine(paragraph.Text);
            }
        }
    }

    class WordParagraphRange
    {
        public bool IsBold { get; set; } = false;
        public bool IsItalic { get; set; } = false;
        public bool IsUnderlined { get; set; } = false;
        public string Color { get; set; } = "";

        public WordParagraphRange(XmlNode rangeNode)
        {
            // look into w:pPr
            // if w:i, then italic
            // if w:b, then bold
            // if w:u, then underlined
        }
    }

    class WordParagraphIndent { }

    class WordParagraphSpacing { }

    class WordParagraph
    {
        public WordParagraphIndent Indent { get; set; } = new WordParagraphIndent();
        public WordParagraphSpacing Spacing { get; set; } = new WordParagraphSpacing();
        public string Style { get; set; } = "";
        public string Text { get; set; } = "";
        public List<WordParagraphRange> Ranges { get; set; } = new List<WordParagraphRange>();

        public WordParagraph(XmlNode paragraphNode, XmlNamespaceManager nsmgr)
        {
            Style = paragraphNode.SelectSingleNode("w:pPr/w:pStyle", nsmgr)?.Attributes?["w:val"]?.Value ?? "";
            Text = paragraphNode.InnerText;

            XmlNodeList? rangeNodes = paragraphNode.SelectNodes("w:r", nsmgr);

            if (rangeNodes != null)
            {
                foreach (XmlNode rangeNode in rangeNodes)
                {
                    Ranges.Add(new WordParagraphRange(rangeNode));
                }
            }
        }
    }

    class WordDocument
    {
        public List<WordParagraph> Paragraphs { get; set; } = new List<WordParagraph>();

        public WordDocument(XmlDocument wordDocument)
        {
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(wordDocument.NameTable);
            nsmgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            XmlNodeList paragraphNodes = wordDocument.GetElementsByTagName("w:p");
            foreach (XmlNode paragraphNode in paragraphNodes)
            {
                Paragraphs.Add(new WordParagraph(paragraphNode, nsmgr));
            }
        }
    }
}