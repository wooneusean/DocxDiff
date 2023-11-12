using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace DocxDiff
{
    internal class Program
    {

        private static void Main(string[] args)
        {
            string originalPath = @"C:\Users\User\Desktop\temp\diff\original.docx";
            string modifiedPath = @"C:\Users\User\Desktop\temp\diff\modified.docx";
            string outputPath = @"C:\Users\User\Desktop\temp\diff\output.docx";

            WordDocument originalDoc = new WordDocument(originalPath);
            WordDocument modifiedDoc = new WordDocument(modifiedPath);

            DiffMatchPatch.diff_match_patch dmp = new DiffMatchPatch.diff_match_patch();
            var diffs = dmp.diff_main(originalDoc.Text, modifiedDoc.Text);

            foreach (var diff in diffs)
            {
                modifiedDoc.Paragraphs.Insert(0, new WordParagraph
                {
                    Ranges = new List<WordParagraphRange>
                    {
                        new WordParagraphRange {
                            Text = $"[{diff.operation}] {diff.text}", 
                            Color = diff.operation == DiffMatchPatch.Operation.DELETE ? "FF0000" : diff.operation == DiffMatchPatch.Operation.INSERT ? "00FF00" : "000000",
                            IsBold = true,
                        }
                    }
                });
            }

            modifiedDoc.SaveTo(outputPath);
        }
    }

    internal class WordParagraphRange
    {
        public bool IsBold { get; set; } = false;
        public bool IsItalic { get; set; } = false;
        public bool IsUnderlined { get; set; } = false;
        public string Color { get; set; } = "";
        public string Text { get; set; } = "";
        public string Highlight { get; set; } = "";
        public bool PreserveWhitespace { get; set; }

        public WordParagraphRange() { }

        public WordParagraphRange(XmlNode rangeNode, XmlNamespaceManager nsmgr)
        {
            Text = rangeNode.InnerText;

            PreserveWhitespace = rangeNode.SelectSingleNode("w:t", nsmgr)?.Attributes?["xml:space"]?.Value != null;

            XmlNodeList? rPrChildList = rangeNode.SelectSingleNode("w:rPr", nsmgr)?.SelectNodes("*");

            if (rPrChildList == null)
            {
                return;
            }

            foreach (XmlNode child in rPrChildList)
            {
                switch (child.Name)
                {
                    case "w:i":
                        IsItalic = true;
                        break;

                    case "w:b":
                        IsBold = true;
                        break;

                    case "w:u":
                        IsUnderlined = true;
                        break;

                    case "w:color":
                        Color = child.Attributes?["w:val"]?.Value ?? "";
                        break;

                    case "w:highlight":
                        Highlight = child.Attributes?["w:val"]?.Value ?? "";
                        break;

                    default:
                        break;
                }
            }
        }



        public string ToHTML()
        {
            StringBuilder html = new StringBuilder();
            html.Append("<span");
            AttachStyles(html);
            html.Append(">");
            if (IsItalic)
            {
                html.Append("<i>");
            }
            if (IsBold)
            {
                html.Append("<b>");
            }
            if (IsUnderlined)
            {
                html.Append("<u>");
            }
            html.Append(Text);
            if (IsUnderlined)
            {
                html.Append("</u>");
            }
            if (IsBold)
            {
                html.Append("</b>");
            }
            if (IsItalic)
            {
                html.Append("</i>");
            }
            html.Append("</span>");

            return html.ToString();
        }

        private void AttachStyles(StringBuilder html)
        {
            var styleList = new Dictionary<string, string>
            {
                {"color", Color },
                {"background-color", Highlight }
            }
            .Where(x => x.Value != "")
            .Select(x => $"{x.Key}: {x.Value}")
            .ToList();

            if (styleList.Count > 0)
            {
                html.Append($" style=\"{string.Join(';', styleList)}\"");
            }
        }

        public void AppendTo(XmlNode node)
        {
            XmlDocument document = node.OwnerDocument == null ? (XmlDocument)node : node.OwnerDocument;

            // Create the main 'w:r' element
            XmlElement rElement = document.CreateElement("w", "r", document.DocumentElement.NamespaceURI); // Replace with the actual namespace URI

            // Create the 'w:rPr' element
            XmlElement rPrElement = document.CreateElement("w", "rPr", document.DocumentElement.NamespaceURI);

            // Create the 'w:b' element
            XmlElement bElement = document.CreateElement("w", "b", document.DocumentElement.NamespaceURI);

            // Append the 'w:b' element to 'w:rPr'
            if (IsBold)
            {
                rPrElement.AppendChild(bElement);

                XmlElement bCsElement = document.CreateElement("w", "bCs", document.DocumentElement.NamespaceURI);
                // Append the 'w:bCs' element to 'w:rPr'
                rPrElement.AppendChild(bCsElement);
            }

            // Create the 'w:i' element
            if (IsItalic)
            {
                XmlElement iElement = document.CreateElement("w", "i", document.DocumentElement.NamespaceURI);

                // Append the 'w:i' element to 'w:rPr'
                rPrElement.AppendChild(iElement);

                // Create the 'w:iCs' element (only if 'w:i' exists)
                if (true /* Add your condition here based on your logic */)
                {
                    XmlElement iCsElement = document.CreateElement("w", "iCs", document.DocumentElement.NamespaceURI);
                    // Append the 'w:iCs' element to 'w:rPr'
                    rPrElement.AppendChild(iCsElement);
                }
            }

            if (IsUnderlined)
            {
                // Create the 'w:u' element
                XmlElement uElement = document.CreateElement("w", "u", document.DocumentElement.NamespaceURI);

                // Create the 'w:u' attribute 'w:val="single"'
                uElement.SetAttribute("val", document.DocumentElement.NamespaceURI, "single");

                // Append the 'w:u' element to 'w:rPr'
                rPrElement.AppendChild(uElement);
            }

            if (Color != "")
            {
                // Create the 'w:color' element
                XmlElement colorElement = document.CreateElement("w", "color", document.DocumentElement.NamespaceURI);

                colorElement.SetAttribute("val", document.DocumentElement.NamespaceURI, Color);

                // Append the 'w:color' element to 'w:rPr'
                rPrElement.AppendChild(colorElement);
            }


            if (Highlight != "")
            {
                // Create the 'w:highlight' element
                XmlElement highlightElement = document.CreateElement("w", "highlight", document.DocumentElement.NamespaceURI);

                highlightElement.SetAttribute("val", document.DocumentElement.NamespaceURI, Highlight);

                // Append the 'w:highlight' element to 'w:rPr'
                rPrElement.AppendChild(highlightElement);
            }

            // Append the 'w:rPr' element to 'w:r'
            rElement.AppendChild(rPrElement);

            // Create the 'w:t' element
            XmlElement tElement = document.CreateElement("w", "t", document.DocumentElement.NamespaceURI);

            if (PreserveWhitespace)
            {
                var spaceAttr = document.CreateAttribute("xml", "space", "urn:xml");
                spaceAttr.Value = "preserve";
                tElement.Attributes.Append(spaceAttr);
            }

            tElement.InnerText = Text;

            // Append the 'w:t' element to 'w:r'
            rElement.AppendChild(tElement);

            // Append the 'w:r' element to the provided 'node'
            node.AppendChild(rElement);
        }
    }

    internal class WordParagraphIndent
    { }

    internal class WordParagraphSpacing
    { }

    internal class WordParagraph
    {
        // TODO: Add indent and spacing
        public WordParagraphIndent Indent { get; set; } = new WordParagraphIndent();

        public WordParagraphSpacing Spacing { get; set; } = new WordParagraphSpacing();
        public string Style { get; set; } = "";
        public string Text { get; set; } = "";
        public List<WordParagraphRange> Ranges { get; set; } = new List<WordParagraphRange>();

        public WordParagraph() { }

        public WordParagraph(XmlNode paragraphNode, XmlNamespaceManager nsmgr)
        {
            Style = paragraphNode.SelectSingleNode("w:pPr/w:pStyle", nsmgr)?.Attributes?["w:val"]?.Value ?? "";
            Text = paragraphNode.InnerText;

            XmlNodeList? rangeNodes = paragraphNode.SelectNodes("w:r", nsmgr);

            if (rangeNodes != null)
            {
                foreach (XmlNode rangeNode in rangeNodes)
                {
                    Ranges.Add(new WordParagraphRange(rangeNode, nsmgr));
                }
            }
        }
        public void AppendTo(XmlNode node)
        {
            XmlDocument document = node.OwnerDocument == null ? (XmlDocument)node : node.OwnerDocument;

            // Create the main 'w:p' element
            XmlElement mainElement = document.CreateElement("w:p", document.DocumentElement.NamespaceURI);

            // Create the 'w:pPr' element
            XmlElement pPrElement = document.CreateElement("w:pPr", document.DocumentElement.NamespaceURI);

            // Create the 'w:pStyle' element with the 'w:val' attribute
            XmlElement pStyleElement = document.CreateElement("w:pStyle", document.DocumentElement.NamespaceURI);
            pStyleElement.SetAttribute("val", document.DocumentElement.NamespaceURI, "StyleValue");

            // Append the 'w:pStyle' element to the 'w:pPr' element
            pPrElement.AppendChild(pStyleElement);

            // Append the 'w:pPr' element to the main 'w:p' element
            mainElement.AppendChild(pPrElement);

            // Add all ranges in paragraph
            foreach (var range in Ranges)
            {
                range.AppendTo(mainElement);
            }

            // Append the main 'w:p' element to the parent document
            node.AppendChild(mainElement);
        }

        public void AppendToEx(XmlNode node)
        {
            XmlDocument document = node.OwnerDocument == null ? (XmlDocument)node : node.OwnerDocument;

            // Create the main 'w:p' element
            XmlElement mainElement = document.CreateElement("w", "p", "urn:bitch");

            // Add a text document to the 'w:p' element
            XmlElement textElement = document.CreateElement("w", "t", "urn:bitch");
            textElement.InnerText = Text;
            mainElement.AppendChild(textElement);

            // Create the 'w:pPr' element
            XmlElement pPrElement = document.CreateElement("w", "pPr", "urn:bitch");

            // Create the 'w:pStyle' element with the 'w:val' attribute
            XmlElement pStyleElement = document.CreateElement("w", "pStyle", "urn:bitch");
            pStyleElement.SetAttribute("w:val", "StyleValue");

            // Append the 'w:pStyle' element to the 'w:pPr' element
            pPrElement.AppendChild(pStyleElement);

            // Append the 'w:pPr' element to the main 'w:p' element
            mainElement.AppendChild(pPrElement);

            // Append the main 'w:p' element to the parent document
            node.AppendChild(mainElement);
        }
    }

    /// <summary>
    /// Represents a Word document, containing all paragraphs within it.
    ///
    /// <para>TODO: Add `w:sectPr` page info</para>
    /// </summary>
    internal class WordDocument
    {
        private string filePath = "";
        public List<WordParagraph> Paragraphs { get; set; } = new List<WordParagraph>();
        public string Text { get; set; }

        public WordDocument(string path)
        {
            filePath = path;
            XmlDocument? wordDocument = getDocumentInArchive(path, "document.xml");

            if (wordDocument == null)
            {
                throw new Exception("File not found.");
            }

            Text = wordDocument.InnerText;

            XmlNamespaceManager nsmgr = new XmlNamespaceManager(wordDocument.NameTable);
            nsmgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            XmlNodeList paragraphNodes = wordDocument.GetElementsByTagName("w:p");
            foreach (XmlNode paragraphNode in paragraphNodes)
            {
                Paragraphs.Add(new WordParagraph(paragraphNode, nsmgr));
            }
        }

        public Stream ToStream()
        {
            var xmlDocument = new XmlDocument();
            xmlDocument.LoadXml(Constants.DOCUMENT_HEADER + Constants.DOCUMENT_FOOTER);

            XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDocument.NameTable);
            nsmgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            foreach (var paragraph in Paragraphs)
            {
                paragraph.AppendTo(xmlDocument.GetElementsByTagName("w:body")[0]!);
            }

            var stream = new MemoryStream();
            xmlDocument.Save(stream);
            stream.Position = 0;

            return stream;
        }

        private XmlDocument? getDocumentInArchive(string path, string search)
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

        public void SaveTo(string path)
        {
            using (var memoryStream = new MemoryStream())
            using (var originalFile = File.OpenRead(filePath))
            {
                using (var archive = new ZipArchive(memoryStream, ZipArchiveMode.Create, true))
                using (var modifiedZip = new ZipArchive(originalFile, ZipArchiveMode.Read))
                {
                    foreach (var entry in modifiedZip.Entries)
                    {
                        if (entry.Name == "document.xml") continue;

                        var newEntry = archive.CreateEntry(entry.FullName);

                        using (var newEntryStream = newEntry.Open())
                        using (var entryStream = entry.Open())
                        {
                            entryStream.CopyTo(newEntryStream);
                        }
                    }

                    var documentEntry = archive.CreateEntry("word/document.xml");

                    using (var documentStream = ToStream())
                    using (var documentEntryStream = documentEntry.Open())
                    {
                        documentStream.CopyTo(documentEntryStream);
                    }
                }

                using (var fileStream = new FileStream(path, FileMode.Create))
                {
                    memoryStream.Seek(0, SeekOrigin.Begin);
                    memoryStream.CopyTo(fileStream);
                }
            }
        }
    }
}
