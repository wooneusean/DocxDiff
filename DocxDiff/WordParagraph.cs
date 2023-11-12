using System.Xml;

namespace DocxDiff
{
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
            pStyleElement.SetAttribute("val", document.DocumentElement.NamespaceURI, Style);

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
    }
}
