using System.Text;
using System.Xml;

namespace DocxDiff
{
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
}
