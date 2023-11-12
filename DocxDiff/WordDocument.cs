using System.IO.Compression;
using System.Xml;

namespace DocxDiff
{
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
