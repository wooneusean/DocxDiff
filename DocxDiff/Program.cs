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


            modifiedDoc.Paragraphs.Add(new WordParagraph
            {
                Style = "Heading1",
                Ranges = new List<WordParagraphRange>
                    {
                        new WordParagraphRange {
                            Text = "Summary",
                        }
                    }
            });

            foreach (var diff in diffs)
            {
                modifiedDoc.Paragraphs.Add(new WordParagraph
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
}
