using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;

class Program
{
    static void Main()
    {
        string filePath = "C:\\Users\\Christopher\\Documents\\raizen\\doc.docx";

        DuplicateBlock(filePath, 2);

        Console.WriteLine("Bloco duplicado com sucesso no mesmo arquivo!");
    }

    static void DuplicateBlock(string filePath, int numberOfCopies)
    {
        using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
        {
            Body body = doc.MainDocumentPart.Document.Body;

            var startParagraph = body.Elements<Paragraph>().FirstOrDefault(para => para.InnerText.Contains("@start"));
            var endParagraph = body.Elements<Paragraph>().FirstOrDefault(para => para.InnerText.Contains("@end"));
            if (startParagraph != null && endParagraph != null)
            {

                var elementsToDuplicate = body.Elements()
                    .SkipWhile(el => el != startParagraph)
                    .TakeWhile(el => el != endParagraph)
                    .ToList();

                for (int i = 0; i < numberOfCopies; i++)
                {
                    var clonedElements = elementsToDuplicate.Select(el => CloneElement(el)).ToList();

                    foreach (var clonedElement in clonedElements)
                    {
                        body.InsertBefore(clonedElement, endParagraph);
                    }
                }

            }
        }
    }

    static OpenXmlElement CloneElement(OpenXmlElement element)
    {
        return element.CloneNode(true);
    }
}