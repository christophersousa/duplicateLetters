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
using duplicateLetters.model;
using Newtonsoft.Json;

class Program
{
    static void Main()
    {
        string filePath = "C:\\Users\\Christopher\\Documents\\raizen\\doc.docx";
        string filePathCopy = "C:\\Users\\Christopher\\Documents\\raizen\\doc_copy.docx";

        string jsonContent = File.ReadAllText("C:\\Users\\Christopher\\Documents\\raizen\\data.json");

        var dataJson = JsonConvert.DeserializeObject<DataJson>(jsonContent);

        File.Copy(filePath, filePathCopy, true);

        DuplicateBlock(filePathCopy, dataJson);

        Console.WriteLine("Bloco duplicado com sucesso no mesmo arquivo!");
    }

    static void DuplicateBlock(string filePath, DataJson dataJson)
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
                    .Where(el => el != startParagraph && el != endParagraph)
                    .ToList();

                for (int i = 0; i < dataJson.Periodo.Count; i++)
                {
                    var clonedElements = elementsToDuplicate.Select(el => CloneElement(el)).ToList();

                    // Created page
                    Paragraph pageSeparator = new Paragraph(new Run(new Break() { Type = BreakValues.Page }));

                    // Insert the page separator before the endParagraph
                    body.InsertBefore(pageSeparator, endParagraph);


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