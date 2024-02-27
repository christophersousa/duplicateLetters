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
        string filePathCopy = "C:\\Users\\Christopher\\Documents\\raizen\\doc_copy.docx";

        CopyFile(filePath, filePathCopy);

        DuplicateBlock(filePathCopy, 2);

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

    static void CopyFile(string fileOrigin, string fileDestiny)
    {
        try
        {
            using (FileStream fsOrigin = new FileStream(fileOrigin, FileMode.Open, FileAccess.Read))
            using (FileStream fsDestiny = new FileStream(fileDestiny, FileMode.Create, FileAccess.Write))
            {
                fsOrigin.CopyTo(fsDestiny);
            }
        }
        catch (IOException ex)
        {
            Console.WriteLine($"Erro ao copiar arquivo: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro geral: {ex.Message}");
        }
    }

    static OpenXmlElement CloneElement(OpenXmlElement element)
    {
        return element.CloneNode(true);
    }
}