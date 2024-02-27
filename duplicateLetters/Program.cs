using System;
using System.Collections;
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
            Dictionary<string, string> dictGeneral = new Dictionary<string, string>
                    {
                        {"@nomedapessoa", dataJson.Nomedapessoa},
                        {"@DATADEADMISSÃO", dataJson.Datadeadmissao},
                        {"@nomedogestor", dataJson.Gestor}
                    };
            ReplaceMarking(doc, dictGeneral);
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
                    Dictionary<string, string> dictPeriod = new Dictionary<string, string>
                    {
                        {"@period", $"Período {i +1}"},
                        {"@AREADAEMPRESA", dataJson.Periodo[i].Nome},
                        {"@CARGODAPESSOA", dataJson.Periodo[i].Data}
                    };

                    var clonedElements = elementsToDuplicate.Select(el => ReplaceAndCloneElement(el, dictPeriod)).ToList();

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

    static void ReplaceMarking(WordprocessingDocument documento, Dictionary<string, string> dictGeneral)
    {
        foreach (var texto in documento.MainDocumentPart.Document.Descendants<Text>())
        {
            foreach (var item in dictGeneral)
            {
                if (texto.Text.Contains(item.Key))
                {
                    texto.Text = texto.Text.Replace(item.Key, item.Value);
                }

            }
        }
    }

    static OpenXmlElement ReplaceAndCloneElement(OpenXmlElement element, Dictionary<string, string> replacements)
    {
        var clonedElement = CloneElement(element);

        foreach (var run in clonedElement.Descendants<Run>())
        {
            foreach (var text in run.Descendants<Text>())
            {
                foreach (var replacement in replacements)
                {
                    text.Text = text.Text.Replace(replacement.Key, replacement.Value);
                }
            }
        }

        return clonedElement;
    }

    static OpenXmlElement CloneElement(OpenXmlElement element)
    {
        return element.CloneNode(true);
    }
}