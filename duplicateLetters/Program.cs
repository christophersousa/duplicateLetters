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
                        {"@period", $"Período {i}"},
                        {"@CARGODAPESSOA", dataJson.Periodo[i].Data},
                        {"@AREADAEMPRESA", dataJson.Periodo[i].Nome}
                    };

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

    static OpenXmlElement CloneElement(OpenXmlElement element)
    {
        return element.CloneNode(true);
    }
}