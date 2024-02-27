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
using duplicateLetters.service;
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

        LettersService service = new LettersService();

        service.DuplicateBlock(filePathCopy, dataJson);

        Console.WriteLine("Bloco duplicado com sucesso no mesmo arquivo!");
    }
}