using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using System.Threading;
using System.Diagnostics;

namespace ausbildungsnachweise
{
    class Program
    {
        public static string current_folder = @"C:\Users\VW1U904\OneDrive - Volkswagen AG\Docs\Documents\Ausbildungsnachweise\04_VVGG\Word";
        static void Main(string[] args)
        { 
            bool started = false;
            foreach(var arg in args)
            {
                if(arg.StartsWith("-"))
                {
                    for(int count = Int32.Parse(arg.Remove(0, 1)); count > 0; count--)
                    {
                        Console.WriteLine("Files left: " + count);
                        StartIteration();
                        started = true;
                    }
                }                
            }  
            if(!started)
            {
                StartIteration();
            }
                     
        }


        static void StartIteration()
        {
            string last_path = Directory.GetFiles(current_folder).OrderBy(x => Int32.Parse(x.Split("\\").Last().Split("_")[1])).Last();
            string last_file = last_path.Split("\\").Last();

            int last_iteration = Int32.Parse(last_file.Split("_")[1]);
            int next_iteration = last_iteration + 1;

            string[] dot_split = last_file.Split("_").Last().Split(".");
            

            string last_firstDate = $"{dot_split[0]}.{dot_split[1]}.{dot_split[4]}";
            string last_secondDate = $"{dot_split[2].Remove(0, 1)}.{dot_split[3]}.{dot_split[4]}";

            DateTime startDate = DateTime.Parse(last_firstDate).AddDays(7);
            DateTime endDate = DateTime.Parse(last_secondDate).AddDays(7);
            
            string newFile = $"Ausbildungsnachweis_{next_iteration}_{startDate.Day.ToString("00")}.{startDate.Month.ToString("00")}.-{endDate.ToString("d", CultureInfo.GetCultureInfo("de-DE"))}.docx";
            string newPath = last_path.Replace(last_file, newFile);

            Console.WriteLine("New Filename: " + newFile);

            Microsoft.Office.Interop.Word.Application fileOpen = new Microsoft.Office.Interop.Word.Application();

            Microsoft.Office.Interop.Word.Document document = fileOpen.Documents.Open(last_path, ReadOnly: false);

            fileOpen.Visible = true;
            document.Activate();

            FindAndReplace(fileOpen, last_firstDate, startDate.ToString("d", CultureInfo.GetCultureInfo("de-DE")));
            FindAndReplace(fileOpen, last_secondDate, endDate.ToString("d", CultureInfo.GetCultureInfo("de-DE")));
            FindAndReplace(fileOpen, last_iteration, next_iteration);

            Console.WriteLine("Saving");
            document.SaveAs2(newPath);
            fileOpen.Quit();
            Console.WriteLine("Done");
            Thread.Sleep(3000);
        }


        static void FindAndReplace(Microsoft.Office.Interop.Word.Application fileOpen, object findText, object replaceWithText)
        {
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            fileOpen.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }   
    }
}
