using Aylien.TextApi;
using System;
using Microsoft.Office.Interop.Excel;
using System.Linq;
using Microsoft.CSharp;
namespace ConsoleApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            Client client = new Client("fcd04c9b", "0bfa15d2ec326dc77265fdf6f2461ff2");
            //String text = "for architecture is not mere building, but beautiful building. It began when for the first time a man or a woman thought of a dwelling in terms of appearance as well as of use. Probably this effort to give beauty or sublimity to a structure was directed first to graves rather than to homes; while the commemorative pillar developed into statuary, the tomb grew into a temple";

            //var hashtags = client.Hashtags(text: text);
            //Console.WriteLine(string.Join(", ", hashtags.HashtagsMember));
            //Console.ReadLine();
            Microsoft.Office.Interop.Excel.Application xlsApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlsApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }

            //Displays Excel so you can see what is happening
            //xlsApp.Visible = true;
            Workbook wb = xlsApp.Workbooks.Open("C:\\Users\\mobin\\Csharp\\aylien_textapi_csharp-master\\ConsoleApplication\\bin\\Debug\\Citations_History_OurOrientalHeritage.xlsx",
                                             ReadOnly: false);
            xlsApp.AlertBeforeOverwriting = false;
            try
            {
                Sheets sheets = wb.Worksheets;
                Worksheet ws = (Worksheet)sheets.get_Item(1);

                Range firstColumn = (Range)ws.UsedRange.Columns[1];
                System.Array myvalues = (System.Array)firstColumn.Cells.Value;
                string[] strArray = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();
                for (int i = 400; i < 1000; i++)
                {
                    var hashtags = client.Hashtags(text: strArray[i]);
                    string hashtag = String.Join(", ", hashtags.HashtagsMember);
                    if (hashtag != "" && hashtag != null)
                        ws.Cells[i + 1, 9] = hashtag;
                    wb.Save();
                    Console.WriteLine(i);
                }
            }
            finally
            {
                
                
                wb.Close();
            }
        }
    }
}