using System;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using SprawdzPogode.Handlers;
using SprawdzPogode.Exceptions;
using System.Collections.Generic;

namespace SprawdzPogode.Extractors
{
    class Extractor: IExtractor
    {
        private string[] Lines { set; get; }

        public Extractor(string[] lines)
        {
            Lines = lines;
        }

        public void ExtractData(IOutputHandler ex, IFetchableHandler ch)
        {
            ex.Start();
            ch.Start();
            string[] tab = new string[Lines.Length * 6];
            int i = 0;

            foreach (string line in Lines)
            {

                ch.Handle(new string[1] { line + " pogoda" });
                tab[i] = String.Format("{0:yyyy/MM/dd HH:mm:ss}", DateTime.Now);
                tab[i+1] = line;

                try
                {
                    tab[i + 2] = ch.GetData("wob_tm").ToString();
                    tab[i + 3] = ch.GetData("wob_pp").ToString();
                    tab[i + 4] = ch.GetData("wob_ws").ToString();
                    tab[i + 5] = "Success";
                }
                catch (DataNotFoundException)
                {
                    tab[i + 5] = "Fail";
                }

                i += 6;
            }

            ex.Handle(tab);
            ex.Finish();
            ch.Finish();
        }
    }
}


