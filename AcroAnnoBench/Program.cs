using Acrobat;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace AcroAnnoBench
{
    internal class Program
    {
        static int REPETITION = 100;
        static void Main(string[] args)
        {
            AcroPDDoc pddoc = null;
            long[] processTime = new long[REPETITION];
            long numOfAnnotations = 0;
            Console.WriteLine("Processing " + REPETITION + " documents....");

            for(int i=0; i<REPETITION; i++)
            {
                long startTime = CurrentTimeMillis();
                try
                {
                    pddoc = new AcroPDDoc();
                    pddoc.Open(@"C:\Users\Administrator\source\repos\AcroAnnoBench\AcroAnnoBench\acrobat_reference.pdf");

                    numOfAnnotations = 0;
                    for (var pageIndex = 0; pageIndex < pddoc.GetNumPages(); pageIndex++)
                    {
                        CAcroPDPage pdpage = (CAcroPDPage)pddoc.AcquirePage(pageIndex);
                        
                        for (var annotIndex = 0; annotIndex < pdpage.GetNumAnnots(); annotIndex++)
                        {
                            CAcroPDAnnot annotation = (CAcroPDAnnot)pdpage.GetAnnot(annotIndex);
                            string subType = annotation.GetSubtype();
                        }
                        numOfAnnotations += pdpage.GetNumAnnots();
                    }

                }
                finally
                {
                    pddoc?.Close();
                    Marshal.ReleaseComObject(pddoc);
                    pddoc = null;
                }
                long endTime = CurrentTimeMillis();
                processTime[i] = endTime - startTime;
            }
            Console.WriteLine("Average Process Time (sec): " + CalculateAverageProcessTime(processTime) / 1000.0);
            Console.WriteLine("Num of annotations : " + numOfAnnotations);
        }
        private static double CalculateAverageProcessTime(long[] processTime)
        {
            double sum = 0;
            foreach(long num in processTime)
            {
                sum += num;
            }
            
            return (double)sum / processTime.Length;
        }

        public static long CurrentTimeMillis()
        {
            long ticks = DateTime.Now.Ticks;
            return  ticks / TimeSpan.TicksPerMillisecond;
        }
    }


}
