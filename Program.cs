using System;
using System.Xml;
using System.Xml.Linq;
using System.Drawing;

// INFO
/*
 * the C# CSharpConsole.exe and all relevant file must be copied in the CaaG/public/xml2html folder
 * of the CaaG Web application
 * -
 * to test this Console app before using in in the CaaG Web app
 * proceed as follow:
 * in the "Project" menu, select the "CSharpConsole Properties" sub-menu (the last one)
 * and in the "debug" vertical tab, edit the "Application arguments" that could be like:
 * "C:\Users\htaverni\source\repos\CaaG\CaaG\public\folders\7fcbf81b-4877-4e70-8d15-c5322762ce94\Walgreens - Fully Executed MPSA_XXX" 
 * it is now a unique parameters, as previous multiple parameters are now automatically produce from this unique one
 * 
 * and it could be as simple as:
 * "C:\Users\htaverni\Desktop\simple.docx"
 * 
 * to test this, you must 
 * - unzip the simple.docx file into C:\Users\htaverni\Desktop\simple.docx_unzipped
 * - rename the simple.docx file, as the exec will create a folder with the same name
 * 
 * to use the exec as part of CaaG, you must copy the content of the folder:
 * C:\Users\htaverni\source\repos\CaaG\CSharpConsole\bin\Debug\net5.0
 * Please, check that you have build the code in "debug" more
 * to the following place:
 * C:\Users\htaverni\source\repos\CaaG\CaaG\public\xml2html
 */


namespace CSharpConsole
{
    
    class Program
    {



        static void Main(string[] args)
        {
            string uniqueParameter;
            try { uniqueParameter = args[0];}
            catch (Exception e) {Console.WriteLine("Missing parameter in command line (like \"C:\\Users\\htaverni\\Desktop\\simple.docx\") => " + e.Message);return;}
            Console.WriteLine("uniqueParameter => " + uniqueParameter);
            int slash = uniqueParameter.LastIndexOf(@"\");
            Console.WriteLine("slash=" + slash);
            string fileNameDstWithoutExtension = uniqueParameter.Substring(slash + 1);
            string folderNameSrc = uniqueParameter + "_unzipped";
            string folderNameDst = uniqueParameter.Substring(0, slash);
            // string folderNameSrc = args[0];
            // string folderNameDst = args[1];
            // string fileNameDstWithoutExtension = args[2];
            Console.WriteLine("folderNameSrc => " + folderNameSrc);
            Console.WriteLine("folderNameDst => " + folderNameDst);
            Console.WriteLine("fileNameDst => " + fileNameDstWithoutExtension);
            ConvertDocx2Html.henri(folderNameSrc, folderNameDst, fileNameDstWithoutExtension);
        }
    }
}
