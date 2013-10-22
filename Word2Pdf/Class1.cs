using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.IO;

namespace Word2Pdf
{
    static public class Doc2PDFAtServerClass
    {

        static public void word2PdfFcih(object SourceFileName, object newFileName)
        {

            //Pid++;

            Microsoft.Office.Interop.Word.ApplicationClass MSdoc = null;

            //object Source = "d:\\Document" + Pid.ToString(System.Globalization.CultureInfo.CurrentCulture) + ".doc";

            object Source = SourceFileName;

            object readOnly = false;

            object Unknown = System.Reflection.Missing.Value; //Type.Missing;

            object missing = Type.Missing;

            try
            {

                //Creating the instance of Word Application

                if (MSdoc == null)

                    MSdoc = new Microsoft.Office.Interop.Word.ApplicationClass();

                MSdoc.Visible = false;

                MSdoc.Documents.Open(ref Source, ref Unknown,

                ref readOnly, ref Unknown, ref Unknown,

                ref Unknown, ref Unknown, ref Unknown,

                ref Unknown, ref Unknown, ref Unknown,

                ref Unknown, ref Unknown, ref Unknown, ref Unknown, ref Unknown);

                MSdoc.Application.Visible = false;

                MSdoc.WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateMinimize;

                object FileName = newFileName;

                object FileFormat = WdSaveFormat.wdFormatPDF;

                object LockComments = false;

                object AddToRecentFiles = false;

                object ReadOnlyRecommended = false;

                object EmbedTrueTypeFonts = true;

                object SaveNativePictureFormat = false;

                object SaveFormsData = false;

                object SaveAsAOCELetter = false;

                //object Encoding = MsoEncoding.msoEncodingUSASCII;

                object InsertLineBreaks = false;

                object AllowSubstitutions = false;

                object LineEnding = WdLineEndingType.wdCRLF;

                object AddBiDiMarks = false;

                /*

                to get more details about SaveAs(...) function and it's parameter ,read this microsoft's link

                http://msdn2.microsoft.com/en-us/library/aa662158(office.10).aspx

                */

                MSdoc.ActiveDocument.SaveAs(ref FileName, ref FileFormat, ref LockComments,

                ref missing, ref AddToRecentFiles, ref missing,

                ref ReadOnlyRecommended, ref EmbedTrueTypeFonts,

                ref SaveNativePictureFormat, ref SaveFormsData,

                ref SaveAsAOCELetter, ref /*Encoding*/missing, ref InsertLineBreaks,

                ref AllowSubstitutions, ref LineEnding, ref AddBiDiMarks);

            }

            catch (FileLoadException e)
            {

                Console.WriteLine(e.Message + "Error");

            }

            catch (FileNotFoundException e)
            {

                Console.WriteLine(e.Message + "Error");

            }

            catch (FormatException e)
            {

                Console.WriteLine(e.Message + "Error");

            }

            finally
            {

                if (MSdoc != null)
                {

                    MSdoc.Documents.Close(ref Unknown, ref Unknown, ref Unknown);

                    //WordDoc.Application.Quit(ref Unknown, ref Unknown, ref Unknown);

                }

                // for closing the application

                // WordDoc.Quit(ref Unknown, ref Unknown, ref Unknown);

                MSdoc.Quit(ref Unknown, ref Unknown, ref Unknown);

                MSdoc = null;

            }

        }

    }
}
