using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace InteropWord
{
    public class Metodos
    {
        public static string caminho { get; set; }
        public static Stream BuscarArquivo()
        {
            string caminhoArquivo;

            try
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.InitialDirectory = "c:\\";
                    openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                    openFileDialog.FilterIndex = 2;
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        //Get the path of specified file
                        caminhoArquivo = openFileDialog.FileName;
                        caminho = caminhoArquivo.Trim();    
                        //Read the contents of the file into a stream
                        var fileStream = openFileDialog.OpenFile();
                        return fileStream;
                    }
                    else
                    {
                        return new MemoryStream();
                    }
                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Aviso!", MessageBoxButtons.OK);
            }

            return new MemoryStream();
        }

        public static void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object toFindText, object replaceWithText)
        {
            try
            {


                object matchCase = true;
                object matchwholeWord = true;
                object matchwildCards = false;
                object matchSoundLike = false;
                object nmatchAllforms = false;
                object forward = true;
                object format = false;
                object matchKashida = false;
                object matchDiactitics = false;
                object matchAlefHamza = false;
                object matchControl = false;
                object read_only = false;
                object visible = true;
                object replace = -2;
                object wrap = 1;

                wordApp.Selection.Find.Execute(ref toFindText, ref matchCase,
                                                ref matchwholeWord, ref matchwildCards, ref matchSoundLike,

                                                ref nmatchAllforms, ref forward,

                                                ref wrap, ref format, ref replaceWithText,

                                                    ref replace, ref matchKashida,

                                                ref matchDiactitics, ref matchAlefHamza,

                                                 ref matchControl);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Aviso!", MessageBoxButtons.OK);
            }

        }

        public static void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object toFindText, object replaceWithText, object filename)
        {
            try
            {
                object missing = Missing.Value;

                Microsoft.Office.Interop.Word.Document myWordDoc = null;

                if (File.Exists((string)filename))
                {
                    object readOnly = false;

                    object isvisible = false;

                    wordApp.Visible = false;
                    myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                                        ref missing, ref missing, ref missing,
                                                        ref missing, ref missing, ref missing,
                                                        ref missing, ref missing, ref missing,
                                                         ref missing, ref missing, ref missing, ref missing);
                    myWordDoc.Activate();


                    object matchCase = true;
                    object matchwholeWord = true;
                    object matchwildCards = false;
                    object matchSoundLike = false;
                    object nmatchAllforms = false;
                    object forward = true;
                    object format = false;
                    object matchKashida = false;
                    object matchDiactitics = false;
                    object matchAlefHamza = false;
                    object matchControl = false;
                    object read_only = false;
                    object visible = true;
                    object replace = -2;
                    object wrap = 1;

                    wordApp.Selection.Find.Execute(ref toFindText, ref matchCase,
                                                    ref matchwholeWord, ref matchwildCards, ref matchSoundLike,

                                                    ref nmatchAllforms, ref forward,

                                                    ref wrap, ref format, ref replaceWithText,

                                                        ref replace, ref matchKashida,

                                                    ref matchDiactitics, ref matchAlefHamza,

                                                     ref matchControl);
                    //myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                    //                                                ref missing, ref missing, ref missing,
                    //                                                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    //                                                ref missing, ref missing, ref missing);

                    //myWordDoc.Close();
                    wordApp.Quit();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Aviso!", MessageBoxButtons.OK);
            }

        }


        public static void CreateWordDocument(object filename, object SaveAs, string textoAntigo, string textoNovo)
        {


            try
            {
                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                object missing = Missing.Value;

                Microsoft.Office.Interop.Word.Document myWordDoc = null;

                if (File.Exists((string)filename))
                {
                    object readOnly = false;

                    object isvisible = false;

                    wordApp.Visible = false;
                    myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                                        ref missing, ref missing, ref missing,
                                                        ref missing, ref missing, ref missing,
                                                        ref missing, ref missing, ref missing,
                                                         ref missing, ref missing, ref missing, ref missing);
                    myWordDoc.Activate();
                    myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                                                                    ref missing, ref missing, ref missing,
                                                                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                                    ref missing, ref missing, ref missing);

                    myWordDoc.Close();
                    wordApp.Quit();


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Aviso!", MessageBoxButtons.OK);
            }

        }

        public static void CreateWordDocument(object filename, object SaveAs)
        {
            try
            {
                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                object missing = Missing.Value;

                Microsoft.Office.Interop.Word.Document myWordDoc = null;

                if (File.Exists((string)filename))
                {
                    object readOnly = false;

                    object isvisible = false;

                    wordApp.Visible = false;
                    myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                                        ref missing, ref missing, ref missing,
                                                        ref missing, ref missing, ref missing,
                                                        ref missing, ref missing, ref missing,
                                                         ref missing, ref missing, ref missing, ref missing);
                    myWordDoc.Activate();
                    myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                                                                    ref missing, ref missing, ref missing,
                                                                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                                    ref missing, ref missing, ref missing);

                    myWordDoc.Close();
                    wordApp.Quit();


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Aviso!", MessageBoxButtons.OK);
            }

        }


    }
}
