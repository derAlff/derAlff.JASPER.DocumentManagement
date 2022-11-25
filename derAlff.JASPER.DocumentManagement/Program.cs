using System.Net.NetworkInformation;
using TesseractOCR;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using PdfSharpCore;
using PdfSharpCore.Pdf.IO;
using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.Advanced;
using TesseractOCR.Enums;
using derAlff.JASPER.Logger;

namespace derAlff.JASPER.DocumentManagement
{
    internal class Program
    {
        

        public static Logging logging = new Logging();

        public static String WorkingDirectory = String.Empty;
        public static String PathToPdfDocuments = "\\input";
        public static String PathToTessdata = ".\\tessdata";
        public static String PathToConfigFolder = "\\config";
        public static String PathToRulesJson = "\\rules.json";

        public static String Search = "";

        public static List<String> PdfDocuments = new List<String>();
        static void Main(string[] args)
        {
            logging.Init(@"D:\", "derAlff.JASPER.DocumentManagement.log");
            logging.SetLoglevel(LogLevel.debug);
            logging.Log("Started 'Main()'");

            String[] tmp = null;
            String ocredText = String.Empty;
            Console.WriteLine("Start program");
            Console.WriteLine(@$"Check folder '{PathToPdfDocuments}'");

            WorkingDirectory = Directory.GetCurrentDirectory();
            PathToPdfDocuments = WorkingDirectory + PathToPdfDocuments;

            if (!Directory.Exists(PathToPdfDocuments))
            {
                logging.Log(@$"The folder '{PathToPdfDocuments}' do not exists. Create the folder now...");
                Directory.CreateDirectory(PathToPdfDocuments);
            }

            logging.Log($@"The folder '{PathToPdfDocuments}' exists. Check for pdf documents", LogLevel.debug);
            logging.Log($@"Get all PDF files from '{PathToPdfDocuments}'", LogLevel.debug);
            
            PdfDocuments = Directory.GetFiles(PathToPdfDocuments, "*.pdf", SearchOption.AllDirectories).ToList();

            logging.Log($@"Found '{PdfDocuments.Count}' documents.", LogLevel.debug);

            foreach(string doc in PdfDocuments)
            {
                String targetFolder;

                //ReadPdf(doc, out ocredText);
                ReadPdfTesseract(doc, out ocredText);
                RunRules(ocredText, out targetFolder);
                if(CheckFolder(targetFolder))
                {
                    MovePdf(doc, targetFolder, Path.GetFileName(doc));
                }
            }

            logging.Log("Ended 'Main()'");
        }

        static bool ReadConfig()
        {
            bool returnVal = false;

            try
            {

            }
            catch { 
                returnVal = false; 
            }

            return returnVal;
        }
        static bool ReadPdf(String PdfFile, out String Result)
        {
            String functionName = "ReadPdf" + "()";
            bool returnVal = false;
            Result = String.Empty;

            try
            {
                /*var ocr = new IronTesseract();
                using (var input = new OcrInput())
                {
                    input.AddPdf(PdfFile);
                    var r = ocr.Read(input);
                    Result = r.Text;
                    Console.WriteLine(r.Text);
                }*/

                returnVal = true;
            }
            catch(Exception ex){
                Console.WriteLine($"Error in function '{functionName}'\n{ex.Message}");
                returnVal = false;
            }

            return returnVal;
        }

        static bool ReadPdfTesseract(String PdfFile, out String Result)
        {
            String functionName = "ReadPdfTesseract" + "()";
            logging.Log("Started " + functionName + "()");
            
            bool returnVal = false;
            Result = String.Empty;

            try
            {
                PdfSharpCore.Pdf.PdfDocument document = PdfReader.Open(PdfFile);
                int imageCount = 0;

                logging.Log(@$"Get document pages from '{PdfFile}'", LogLevel.debug);
                foreach(PdfPage page in document.Pages)
                {
                    PdfDictionary resources = page.Elements.GetDictionary("/Resources");
                    if(resources != null)
                    {
                        PdfDictionary xObjects = resources.Elements.GetDictionary("/XObject");
                        if (xObjects != null)
                        {
                            ICollection<PdfItem> items = xObjects.Elements.Values;
                            
                            // Iterate references to external objects
                            foreach (PdfItem item in items)
                            {                                
                                PdfReference reference = item as PdfReference;
                                
                                if (reference != null)
                                {
                                    PdfDictionary xObject = reference.Value as PdfDictionary;
                                    // Is external object an image?
                                    if (xObject != null && xObject.Elements.GetString("/Subtype") == "/Image")
                                    {
                                        String imageFilePath = String.Empty;
                                        MemoryStream ms = null;
                                                                                
                                        ExportJpegImage(xObject, ref imageCount, Path.GetFileName(PdfFile), out imageFilePath);
                                        //OCR
                                        Result += "";

                                        logging.Log(@$"Export JPEG to '{imageFilePath}'", LogLevel.debug);

                                        using var engine = new Engine(@"./tessdata", Language.German, EngineMode.Default);
                                        using var img = TesseractOCR.Pix.Image.LoadFromFile(imageFilePath);
                                        using var p = engine.Process(img);
                                        Result += p.Text;
                                    }
                                }

                            }
                        }
                    }
                }

                RemoveAllTempFiles(Directory.GetCurrentDirectory() + "\\input\\temp\\");

                returnVal = true;
            }
            catch (Exception ex)
            {
                logging.Log($"Error in function '{functionName}'\n{ex.Message}", LogLevel.error);
                returnVal = false;
            }

            logging.Log("Ended " + functionName + "()");
            return returnVal;
        }

        static void ExportJpegImage(PdfDictionary image, ref int count, String filename, out String imagePath)
        {
            String functionName = "ExportJpegImage()";
            logging.Log("Started " + functionName + "");

            try
            {
                // Fortunately JPEG has native support in PDF and exporting an image is just writing the stream to a file.
                byte[] stream = image.Stream.Value;
                imagePath = String.Empty;
                imagePath = Directory.GetCurrentDirectory() + "\\input\\temp\\" + filename + count.ToString() + ".jpeg";

                FileStream fs = new FileStream(String.Format(imagePath, count++), FileMode.Create, FileAccess.Write);

                BinaryWriter bw = new BinaryWriter(fs);
                bw.Write(stream);
                bw.Close();
            }
            catch (Exception ex)
            {
                logging.Log($"Error in function '{functionName}'\n{ex.Message}", LogLevel.error);
                imagePath = String.Empty;
            }

            logging.Log("Ended " + functionName + "");

        }

        static bool RemoveAllTempFiles(String Path)
        {
            String functionName = "RemoveAllTempFiles" + "()";
            logging.Log("Started " + functionName + "");

            bool returnVal = false;

            try
            {
                logging.Log(@$"Delete all temp files in '{Path}'", LogLevel.debug);
                DirectoryInfo di = new DirectoryInfo(Path);
                foreach(String file in Directory.GetFiles(Path))
                {
                    File.Delete(file);
                }

            }
            catch (Exception ex)
            {
                logging.Log($"Error in function '{functionName}'\n{ex.Message}", LogLevel.error);
                returnVal = false;
            }

            logging.Log("Ended " + functionName + "");
            return returnVal;
        }

        static bool RunRules(String PdfText, out String TargetFolder)
        {
            String functionName = "RunRules" + "()";
            logging.Log("Started " + functionName + "");

            bool returnVal = false;
            TargetFolder = String.Empty;

            bool found = false;
            String searchVal = String.Empty;
            String globalTargetFolder = String.Empty;
            String targetFolder = String.Empty;
            String targetSubFolder = String.Empty;
            // Open configuration
            try
            {
                if(Directory.Exists(WorkingDirectory + PathToConfigFolder))
                {
                    if(File.Exists(WorkingDirectory + PathToConfigFolder + PathToRulesJson))
                    {
                        // Open rules
                        JObject jRules = JObject.Parse(File.ReadAllText(WorkingDirectory + PathToConfigFolder + PathToRulesJson));

                        foreach(JProperty rule in (JToken)jRules["Rules"])
                        {
                            if(!found)
                            {
                                String ruleJson = jRules["Rules"][rule.Name].ToString();

                                JObject jRule = JObject.Parse(ruleJson);

                                if (jRule.ContainsKey("Invoice"))
                                {
                                    searchVal = jRules["Rules"][rule.Name]["Invoice"]["SearchValues"].ToString();   
                                    targetSubFolder = jRules["Rules"][rule.Name]["Invoice"]["Subfolder"].ToString();
                                }
                                else if(jRule.ContainsKey("AB"))
                                {
                                    searchVal = jRules["Rules"][rule.Name]["AB"]["SearchValues"].ToString();
                                    targetSubFolder = jRules["Rules"][rule.Name]["AB"]["Subfolder"].ToString();
                                }
                                else if (jRule.ContainsKey("Document"))
                                {
                                    searchVal = jRules["Rules"][rule.Name]["Document"]["SearchValues"].ToString();
                                    targetSubFolder = jRules["Rules"][rule.Name]["Document"]["Subfolder"].ToString();
                                }

                                if (!String.IsNullOrEmpty(searchVal))
                                {
                                    bool matched = true;
                                    List<String> searchList = searchVal.Split(' ').ToList();

                                    foreach(String searchItem in searchList)
                                    {
                                        if(PdfText.ToLower().Contains(searchItem.ToLower()))
                                        {
                                            matched &= true;
                                        }
                                        else
                                        {
                                            matched &= false;
                                        }
                                    }

                                    if(matched)
                                    {
                                        found = true;
                                        globalTargetFolder = jRules["GlobalTargetFolder"].ToString();
                                        targetFolder = jRules["Rules"][rule.Name]["TargetFolder"].ToString();
                                        TargetFolder = globalTargetFolder + targetFolder + targetSubFolder;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        throw new Exception(@$"'rules.json' is not available! in folder {WorkingDirectory + PathToConfigFolder + PathToRulesJson}");
                    }
                }

                returnVal = true;
            }
            catch(Exception ex)
            {

                logging.Log($"Error in function '{functionName}'\n{ex.Message}", LogLevel.error);
                returnVal = false;
            }

            logging.Log("Ended " + functionName + "");
            return returnVal;
        }

        static bool CheckFolder(String Path)
        {
            String functionName = "CheckFolder" + "()";
            logging.Log("Started " + functionName + "");

            bool returnVal = false;

            try
            {
                if(!String.IsNullOrEmpty(Path))
                {
                    if (Directory.Exists(Path))
                    {
                        returnVal = true;
                    }
                    else
                    {
                        Directory.CreateDirectory(Path);
                    }
                }
                
            }
            catch(Exception ex)
            {
                logging.Log($"Error in function '{functionName}'\n{ex.Message}", LogLevel.error);
                returnVal = false;
            }

            logging.Log("Ended " + functionName + "");
            return returnVal;
        }

        static bool MovePdf(String Source, String Target, String FileName)
        {
            String functionName = "MovePdf" + "()";
            logging.Log("Started " + functionName + "");

            bool returnVal = false;
                        
            try
            {
                if(!String.IsNullOrEmpty(Target))
                {
                    if (Target[0] == '.')
                    {
                        String tmp = Directory.GetCurrentDirectory();
                        Target = Target.Substring(1, Target.Length - 1);
                        Target = tmp + Target;
                    }
                    if (FileName[0] != '\\')
                    {
                        FileName = "\\" + FileName;
                    }

                    File.Move(Source, Target + FileName);
                    returnVal = true;
                }
                                
            }
            catch (Exception ex)
            {
                logging.Log($"Error in function '{functionName}'\n{ex.Message}", LogLevel.error);
                returnVal = false;
            }

            logging.Log("Ended " + functionName + "");
            return returnVal;
        }
    }
}