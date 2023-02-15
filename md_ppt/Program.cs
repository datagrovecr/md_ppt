using System.IO;
using System.IO.Compression;
using Ppt_lib;

internal class Program
{
    static async Task Main(string[] args)
    {

        var outdir = @"./../../../../md_ppt/test_results/";
        var outdirMedia = @"./../../../../md_ppt/test_results/results/ppt/media/";
        string[] files = Directory.GetFiles(@"./../../../../md_ppt/folder_tests/", "*.md", SearchOption.TopDirectoryOnly);

        foreach (var mdFile in files)
        {
            //Just getting the end route
            string fn = Path.GetFileNameWithoutExtension(mdFile);
            string root = outdir + fn.Replace("_md", "");
            string rootResult = outdir + "results/"+fn.Replace("_md", "");
            var pptxFile = root + ".pptx";
            try
            {
                #region write file

                // markdown to docx
                var md = File.ReadAllText(mdFile);
                var inputStream = new MemoryStream();
                await DgPpt.md_to_ppt(md, inputStream);

                //inputStream is writing into the .docx file
                File.WriteAllBytes(pptxFile , inputStream.ToArray());
                #endregion

                #region PPT back to markdown

                // convert the docx back to markdown.

             /*   using (var instream = File.Open(pptxFile, FileMode.Open))
                {
                    Directory.CreateDirectory(outdirMedia);
                    Directory.CreateDirectory(rootResult);
                    var outstream = new MemoryStream();
                    await DgPpt.ppt_to_md(instream, outstream, rootResult);

                    //pull the images from "ppt/media"
                    using (ZipArchive archive = new ZipArchive(instream ,ZipArchiveMode.Update, true))
                    {
                        string subDirectory = "ppt/media";
                        // Loop through each entry in the zip file
                        foreach (ZipArchiveEntry entry in archive.Entries)
                        {
                            // Check if the entry is a directory and its name matches the specified subdirectory
                            if (entry.FullName.Contains(subDirectory) && !entry.Name.EndsWith("/"))
                            {
                                
                                // Extract the entry to the specified extract path
                                entry.ExtractToFile(outdirMedia + entry.Name, true);



                            }
                        }

                    }

                }*/
                #endregion

                using (ZipArchive archive = ZipFile.OpenRead(outdir + "test.pptx"))
                {
                    

                    archive.ExtractToDirectory(outdir + "test.unzipped", true);

                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"{mdFile} failed {e}");
            }
        }
    }


}