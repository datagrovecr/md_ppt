using System.IO.Compression;
using Ppt_lib;

internal class Program
{
    static async Task Main(string[] args)
    {

        var outdir = @"./../../../../md_ppt/test_results/";
        string[] files = Directory.GetFiles(@"./../../../../md_ppt/folder_tests/", "*.md", SearchOption.TopDirectoryOnly);

        foreach (var mdFile in files)
        {
            //Just getting the end route
            string fn = Path.GetFileNameWithoutExtension(mdFile);
            string root = outdir + fn.Replace("_md", "");
            var docxFile = root + ".ppt";
            try
            {
                // markdown to docx
                var md = File.ReadAllText(mdFile);
                var inputStream = new MemoryStream();
                await DgPpt.md_to_ppt(md, inputStream);

                //inputStream is writing into the .docx file
                File.WriteAllBytes(docxFile, inputStream.ToArray());


                // convert the docx back to markdown.
                
                /*using (var instream = File.Open(docxFile, FileMode.Open))
                {
                    var outstream = new MemoryStream();
                    await DgDocx.docx_to_md(instream, outstream, root);//Previous: instream, outstream, fn.Replace("_md", "")

                    
                }*/

                using (ZipArchive archive = ZipFile.OpenRead(outdir + "test.docx"))
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