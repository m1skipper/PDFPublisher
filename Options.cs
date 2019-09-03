using CommandLine;
using CommandLine.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace PDFPublisher
{
    /// <summary>
    /// Опции командной строки
    /// </summary>
    class Options
    {
        // Values is a command
        [ValueList(typeof(List<string>), MaximumElements = 1)]
        public IList<string> Commands { get; set; }

        [Option('i', "input", 
         HelpText = "Input file or files(divided by ',').")]
        public string Input { get; set; }

        [Option('o', "output",
          HelpText = "Output file.")]
        public string Output { get; set; }

        [Option('t', "type", DefaultValue = "ean8",
         HelpText = "Barcode type: ean8, ean13, code128.")]
        public string Type { get; set; }

        [Option('c', "code", 
         HelpText = "Barcode code.")]
        public string Code { get; set; }

        [Option('x', "offsetx", DefaultValue = 4,
          HelpText = "X offset for barcode")]
        public int OffsetX { get; set; }

        [Option('y', "offsety", DefaultValue = 4,
          HelpText = "Y offset for barcode")]
        public int OffsetY { get; set; }

        //[Option('r', "rotate", DefaultValue = 270, - было до того(в версии LGPL) как я сделал правильный поворот. Сейчас по идее нужно добавить параметр откуда брать offset.
        [Option('r', "rotate", DefaultValue = 0,
          HelpText = "Rotate barcode(degrees)")]
        public int Rotate { get; set; }

        [Option('p', "page", DefaultValue = 0,
            HelpText = "Page number(start from 1), value 0 - all pages, or first page")]
        public int Page { get; set; }
        
        [Option('l', "label", 
            HelpText = "text string")]
        public string Label { get; set; }

        [Option('w', "width", DefaultValue = "",
        HelpText = "Width - number, or word 'fit'")]
        public string Width { get; set; }

        [Option('h', "height", DefaultValue = "",
        HelpText = "Height - number, or word 'fit'")]
        public string Height { get; set; }

        [Option('n', "noText", DefaultValue = false,
        HelpText = "Not display code text under bar")]
        public bool NoText { get; set; }

        [Option('b', "barHeight", DefaultValue = 0.0,
        HelpText = "Height of the bar relative to bar width")]
        public double BarHeight { get; set; }

        [Option('m', "center", DefaultValue = false,
        HelpText = "Center item when insert")]
        public bool Center { get; set; }

        [Option('f', "imageFile",
         HelpText = "Image file(pdf, jpg)")]
        public string ImageFile { get; set; }

        [Option('s', "scale", DefaultValue = 2.0,
            HelpText = "Image scale")]
        public double Scale { get; set; }

        [Option('u', "noClean", DefaultValue = false,
            HelpText = "Not clean temp files")]
        public bool NoClean { get; set; }

        [Option('a', "from", DefaultValue = 0,
        HelpText = "Page number(start from 1), value 0 or 1 - from first page")]
        public int FromPage { get; set; }

        [Option('z', "to", DefaultValue = 0,
        HelpText = "Page number(start from 1), value 0 - all pages")]
        public int ToPage { get; set; }

        [Option("log", DefaultValue = "",
            HelpText = "Log file")]
        public string Log { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            var ht = HelpText.AutoBuild(this,
              (HelpText current) => HelpText.DefaultParsingErrorsHandler(this, current));

            var sb = new StringBuilder();
            sb.AppendLine();
            sb.AppendLine("Available commands:");
            sb.AppendLine();

            sb.AppendLine(" combine - combine several pdf files into one pdf file");
            sb.AppendLine("  options: --input, --output");
            sb.AppendLine("  example: PDFPublisher.exe combine --input=\"file1.pdf,file2.pdf\" --output=\"result.pdf\"");
            sb.AppendLine();

            sb.AppendLine(" cut - cut pages from pdf file");
            sb.AppendLine("  options: --input, --output, --from, --to");
            sb.AppendLine("  example: PDFPublisher.exe cut --input=\"file1.pdf\" --output=\"result.pdf\" --from=1 --to=10");
            sb.AppendLine();

            sb.AppendLine(" barcode - print barcode on each page of pdf document");
            sb.AppendLine("  options: --input, --output, --type, --code, --offsetx, --offsety");
            sb.AppendLine("  example: PDFPublisher.exe barcode --type=ean8 --code=1234567 --input=\"file.pdf\" --output=\"result.pdf\"");
            sb.AppendLine();

            sb.AppendLine(" barcodeonlabel - print barcode over label on each page of pdf document");
            sb.AppendLine("  options: --input, --output, --type, --code, --label, --width, --height, --notext, --barheight");
            sb.AppendLine("  example: PDFPublisher.exe barcodeonlabel --label=\"[<BARCODE_PLACEHOLDER>]\" --type=code128 --code=123456789 --height=fit --input=\"file.pdf\" --output=\"result.pdf\"");
            sb.AppendLine();

            sb.AppendLine(" imageonlabel - print image over label on each page of pdf document");
            sb.AppendLine("  options: --input, --output, --label, --imagefile, --width, --height, --center, --log");
            sb.AppendLine("  example: PDFPublisher.exe imageonlabel --label=\"[<BARCODE_PLACEHOLDER>]\" --imageFile=test.png --width=fit --input=\"file.pdf\" --output=\"result.pdf\" --log=\"log.txt\"");
            sb.AppendLine();

            sb.AppendLine(" barcodereplace - replace(print over) one barcode in pdf document by other(code 128)");
            sb.AppendLine("  options: --input, --output, --label, --code, --notext, --page");
            sb.AppendLine("  example: PDFPublisher.exe barcodereplace --label=123456789 --code=000000000 --input=\"file.pdf\" --output=\"result.pdf\"");
            sb.AppendLine();
            
            sb.AppendLine(" convert - convert image file(png, jpg, tiff) to pdf");
            sb.AppendLine("  options: --input, --output");
            sb.AppendLine("  example: PDFPublisher.exe convert --input=\"file.jpg\" --output=\"result.pdf\"");
            sb.AppendLine();

            sb.AppendLine(" image - take page from pdf as png image");
            sb.AppendLine("  options: --input, --output, --page");
            sb.AppendLine("  example: PDFPublisher.exe image --input=\"file.png\" --output=\"result.png\" --page=1");
            sb.AppendLine();

            sb.AppendLine(" extractimages - extract all images from pdf and save them as jpeg files.");
            sb.AppendLine("  options: --input");
            sb.AppendLine();

            sb.AppendLine(" scanbarcode - scan barcode from pdf or png image");
            sb.AppendLine("  options: --input, --page, --clear, --scale");
            sb.AppendLine("  example: PDFPublisher.exe scanbarcode --input=\"file.pdf\" --page=1");
            sb.AppendLine();

            sb.AppendLine(" find - find text label in pdf");
            sb.AppendLine("  options: --input, --page, --label");
            sb.AppendLine("  example: PDFPublisher.exe find --input=\"file.pdf\" --page=1 --label=labeltext");
            sb.AppendLine();

            sb.AppendLine(" findall - find all text labels in pdf by mask");
            sb.AppendLine("  options: --input, --output, --label");
            sb.AppendLine("  example: PDFPublisher.exe findall --input=\"file.pdf\" --output=\"alllabels.txt\" --label=\"label*with?mask\"");
            sb.AppendLine();

            sb.AppendLine(" pagesizes - display sizes of all pdf pages(width and height divided by ;)");
            sb.AppendLine("  options: --input");
            sb.AppendLine("  example: PDFPublisher.exe pagesizes --input=\"file.pdf\"");
            sb.AppendLine();

            sb.AppendLine(" batchfile - execute list of any commands described above");
            sb.AppendLine("  options: --input");
            sb.AppendLine("  example: PDFPublisher.exe batchfile --input=\"batch.txt\"");
            sb.AppendLine();

            sb.AppendLine(" findpages - find all pages in pdf with text labels by mask");
            sb.AppendLine("  options: --input, --output, --label");
            sb.AppendLine("  example: PDFPublisher.exe findpages --input=\"file.pdf\" --output=\"allpages.txt\" --label=\"label*with?mask\"");
            sb.AppendLine();

            sb.Append("Option description:");

            ht.Heading = "PDFPublisher utility.";
            ht.Copyright = sb.ToString(); // Replace copyright with commands info
            return ht;
        }
    }    
}
