using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PDFPublisher
{
    /// <summary>
    /// Основной класс консольной программы.
    /// Программа реализует операции над pdf файлами.
    /// </summary>
    class Program
    {
        static int Main(string[] args)
        {
            const int RET_OK = 0;
            const int RET_ERR = 1;

            var options = new Options();
            if (CommandLine.Parser.Default.ParseArguments(args, options))
            {
                if (options.Commands.Count == 0)
                {
                    Console.Write(options.GetUsage());
                    return RET_OK;
                }

                try
                {
                    var cmd = options.Commands[0];
                    switch (cmd)
                    {
                        case "combine":
                        {
                            var files = options.Input.Trim('\"').Split(',');
                            var output = options.Output.Trim('\"');
                            PdfOperations.Combine(files, output);
                            break;
                        }
                        case "cut":
                        {
                            var file = options.Input.Trim('\"');
                            var output = options.Output.Trim('\"');
                            var from = options.FromPage;
                            var to = options.ToPage;
                            PdfOperations.Cut(file, output, from, to);
                            break;
                        }
                        case "barcode":
                        {
                            var files = options.Input.Trim('\"').Split(',');
                            var output = options.Output.Trim('\"');
                            var code = options.Code.Trim('\"');
                            string type = options.Type.Trim('\"');
                            int offsetX = options.OffsetX;
                            int offsetY = options.OffsetY;
                            int rotate = options.Rotate;
                            //int type = !string.IsNullOrEmpty(options.Type) ? GetBarCodeType(options.Type) : iTextSharp.text.pdf.Barcode.EAN8;
                            //int offsetX = !string.IsNullOrEmpty(options.Type) ? GetBarCodeType(options.Type) : iTextSharp.text.pdf.Barcode.EAN8;
                            foreach (var file in files)
                            {
                                PdfOperations.BarcodeStamp(file, code, output, type, offsetX, offsetY, rotate);
                            }
                            break;
                        }
                        case "qrcode":
                            {
                                var files = options.Input.Trim('\"').Split(',');
                                var output = options.Output.Trim('\"');
                                var code = options.Code.Trim('\"');
                                int offsetX = options.OffsetX;
                                int offsetY = options.OffsetY;
                                int rotate = options.Rotate;
                                //int type = !string.IsNullOrEmpty(options.Type) ? GetBarCodeType(options.Type) : iTextSharp.text.pdf.Barcode.EAN8;
                                //int offsetX = !string.IsNullOrEmpty(options.Type) ? GetBarCodeType(options.Type) : iTextSharp.text.pdf.Barcode.EAN8;
                                foreach (var file in files)
                                {
                                    PdfOperations.QR_Stamp(file, code, output, offsetX, offsetY, rotate);
                                }
                                break;
                            }
                        case "convert":
                        {
                            var input = options.Input.Trim('\"');
                            var output = options.Output.Trim('\"');
                            PdfOperations.FromImage(input, output);
                            break;
                        }
                        case "image":
                        {
                            var input = options.Input.Trim('\"');
                            var output = options.Output.Trim('\"');
                            var page = options.Page;
                            PdfOperations.ToImage(input, page, output);
                            break;
                        }
                        case "scanbarcode":
                        {
                            var input = options.Input.Trim('\"');
                            var page = options.Page;
                            string barcode = "";
                            string inputExt = System.IO.Path.GetExtension(input);
                            double scale = options.Scale;
                            bool noClean = options.NoClean;

                            bool find = false;
                            

                            if (inputExt == ".pdf")
                            {
                                PdfOperations.ToImage(input, page, input + ".png", (float)scale);
                                if (page == 0)
                                {
                                    // All barcodes
                                    Console.Write("Barcodes: ");
                                    for (int i = 1; ; i++)
                                    {
                                        string imageFileN = PdfOperations.GetFileNameWithIndex(input + ".png", i);
                                        if (!System.IO.File.Exists(imageFileN)) break;
                                        barcode = BarcodeOperations.Scan(imageFileN);

                                        // Для не найденных barcode просто ставим ; чтобы можно было номер страницы сопоставить
                                        if (find == true) Console.Write(";");

                                        if (!string.IsNullOrEmpty(barcode))
                                        {
                                            Console.Write(barcode);
                                            find = true;
                                        }
                                        else
                                        {
                                            Console.Write("-");
                                        }

                                        if (!noClean)
                                            System.IO.File.Delete(imageFileN);
                                    }
                                    Console.WriteLine("");
                                }
                                else
                                {
                                    barcode = BarcodeOperations.Scan(input + ".png");
                                    if (!noClean)
                                        System.IO.File.Delete(input + ".png");
                                    if (!string.IsNullOrEmpty(barcode))
                                    {
                                        Console.WriteLine("Barcode: " + barcode);
                                        find = true;
                                    }
                                }
                            }
                            else
                            {
                                barcode = BarcodeOperations.Scan(input);
                                if (!string.IsNullOrEmpty(barcode))
                                {
                                    Console.WriteLine("Barcode: " + barcode);
                                    find = true;
                                }
                            }

                            if (find==false)
                            {
                                Console.WriteLine(PdfOperations.BARCODE_NOT_FOUND);
                            }
                            break;
                        }
                        case "barcodereplace":
                        {
                            var input = options.Input.Trim('\"');
                            var output = options.Output.Trim('\"');
                            if (input == output)
                            {
                                throw new Exception(PdfOperations.FILE_MUST_BE_DIFFERENT);
                            }
                            var label = options.Label;
                            var code = options.Code;
                            var page = options.Page;
                            bool displayText = !options.NoText;
                            bool barCodeReplaced = PdfOperations.ReplaceBarcode128(input, output, page, label, code, displayText);
                            if (!barCodeReplaced)
                                Console.WriteLine(PdfOperations.BARCODE_NOT_FOUND);

                            break;
                        }
                        case "find":
                        {
                            var input = options.Input.Trim('\"');
                            var page = options.Page;
                            var text = options.Label;

                            var result = PdfOperations.SearchPdfFile(input, text, page);
                            if (result.Count > 0)
                            {
                                foreach (var item in result)
                                {
                                    string pos = string.Format("Label position: {0}, {1}, {2}, {3} (page {4})", item.MinX, item.MinY, item.MaxX, item.MaxY, item.Page);
                                    Console.WriteLine(pos);
                                }
                            }
                            else
                            {
                                Console.WriteLine(PdfOperations.LABEL_NOT_FOUND);
                            }
                            break;
                        }
                        case "barcodeonlabel":
                        {
                            string input = options.Input.Trim('\"');
                            string output = options.Output.Trim('\"');
                            if (input == output)
                            {
                                throw new Exception(PdfOperations.FILE_MUST_BE_DIFFERENT);
                            }

                            string type = options.Type.Trim('\"');
                            string code = options.Code;
                            string text = options.Label;
                            string widthParam = options.Width;
                            string heightParam = options.Height;
                            bool displayText = !options.NoText;
                            double barHeight = options.BarHeight;

                            bool stampReplaced = PdfOperations.PlaceBarcodeOnLabel(input, output, text, type, code, widthParam, heightParam, (float)barHeight, displayText);
                            if(!stampReplaced) 
                                Console.WriteLine(PdfOperations.LABEL_NOT_FOUND);

                            break;
                        }
                        case "imageonlabel":
                        {
                            string input = options.Input.Trim('\"');
                            string output = options.Output.Trim('\"');
                            string imageFile = options.ImageFile.Trim('\"');
                            if (input == output)
                            {
                                throw new Exception(PdfOperations.FILE_MUST_BE_DIFFERENT);
                            }

                            string label = options.Label;
                            string widthParam = options.Width;
                            string heightParam = options.Height;
                            bool center = options.Center;

                            bool labelFinded = PdfOperations.PlaceImageOnLabel(input, output, label, imageFile, widthParam, heightParam, center);
                            if (!labelFinded)
                                Console.WriteLine(PdfOperations.LABEL_NOT_FOUND);
                            break;
                        }
                        case "extractimages":
                        {
                            string input = options.Input.Trim('\"');
                            PdfOperations.ExtractImagesFromPDF(input);
                            break;
                        }
                        case "pagesizes":
                        {
                            string input = options.Input.Trim('\"');
                            IList<System.Drawing.SizeF> sizes = PdfOperations.GetPagesInfo(input);
                            bool firstTime = true;
                            foreach(System.Drawing.SizeF size in sizes) {
                                if (!firstTime)
                                    Console.Write(";");
                                Console.Write("{0};{1}", size.Width, size.Height);
                                firstTime = false;
                            }
                            Console.WriteLine();
                            break;
                        }
                        case "helloworld":
                        {
                            var output = options.Output;
                            if (string.IsNullOrEmpty(output)) output = "helloworld.pdf";
                            else output = output.Trim('\"');
                            PdfOperations.HelloWorld(output);
                            break; 
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    Console.WriteLine(PdfOperations.OPERATION_ERROR);
                    Console.WriteLine(ex.Message);
                    return RET_ERR;
                }
                return RET_OK;
            }
            return RET_ERR;
        }

        static void Test() 
        {
            PdfOperations.HelloWorld(@"d:\PDF1.pdf");

            PdfOperations.FromImage(@"d:\img.png", @"d:\PDF2.pdf");
            PdfOperations.FromImage(@"d:\original.jpg", @"d:\PDF3.pdf");
            // PdfOperations.FromImage(@"d:\test.tiff", @"d:\PDF4.pdf"); не поддерживаются tif с прозрачным слоем (не все tif поддерживаются)
            PdfOperations.FromImage(@"d:\test2.tif", @"d:\PDF4.pdf");
            PdfOperations.Combine(new string[] { @"d:\PDF1.pdf", @"d:\PDF2.pdf", @"d:\PDF3.pdf", @"d:\PDF4.pdf", @"d:\Karyera_menegera.pdf" }, @"d:\PDF5.pdf");

            //PdfOperations.BarcodeStamp(@"d:\PDF5.pdf", "12345670", @"d:\PDF6.pdf");
            PdfOperations.BarcodeStamp(@"d:\PDF5.pdf", "3222231", @"d:\PDF6.pdf");

            PdfOperations.BarcodeStamp(@"d:\Книга-А.Mызникова-С камерой на свободу.pdf", "23342344", @"d:\PDF7.pdf");

            PdfOperations.BarcodeStamp(@"D:\SampleFiles\pdf\Microsoft Private Cloud Licensing Datasheet.pdf", "23342344", @"d:\PDF1-1.pdf");
            PdfOperations.BarcodeStamp(@"D:\SampleFiles\pdf\mostovik_44_hmkh iirlueou.pdf", "23342344", @"d:\PDF1-2.pdf");
            PdfOperations.BarcodeStamp(@"D:\SampleFiles\pdf\Ocean Grill.pdf", "23342344", @"d:\PDF1-3.pdf");
            PdfOperations.BarcodeStamp(@"D:\SampleFiles\pdf\offline.pdf", "23342344", @"d:\PDF1-4.pdf");
            PdfOperations.BarcodeStamp(@"D:\SampleFiles\pdf\Present-KPSB-may2001.pdf", "23342344", @"d:\PDF1-5.pdf");
            PdfOperations.BarcodeStamp(@"D:\SampleFiles\pdf\Q-50-12-4L.pdf", "23342344", @"d:\PDF1-6.pdf");
            PdfOperations.BarcodeStamp(@"D:\SampleFiles\pdf\qashqai360.pdf", "23342344", @"d:\PDF1-7.pdf");
            PdfOperations.BarcodeStamp(@"D:\SampleFiles\pdf\Saga.pdf", "23342344", @"d:\PDF1-8.pdf");
            PdfOperations.BarcodeStamp(@"D:\SampleFiles\pdf\sagas.pdf", "23342344", @"d:\PDF1-9.pdf");
            PdfOperations.BarcodeStamp(@"D:\SampleFiles\pdf\Shema_2_new.pdf", "23342344", @"d:\PDF1-10.pdf");
            PdfOperations.BarcodeStamp(@"D:\SampleFiles\pdf\Sounds Of The Sea.pdf", "23342344", @"d:\PDF1-11.pdf");
            PdfOperations.BarcodeStamp(@"D:\SampleFiles\pdf\Spisok_0_new.pdf", "23342344", @"d:\PDF1-12.pdf");
            PdfOperations.BarcodeStamp(@"D:\SampleFiles\pdf\Tarify_2013_UEK.pdf", "23342344", @"d:\PDF1-13.pdf");
            PdfOperations.BarcodeStamp(@"D:\SampleFiles\pdf\Windows_Azure_Pack_Datasheet.pdf", "23342344", @"d:\PDF1-14.pdf");
            PdfOperations.BarcodeStamp(@"D:\SampleFiles\pdf\Windows_Server_2012_R2_Datasheet.pdf", "23342344", @"d:\PDF1-15.pdf");
            PdfOperations.BarcodeStamp(@"D:\SampleFiles\pdf\ваучер GTA.pdf", "23342344", @"d:\PDF1-16.pdf");
            PdfOperations.BarcodeStamp(@"D:\SampleFiles\pdf\аэропорт Шереметьево.pdf", "23342344", @"d:\PDF1-17.pdf");

            PdfOperations.BarcodeStamp(@"D:\SampleFiles\pdf-dwg\Box-Layout1.pdf", "23342344", @"d:\PDF2-1.pdf");
            PdfOperations.BarcodeStamp(@"D:\SampleFiles\pdf-dwg\test-Layout1.pdf", "23342344", @"d:\PDF2-2.pdf");

            PdfOperations.BarcodeStamp(@"D:\box_box.pdf", "23342344", @"d:\PDF3-1.pdf");
            PdfOperations.BarcodeStamp(@"D:\box-box-nano.pdf", "23342344", @"d:\PDF3-2.pdf");

            PdfOperations.BarcodeStamp(@"D:\box-box-nano.pdf", "23342344", @"d:\PDF3-2.pdf");
            PdfOperations.BarcodeStamp(@"D:\box-nano.pdf", "23342344", @"d:\PDF3-3.pdf");
            PdfOperations.BarcodeStamp(@"D:\зигзаг.pdf", "23342344", @"d:\PDF3-4.pdf");
        }
    }
}
