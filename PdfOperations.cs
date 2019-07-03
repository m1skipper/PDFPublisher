using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using iTextSharp.text.pdf.parser;

namespace PDFPublisher
{
    /// <summary>
    /// Класс реализующий операции над pdf файлами
    /// </summary>
    class PdfOperations
    {
        // Константы, крайне желательно не менять, так как они могут использоваться для анализа результата.
        // (Особенно те, которые используются при возврате значения 0(успех) из программы)
        public const string BARCODE_NOT_FOUND = "Barcode not found.";
        public const string BARCODE_DIFFERENT_FOUND = "Different barcode found: ";
        public const string LABEL_NOT_FOUND = "Label not found.";
        public const string FILE_NOT_FOUND = "File not found";
        public const string FILE_IMAGE_NOT_FOUND = "Image file not found";
        public const string FILE_MUST_BE_DIFFERENT = "Output file name must be different from input file name.";
        public const string OPERATION_ERROR = "Operation error: ";

        /// <summary>
        /// Создает файл pdf с надписью 'HelloWorld'
        /// http://www.c-sharpcorner.com/UploadFile/f2e803/basic-pdf-creation-using-itextsharp-part-i/
        /// </summary>
        /// <param name="args"></param>
        public static void HelloWorld(string fileName)
        {
            System.IO.FileStream fs = new FileStream(fileName, FileMode.Create);

            // Create an instance of the document class which represents the PDF document itself.
            Document document = new Document(PageSize.A4, 25, 25, 30, 30);
            // Create an instance to the PDF file by creating an instance of the PDF 
            // Writer class using the document and the filestrem in the constructor.
            PdfWriter writer = PdfWriter.GetInstance(document, fs);

            // Add meta information to the document
            document.AddAuthor("Document author");
            document.AddCreator("Sample application");
            document.AddKeywords("PDF keywords");
            document.AddSubject("Document subject - Describing the steps creating a PDF document");
            document.AddTitle("The document title");

            // Open the document to enable you to write to the document
            document.Open();
            // Add a simple and wellknown phrase to the document in a flow layout manner
            document.Add(new Paragraph("Hello World!"));
            // Close the document
            document.Close();
            // Close the writer instance
            writer.Close();
            // Always close open filehandles explicity
            fs.Close();
        }

        /// <summary>
        /// Получить информацию о страницах pdf.
        /// </summary>
        /// <param name="fileName">Имя pdf файла</param>
        public static IList<SizeF> GetPagesInfo(string fileName)
        {
            // Через iTextSharp медленне получается
            //List<SizeF> sizes = new List<SizeF>();
            //using (PdfReader reader = new PdfReader(fileName))
            //{
                //for (int i = 0; i < reader.NumberOfPages; i++)
                //{
                    //var rect = reader.GetPageSizeWithRotation(i+1);
                    //sizes.Add(new SizeF(rect.Width, rect.Height));
                //}
            //}
            //return sizes;

            // Поэтому через viewer будем получать.
            using (var document = PdfiumViewer.PdfDocument.Load(fileName))
            {
                var info = document.GetInformation();
                IList<SizeF> sizes = document.PageSizes;
                return sizes;
            }
        }

        /// <summary>
        /// Слияние нескольких pdf файлов в один
        /// </summary>
        /// <param name="fileNames">Массив файлов для слияния</param>
        /// <param name="outFile">Файл в который вывести результат</param>
        public static void Combine(string[] fileNames, string outFile)
        {
            using (FileStream stream = new FileStream(outFile, FileMode.Create))
            {
                Document document = new Document();
                PdfCopy pdf = new PdfCopy(document, stream);
                PdfReader reader = null;
                try
                {
                    document.Open();
                    foreach (string file in fileNames)
                    {
                        reader = new PdfReader(file);

                        pdf.AddDocument(reader);

                        reader.Close();
                    }
                }
                finally
                {
                    if (document != null)
                    {
                        document.Close();
                    }
                }
            }
        }

        /// <summary>
        /// Вырезать несколько страниц из файла pdf в другой файл pdf.
        /// </summary>
        /// <param name="fileNames">Исходный файл</param>
        /// <param name="outFile">Файл в который вывести результат</param>
        /// <param name="outFile">Файл в который вывести результат</param>
        /// <param name="from">Номер страницы начиная с 1 с которого начинать</param>
        /// <param name="to">Номер страницы начиная с 1 по который вырезать, если = 0 то до конца</param>
        public static void Cut(string fileName, string outFile, int from, int to = 0)
        {
            using (FileStream stream = new FileStream(outFile, FileMode.Create))
            {
                Document document = new Document();
                PdfCopy pdf = new PdfCopy(document, stream);
                try
                {
                    document.Open();
                    using(PdfReader reader = new PdfReader(fileName))
                    {
                        var pages = new List<int>();
                        if(from <= 0) from = 1;
                        if(to <= 0) to = reader.Info.Count + 1;
                        if (to > reader.Info.Count + 1) to = reader.Info.Count + 1;
                        for(int i=from;i<=to;i++) pages.Add(i);
                        reader.SelectPages(pages);
                        pdf.AddDocument(reader);
                        reader.Close();
                    }
                }
                finally
                {
                    if (document != null)
                    {
                        document.Close();
                    }
                }
            }
        }

        /// <summary>
        /// Создать файл png из страницы pdf
        /// </summary>
        /// <param name="pdfFile">Исходный файл pdf</param>
        /// <param name="page">Номер страницы</param>
        /// <param name="imageFile">Выходной файл png</param>
        static public void ToImage(string pdfFile, int page, string imageFile, float scale = 1)
        {
            using (var document = PdfiumViewer.PdfDocument.Load(pdfFile))
            {
                var info = document.GetInformation();
                IList<SizeF> sizes = document.PageSizes;
                int width = (int)(sizes[0].Width * scale);
                int height = (int)(sizes[0].Height * scale);

                if (page == 0)
                {
                    // All pages
                    for (int i = 0; i < document.PageCount; i++)
                    {
                        using (var image = document.Render(i, width, height, 300, 300, PdfiumViewer.PdfRenderFlags.Grayscale))
                        {
                            string imageFileN = GetFileNameWithIndex(imageFile, i + 1);
                            image.Save(imageFileN, System.Drawing.Imaging.ImageFormat.Png);
                        }
                    }
                }
                else
                {
                    using (var image = document.Render(page - 1, width, height, 300, 300, PdfiumViewer.PdfRenderFlags.Grayscale))
                    {
                        image.Save(imageFile, System.Drawing.Imaging.ImageFormat.Png);
                    }
                }
            }
        }

        /// <summary>
        /// Создать имя файла с индексом из имени файла.
        /// (Всмомогательная функция)
        /// </summary>
        /// <param name="fileName">Имя файла</param>
        /// <param name="index">Индекс</param>
        /// <returns>Имя с индексом в конце и тем же расширением</returns>
        public static string GetFileNameWithIndex(string fileName, int index)
        {
            return System.IO.Path.Combine(System.IO.Path.GetDirectoryName(fileName), System.IO.Path.GetFileNameWithoutExtension(fileName) + index + System.IO.Path.GetExtension(fileName));
        }

        /// <summary>
        /// Создать файл pdf из файла картинки
        /// </summary>
        /// <param name="imageFile">Файл картинки типа jpg, jpeg, tif, png</param>
        /// <param name="outFile">Выходной файл pdf</param>
        public static void FromImage(string imageFile, string outFile)
        {
            System.IO.FileStream fs = new FileStream(outFile, FileMode.Create);

            iTextSharp.text.Rectangle pageSize = null;
            using (var srcImage = new Bitmap(imageFile))
            {
                pageSize = new iTextSharp.text.Rectangle(0, 0, srcImage.Width, srcImage.Height);
            }

            iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(imageFile);
            //img.ScaleAbsolute(216, 70);
            //img.ScalePercent(50);
            img.SetAbsolutePosition(0, 0);

            // Create an instance of the document class which represents the PDF document itself.
            Document document = new Document(pageSize);
            // Create an instance to the PDF file by creating an instance of the PDF 
            // Writer class using the document and the filestrem in the constructor.
            PdfWriter writer = PdfWriter.GetInstance(document, fs);

            // Add meta information to the document
            //document.AddAuthor("Author");
            //document.AddCreator("Sample application using iTextSharp");
            //document.AddKeywords("PDF tutorial education");
            //document.AddSubject("Document subject - Describing the steps creating a PDF document");
            //document.AddTitle("The document title - PDF creation using iTextSharp");

            // Open the document to enable you to write to the document
            document.Open();

            // Add a simple and wellknown phrase to the document in a flow layout manner
            document.Add(img);

            // Close the document
            document.Close();
            // Close the writer instance
            writer.Close();
            // Always close open filehandles explicity
            fs.Close();
        }

        /// <summary>
        /// Добавить штрихкод в каждую страницу файла pdf
        /// </summary>
        /// <param name="pdfFile">Файл pdf</param>
        /// <param name="code">Значение штрихкода</param>
        /// <param name="outFile">Выходной файл pdf</param>
        /// <param name="type">Тип бар кода, какие поддерживаются см. функцию CreateBarcode</param>
        /// <param name="offsetX">Смещение по X</param>
        /// <param name="offsetY">Смещение по Y</param>
        public static void BarcodeStamp(string pdfFile, string code, string outFile, string type = BarcodeType.Code128, int offsetX = 4, int offsetY = 4, int rotateDegrees = 270)
        {
            System.IO.FileStream fs = new FileStream(outFile, FileMode.Create);
            PdfReader reader = new PdfReader(pdfFile);
            PdfStamper stamper = new PdfStamper(reader, fs);
            int n = reader.NumberOfPages;
            iTextSharp.text.Rectangle pagesize;

            Barcode barcode = CreateBarcode(type, code);
            for (int i = 1; i <= n; i++)
            {
                PdfContentByte over = stamper.GetOverContent(i);

                // GetPageSize - возвращает размер без учета ориентации(альбомная, вертикальная), т.е. в случае альбомной width и height
                // будут перепутаны при отображении на экране и позиционировании, поэтому нужно делать GetPageSizeWithRotation.
                pagesize = reader.GetPageSizeWithRotation(i);

                PdfTemplate template = CreateBarcodeTemplate(over, barcode, true, true);
                
                // Добавить с поворотом
                var matrix = new System.Drawing.Drawing2D.Matrix();
                matrix.Translate(offsetX, offsetY);
                matrix.Rotate(rotateDegrees);

            	float[] elements = matrix.Elements;
                over.AddTemplate(template, elements[0], elements[1], elements[2], elements[3], elements[4], elements[5]);
            }
            stamper.Close();
            reader.Close();
        }

        /// <summary>
        /// Посчитать контрольную цифру
        /// </summary>
        /// <param name="code"></param>
        /// <returns></returns>
        private static int BarCodeEANCheckSum(string code, int len = 8)
        {
            if (code.Length != len - 1) throw new Exception("EAN" + len + "code length must have " + (len - 1).ToString() + " digits");

            int sum = 0;
            for (int i = 0; i < code.Length; i++)
            {
                int codeDigit = code[i] - '0';
                bool odd = ((len - i) % 2 == 0); // Четность с конца определяется. Пример: ("5512345"->"OEOEOEO")
                sum += codeDigit * (odd ? 3 : 1);
            }

            int checksumDigit = 10 - (sum % 10);
            if (checksumDigit == 10) checksumDigit = 0;

            return checksumDigit;
        }

        /// <summary>
        /// Добавить цифру контрольной суммы, если таковой нет. 
        /// Если есть проверить корректность.
        /// </summary>
        /// <param name="code">Значение кода</param>
        /// <param name="len">Размерность кода(сколько символов должно быть)</param>
        /// <returns>Исходную строку если код правильный, либо строку кода с добавленной цифрой контрольной суммы</returns>
        private static string AddEANCheckSum(string code, int len) 
        {
            if(code.Length == len) 
            {
                // Проверим последнюю цифру, что код верный
                int checkSumDigit = BarCodeEANCheckSum(code.Substring(0, len-1), len);
                int codeDigit = code[len - 1] - '0';
                if (checkSumDigit != codeDigit)
                {
                    throw new Exception("Incorrect check sum of EAN" + len + " code.");
                }
                return code;
            }
            else if(code.Length == len-1)
            {
                // Дополним цифрой чек суммы
                return code + BarCodeEANCheckSum(code, len).ToString();
            }
            else 
            {
                throw new Exception("EAN" + len + " code must have " + (len - 1).ToString() + " digits + check sum digit");
            }
        }

        /// <summary>
        /// Создать объект штрих кода по типу и самому коду.
        /// </summary>
        /// <param name="type">Тип кода, поддерживаются Barcode.EAN8, Barcode.EAN13, Barcode.CODE128</param>
        /// <param name="code">Содержание кода</param>
        /// <returns></returns>
        private static Barcode CreateBarcode(string type, string code, bool displayText = true, float barHeight = 0)
        {
            Barcode barcode = null;
            switch (type)
            {
                case BarcodeType.Ean8:
                    {
                        barcode = new BarcodeEAN();
                        barcode.CodeType = Barcode.EAN8;
                        barcode.Code = AddEANCheckSum(code, 8);
                        break;
                    }
                case BarcodeType.Ean13:
                    {
                        barcode = new BarcodeEAN();
                        barcode.CodeType = Barcode.EAN13;
                        barcode.Code = AddEANCheckSum(code, 13);
                        break;
                    }
                case BarcodeType.Code128:
                case BarcodeType.Ean128:
                    {
                        barcode = new Barcode128();
                        barcode.Code = code;
                        barcode.ChecksumText = true;
                        barcode.GenerateChecksum = true;
                        barcode.StartStopText = true;
                        break;
                    }
                default:
                    throw new System.Exception("Not supported barcode type");
            }

            // Установим дополнительные параметры
            if (barHeight > 0)
            {
                barcode.BarHeight = barHeight; // Высота одной полоски
            }
            if (displayText == false)
            {
                barcode.Font = null; // no text
            }
            // barcode.X = 5; // длина 1-й полоски
            // barcode.Size = 8; // размер шрифта

            return barcode;
        }

        public static List<PDFCompare.Comparer.TextItem> SearchPdfFile(string fileName, String searchText, int page = -1)
        {
            if (!File.Exists(fileName))
                throw new FileNotFoundException(FILE_NOT_FOUND, fileName);

            var findedItems = new List<PDFCompare.Comparer.TextItem>();

            using (PdfReader pdfReader = new PdfReader(fileName))
            {
                var strategy = new PDFCompare.Comparer.LocationExtractionStrategy();

                bool oneTime = page > 0;
                if (page == -1)
                    page = 1;

                for (; page <= pdfReader.NumberOfPages; page++)
                {
                    strategy.Page = page;
                    PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);
                    List<PDFCompare.Comparer.TextItem> items = strategy.GetTextItems();
                    foreach (var item in items)
                    {
                        var word = item.Text;
                        if (word == searchText)
                        {
                            findedItems.Add(item);
                        }
                    }
                    if (oneTime) break;
                }
            }
            return findedItems;
        }

        /// <summary>
        /// https://stackoverflow.com/questions/802269/extract-images-using-itextsharp
        /// </summary>
        private static IList<System.Drawing.Image> GetImagesFromPdfDict(PdfDictionary dict, PdfReader doc)
        {
            List<System.Drawing.Image> images = new List<System.Drawing.Image>();
            PdfDictionary res = (PdfDictionary)(PdfReader.GetPdfObject(dict.Get(PdfName.RESOURCES)));
            PdfDictionary xobj = (PdfDictionary)(PdfReader.GetPdfObject(res.Get(PdfName.XOBJECT)));

            if (xobj != null)
            {
                foreach (PdfName name in xobj.Keys)
                {
                    PdfObject obj = xobj.Get(name);
                    if (obj.IsIndirect())
                    {
                        PdfDictionary tg = (PdfDictionary)(PdfReader.GetPdfObject(obj));
                        PdfName subtype = (PdfName)(PdfReader.GetPdfObject(tg.Get(PdfName.SUBTYPE)));
                        if (PdfName.IMAGE.Equals(subtype))
                        {
                            int xrefIdx = ((PRIndirectReference)obj).Number;
                            PdfObject pdfObj = doc.GetPdfObject(xrefIdx);
                            PdfStream str = (PdfStream)(pdfObj);

                            iTextSharp.text.pdf.parser.PdfImageObject pdfImage =
                                new iTextSharp.text.pdf.parser.PdfImageObject((PRStream)str);
                            System.Drawing.Image img = pdfImage.GetDrawingImage();

                            images.Add(img);
                        }
                        else if (PdfName.FORM.Equals(subtype) || PdfName.GROUP.Equals(subtype))
                        {
                            images.AddRange(GetImagesFromPdfDict(tg, doc));
                        }
                    }
                }
            }

            return images;
        }

        public static void ExtractImagesFromPDF(string input)
        {
            string output = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(input), System.IO.Path.GetFileNameWithoutExtension(input));
            using (PdfReader pdf = new PdfReader(input))
            {
                for (int pageNumber = 1; pageNumber <= pdf.NumberOfPages; pageNumber++)
                {
                    PdfDictionary pg = pdf.GetPageN(pageNumber);
                    IList<System.Drawing.Image> images = GetImagesFromPdfDict(pg, pdf);
                    int i = 0;
                    foreach(var image in images) {
                        string outputFile = output + "-" + pageNumber + "-" + i + ".jpg";
                        image.Save(outputFile);
                        i++;
                    }
                }
            }
        }

        static public bool PlaceBarcodeOnLabel(string input, string output, string label, string type, string code, string widthParam, string heightParam, float barHeight, bool displayText)
        {
            if (!File.Exists(input))
                throw new FileNotFoundException(FILE_NOT_FOUND, input);

            // Проверим параметры
            // widthParam, heightParam должны принимать значение "", null, fit, или число.
            bool scaleByWidth = false;
            bool scaleByHeight = false;
            float width = 0;
            float height = 0;
            int result;
            if (string.IsNullOrEmpty(widthParam)) { ; } // пропускаем
            else if (Int32.TryParse(widthParam, out result))
            {
                width = result;
                scaleByWidth = true;
            }
            else if (widthParam == "fit")
            {
                scaleByWidth = true;
            }
            else throw new ArgumentException();

            if (string.IsNullOrEmpty(heightParam)) { ; } // пропускаем
            else if (Int32.TryParse(heightParam, out result))
            {
                height = result;
                scaleByHeight = true;
            }
            else if (heightParam == "fit")
            {
                scaleByHeight = true;
            }
            else throw new ArgumentException();

            System.IO.FileStream fs = new FileStream(output, FileMode.Create);

            Barcode barcode = CreateBarcode(type, code, displayText, barHeight);
            bool stampFinded = false;

            using (PdfReader pdfReader = new PdfReader(input))
            {
                PdfStamper stamper = new PdfStamper(pdfReader, fs);

                for (int page = 1; page <= pdfReader.NumberOfPages; page++)
                {
                    var strategy = new PDFCompare.Comparer.LocationExtractionStrategy();
                    strategy.Page = page;
                    PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);
                    List<PDFCompare.Comparer.TextItem> items = strategy.GetTextItems();

                    PdfContentByte over = stamper.GetOverContent(page);

                    foreach (var item in items)
                    {
                        var word = item.Text;
                        if (word == label)
                        {
                            float rotateDegrees = 0;

                            Vector leftBottom = item.DescentLine.GetStartPoint();
                            Vector rightTop = item.AscentLine.GetEndPoint();

                            float offsetX = leftBottom[0];
                            float offsetY = leftBottom[1];

                            // GetPageSize - возвращает размер без учета ориентации(альбомная, вертикальная), т.е. в случае альбомной width и height
                            // будут перепутаны при отображении на экране и позиционировании, поэтому нужно делать GetPageSizeWithRotation.
                            iTextSharp.text.Rectangle pagesize = pdfReader.GetPageSizeWithRotation(page);

                            PdfTemplate template = CreateBarcodeTemplate(over, barcode, opaque: true, withBorder: true);

                            // Трансформации применяются в обратном порядке
                            //
                            var matrix = new System.Drawing.Drawing2D.Matrix();

                            Vector v = item.DescentLine.GetEndPoint().Subtract(item.DescentLine.GetStartPoint());
                            rotateDegrees = (float)(Math.Atan2((double)v[1], (double)v[0]) * 180 / Math.PI);

                            int pageRotation = pdfReader.GetPageRotation(page);
                            // Если страница повернута, то мы получаем координаты на неповернутой странице,
                            // а AddTemplate делает в координатах повернутой.
                            // Поэтому преобразуем штамп обратно в координаты неповернутой.
                            switch (pageRotation)
                            {
                                case 90: matrix.Translate(0, pagesize.Height); break;
                                case 180: matrix.Translate(pagesize.Width, pagesize.Height); break;
                                case 270: matrix.Translate(pagesize.Width, 0); break;
                            }
                            if (pageRotation != 0)
                                matrix.Rotate(360 - pageRotation);

                            // Сдвинем штрих код на свое место
                            matrix.Translate(offsetX, offsetY);

                            // Повернем штрихкод как текст повернут
                            matrix.Rotate(rotateDegrees);

                            // Отмасштабируем штрихкод чтобы вписать в п/у текста
                            float widthLabel = width != 0 ? width : (item.DescentLine.GetEndPoint().Subtract(item.DescentLine.GetStartPoint())).Length;
                            float heightLabel = height != 0 ? height : (item.AscentLine.GetStartPoint().Subtract(item.DescentLine.GetStartPoint())).Length;

                            float scaleX = scaleByWidth ? widthLabel / barcode.BarcodeSize.Width : 0;
                            float scaleY = scaleByHeight ? heightLabel / barcode.BarcodeSize.Height : 0;
                            if (scaleX != 0 && scaleY == 0) scaleY = scaleX;
                            if (scaleY != 0 && scaleX == 0) scaleX = scaleY;
                            if (scaleX != 0 && scaleY != 0)
                            {
                                matrix.Scale(scaleX, scaleY);
                            }

                            // Применим шаблон и преобразования
                            float[] elements = matrix.Elements;
                            over.AddTemplate(template, elements[0], elements[1], elements[2], elements[3], elements[4], elements[5]);

                            stampFinded = true;
                        }
                    }
                }
                stamper.Close();
            }
            return stampFinded;
        }

        static public bool ReplaceBarcode128(string input, string output, int page, string sourceCode, string newCode, bool displayText, int maxBarCodeCount = 5)
        {
            float scale = 3;
            float dpi = 300;

            bool barCodeFinded = false;

            Barcode barcodeNew = CreateBarcode(BarcodeType.Code128, newCode, displayText);
            BarcodeResult barcode = null;

            using (var fs = new FileStream(output, FileMode.Create))
            using (var pdfReader = new iTextSharp.text.pdf.PdfReader(input))
            using (var document = PdfiumViewer.PdfDocument.Load(input))
            {
                PdfStamper stamper = new PdfStamper(pdfReader, fs);

                var info = document.GetInformation();
                IList<SizeF> sizes = document.PageSizes;

                // Если задана страница, то только в нёё вставим штрихкод, если не задана, то для всех страниц документа
                int count;
                if (page <= 0)
                {
                    page = 0;
                    count = document.PageCount;
                }
                else
                {
                    // В Pdfium нумерация с 0-ля, а у нас приходит с 1-цы.
                    count = page;
                    page -= 1;                    
                }
                for (; page < count; page++)
                {
                    int imageWidth = (int)(sizes[page].Width * scale);
                    int imageHeight = (int)(sizes[page].Height * scale);
                    System.Drawing.Image img = document.Render(page, imageWidth, imageHeight, dpi, dpi, PdfiumViewer.PdfRenderFlags.Grayscale);

                    for (int i = 0; i < maxBarCodeCount; i++)
                    {
                        barcode = BarcodeOperations.ScanCode128(img, input + ".png");
                        if (barcode != null)
                        {
                            barCodeFinded = true;
                            if (barcode.Code == sourceCode)
                            {
                                float offsetX = barcode.Left / scale;
                                float offsetY = barcode.Top / scale;
                                float width = (barcode.Right - barcode.Left) / scale;
                                float rotateDegrees = barcode.Orientation;

                                // Преобразуем к 0-360 и уберем минимальные отклонения
                                // if (rotateDegrees < 0) rotateDegrees += 360;
                                // else if (rotateDegrees >= 360) rotateDegrees -= 360;
                                // if (Math.Abs(rotateDegrees - 0) <= 4) rotateDegrees = 0;
                                // if (Math.Abs(rotateDegrees - 90) <= 4) rotateDegrees = 90;
                                // if (Math.Abs(rotateDegrees - 180) <= 4) rotateDegrees = 180;
                                // if (Math.Abs(rotateDegrees - 270) <= 4) rotateDegrees = 270;

                                iTextSharp.text.Rectangle pagesize = pdfReader.GetPageSizeWithRotation(page + 1);

                                // Поскольку ось координат Y для img,bmp сверху вниз направлена и начало в верхнем левом угле,
                                // а в pdf снизу вверх и в нижним левом угле, то преобразуем координаты
                                offsetY = pagesize.Height - offsetY;
                                rotateDegrees = -rotateDegrees;

                                PdfOperations.PlaceBarcode(stamper, barcodeNew, page + 1, offsetX, offsetY, width, rotateDegrees);
                            }
                            else
                            {
                                Console.WriteLine(BARCODE_DIFFERENT_FOUND + barcode.Code);
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                }
                stamper.Close();
            }

            return barCodeFinded;
        }

        static public void PlaceBarcode(string input, string output, int page, string type, string code, float offsetX, float offsetY, float width, float rotateDegrees = 0, float barHeight = 0, bool displayText = true)
        {
            if (!File.Exists(input))
                throw new FileNotFoundException(FILE_NOT_FOUND, input);

            System.IO.FileStream fs = new FileStream(output, FileMode.Create);

            // По умолчанию страница 1
            if (page <= 0) page = 1;

            Barcode barcode = CreateBarcode(type, code, displayText, barHeight);

            // Установим дополнительные параметры
            if (barHeight > 0)
            {
                barcode.BarHeight = barHeight; // Высота одной полоски
            }
            if (displayText == false)
            {
                barcode.Font = null; // no text
            }
            // barcode.X = 5; // длина 1-й полоски
            // barcode.Size = 8; // размер шрифта

            using (PdfReader pdfReader = new PdfReader(input))
            {
                PdfStamper stamper = new PdfStamper(pdfReader, fs);
                PlaceBarcode(stamper, barcode, page, offsetX, offsetY, width, rotateDegrees, barHeight, displayText);
                stamper.Close();
            }
        }

        static public bool PlaceImageOnLabel(string input, string output, string label, string imgFile, string widthParam, string heightParam, bool center)
        {
            if (!File.Exists(input))
                throw new FileNotFoundException(FILE_NOT_FOUND, input);

            if (!File.Exists(imgFile))
                throw new FileNotFoundException(FILE_IMAGE_NOT_FOUND, imgFile);

            bool labelFinded = false;

            // Проверим параметры
            // widthParam, heightParam должны принимать значение "", null, fit, или число.
            bool scaleByWidth = false;
            bool scaleByHeight = false;
            float width = 0;
            float height = 0;
            int result;
            if (string.IsNullOrEmpty(widthParam)) { ; } // пропускаем
            else if (Int32.TryParse(widthParam, out result))
            {
                width = result;
                scaleByWidth = true;
            }
            else if (widthParam == "fit")
            {
                scaleByWidth = true;
            }
            else throw new ArgumentException();

            if (string.IsNullOrEmpty(heightParam)) { ; } // пропускаем
            else if (Int32.TryParse(heightParam, out result))
            {
                height = result;
                scaleByHeight = true;
            }
            else if (heightParam == "fit")
            {
                scaleByHeight = true;
            }
            else throw new ArgumentException();

            System.IO.FileStream fs = new FileStream(output, FileMode.Create);

            using (PdfReader pdfReader = new PdfReader(input))
            {
                PdfStamper stamper = new PdfStamper(pdfReader, fs);

                iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(imgFile);
                img.SetAbsolutePosition(0, 0);

                for (int page = 1; page <= pdfReader.NumberOfPages; page++)
                {
                    var strategy = new PDFCompare.Comparer.LocationExtractionStrategy();
                    strategy.Page = page;
                    PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);
                    List<PDFCompare.Comparer.TextItem> items = strategy.GetTextItems();

                    PdfContentByte over = stamper.GetOverContent(page);

                    foreach (var item in items)
                    {
                        var word = item.Text;
                        if (word == label)
                        {
                            float rotateDegrees = 0;

                            Vector leftBottom = item.DescentLine.GetStartPoint();
                            Vector rightTop = item.AscentLine.GetEndPoint();

                            float offsetX = leftBottom[0];
                            float offsetY = leftBottom[1];

                            // GetPageSize - возвращает размер без учета ориентации(альбомная, вертикальная), т.е. в случае альбомной width и height
                            // будут перепутаны при отображении на экране и позиционировании, поэтому нужно делать GetPageSizeWithRotation.
                            iTextSharp.text.Rectangle pagesize = pdfReader.GetPageSizeWithRotation(page);

                            PdfTemplate template = over.CreateTemplate(img.Width, img.Height);
                            template.AddImage(img);

                            var matrix = new System.Drawing.Drawing2D.Matrix();

                            // Трансформации применяются в обратном порядке
                            //
                            Vector v = item.DescentLine.GetEndPoint().Subtract(item.DescentLine.GetStartPoint());
                            rotateDegrees = (float)(Math.Atan2((double)v[1], (double)v[0]) * 180 / Math.PI);

                            int pageRotation = pdfReader.GetPageRotation(page);
                            
                            // Если страница повернута, то мы получаем координаты на неповернутой странице,
                            // а AddTemplate делает в координатах повернутой.
                            // Поэтому преобразуем штамп обратно в координаты неповернутой.
                            switch (pageRotation)
                            {
                                case 90: matrix.Translate(0, pagesize.Height); break;
                                case 180: matrix.Translate(pagesize.Width, pagesize.Height); break;
                                case 270: matrix.Translate(pagesize.Width, 0); break;
                            }
                            if (pageRotation != 0)
                                matrix.Rotate(360 - pageRotation);
                          
                            // Сдвинем штрих код на свое место
                            matrix.Translate(offsetX, offsetY);

                            // Повернем штрихкод как текст повернут
                            float w = (item.DescentLine.GetEndPoint().Subtract(item.DescentLine.GetStartPoint())).Length;
                            img.ScaleToFit(w, w*img.Height/img.Width);
                            matrix.Rotate(rotateDegrees);

                            // Отмасштабируем штрихкод чтобы вписать в п/у текста
                            float widthLabel = width != 0 ? width : (item.DescentLine.GetEndPoint().Subtract(item.DescentLine.GetStartPoint())).Length;
                            float heightLabel = height != 0 ? height : (item.AscentLine.GetStartPoint().Subtract(item.DescentLine.GetStartPoint())).Length;

                            float scaleX = scaleByWidth ? widthLabel / img.Width : 0;
                            float scaleY = scaleByHeight ? heightLabel / img.Height : 0;
                            if (scaleX != 0 && scaleY == 0) scaleY = scaleX;
                            if (scaleY != 0 && scaleX == 0) scaleX = scaleY;
                            if (scaleX != 0 && scaleY != 0)
                            {
                                matrix.Scale(scaleX, scaleY);
                            }

                            // Применим шаблон и преобразования
                            float[] elements = matrix.Elements;
                            over.AddTemplate(template, elements[0], elements[1], elements[2], elements[3], elements[4], elements[5]);

                            labelFinded = true;
                        }
                    }
                }
                stamper.Close();
            }
            return labelFinded;
        }

        static private PdfTemplate CreateBarcodeTemplate(PdfContentByte over, Barcode barcode, bool opaque, bool withBorder)
        {
            PdfTemplate template = over.CreateTemplate(0, 0);
            float borderSize = opaque && withBorder ? barcode.X * 3 : 0; // 3 ширины полоски
            if (opaque)
            {
                // Создадим п/у залитый белым цветом
                template.SetColorFill(BaseColor.WHITE);
                template.Rectangle(-borderSize, -borderSize, barcode.BarcodeSize.Width + borderSize * 2, barcode.BarcodeSize.Height + borderSize * 2);
                template.Fill();
            }

            iTextSharp.text.Rectangle boundingBox = barcode.PlaceBarcode(template, BaseColor.BLACK, BaseColor.BLACK);
            if (borderSize!=0.0)
            {
                boundingBox.Right += borderSize;
                boundingBox.Left -= borderSize;
                boundingBox.Top += borderSize;
                if (barcode.Font!=null)
                {
                    // Если текст отображается, то не нужно расширять, потому что там и так место под текстом есть.
                    boundingBox.Bottom -= borderSize;
                }
            }
            // Обрежем шаблон до размера barCode + граница
            template.BoundingBox = boundingBox;
            return template;
        }

        static private PdfTemplate CreateImageTemplate(PdfContentByte over, string imageFile)
        {
            iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(imageFile);
            img.SetAbsolutePosition(0, 0);
            PdfTemplate template = over.CreateTemplate(img.Width, img.Height);
            template.AddImage(img);
            return template;
        }

        static private void PlaceBarcode(PdfStamper stamper, Barcode barcode, int page, float left, float top, float width, float rotateDegrees = 0, float barHeight = 0, bool displayText = true, bool withBorder = true)
        {
            // Установим дополнительные параметры
            if (barHeight > 0)
            {
                barcode.BarHeight = barHeight; // Высота одной полоски
            }
            if (displayText == false)
            {
                barcode.Font = null; // no text
            }
            // barcode.X = 5; // длина 1-й полоски
            // barcode.Size = 8; // размер шрифта

            PdfContentByte over = stamper.GetOverContent(page);
            PdfTemplate template = CreateBarcodeTemplate(over, barcode, opaque:true, withBorder:true);

            // Поместим в нужное место документа.
            // Трансформации применяются в обратном порядке.
            //

            var matrix = new System.Drawing.Drawing2D.Matrix();
            float scale = width / barcode.BarcodeSize.Width;

            // Сдвинем штрих код на свое место
            matrix.Translate(left, top);

            // Повернем штрихкод как текст повернут
            matrix.Rotate(rotateDegrees);

            // Отмасштабируем штрихкод чтобы вписать в п/у текста
            matrix.Scale(scale, scale);

            // Сдвинем штрих на величину высоты, чтобы поворот происходил вокруг верхней левой точки, а не вокруг нижней левой.
            matrix.Translate(0, -barcode.BarcodeSize.Height);

            // Применим шаблон и преобразования
            float[] elements = matrix.Elements;
            over.AddTemplate(template, elements[0], elements[1], elements[2], elements[3], elements[4], elements[5]);
        }
    }
}
