using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using ZXing;

namespace PDFPublisher
{
    /// <summary>
    /// Класс типов штрихкодов поддерживаемый приложением.
    /// Обеспечивает синхронизацию типов штрихкодов приложения, распознавалки штрихкодов ZXing и печати штрихкодов iTextPdf.
    /// </summary>
    public class BarcodeType
    {
        public const string Ean8 = "ean8";
        public const string Ean13 = "ean13";
        public const string Ean128 = "ean128";
        public const string Code128 = "code128";

        /// <summary>
        /// Поличить тип штрих кода из текстового описания
        /// </summary>
        /// <param name="barCode">Текстовый параметр - ean8, ean13, code128</param>
        static public int GetBarcodeType(string barCode) 
        {
            switch(barCode)
            {
                case BarcodeType.Ean8: return iTextSharp.text.pdf.Barcode.EAN8;
                case BarcodeType.Ean13: return iTextSharp.text.pdf.Barcode.EAN13;
                case BarcodeType.Ean128:
                case BarcodeType.Code128:
                    return iTextSharp.text.pdf.Barcode.CODE128;
            }
            throw new Exception("Not supported barcode: " + barCode);
        }

        /// <summary>
        /// Поличить тип штрих кода ZXing
        /// </summary>
        /// <param name="barCode">Текстовый параметр - ean8, ean13, code128</param>
        static public BarcodeFormat GetBarcodeFormat(string barCode)
        {
            switch (barCode)
            {
                case BarcodeType.Ean8: return BarcodeFormat.EAN_8;
                case BarcodeType.Ean13: return BarcodeFormat.EAN_13;
                case BarcodeType.Ean128:
                case BarcodeType.Code128:
                    return BarcodeFormat.CODE_128;
            }
            throw new Exception("Not supported barcode: " + barCode);
        }

        static public string ToString(BarcodeFormat barcodeFormat)
        {
            switch (barcodeFormat)
            {
                case BarcodeFormat.EAN_8:
                    return BarcodeType.Ean8;
                case BarcodeFormat.EAN_13:
                    return BarcodeType.Ean13;
                case BarcodeFormat.CODE_128:
                    return BarcodeType.Code128;
                default: 
                    return barcodeFormat.ToString();
            }
        }        
    };
}
