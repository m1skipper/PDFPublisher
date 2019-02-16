using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using ZXing;
using ZXing.Common;

namespace PDFPublisher
{
    /// <summary>
    /// Класс содержащий информацию о распознанном штрихкоде
    /// </summary>
    class BarcodeResult
    {
        public string Code;
        public string Type;
        public float Orientation;
        public int Top;
        public int Left;
        public int Right;
        public int Bottom;
    };

    /// <summary>
    /// Класс реализующий операции над pdf файлами
    /// </summary>
    class BarcodeOperations
    {
        /// <summary>
        /// Сканировать штрихкод
        /// </summary>
        /// <param name="fileName">Имя файла с картинкой png, jpeg</param>
        /// <param name="format">BarcodeFormat, если 0, то распознавать все поддерживаемые форматы</param>
        public static string Scan(string fileName, int format = 0)
        {
            string barcode = "";            

            using(Image img = Image.FromFile(fileName))
            using (var bmp = new Bitmap(img))
            {
                var reader = new BarcodeReader();
                reader.Options.TryHarder = true;
                reader.AutoRotate = true;
                if (format != 0)
                {
                    reader.Options.PossibleFormats = new BarcodeFormat[] { (BarcodeFormat)format };
                }

                Result result = reader.Decode(bmp);
                if (result != null)
                {
                    barcode = result.Text;
                }
                return barcode;
            }
        }

        /// <summary>
        /// Найти штрихкод code128 и его ограничивающий п/у
        /// </summary>
        /// <param name="fileName">Имя файла с картинкой png, jpeg</param>
        public static BarcodeResult ScanCode128(Image img, string debugFileName = null)
        {
            BarcodeResult barcodeResult = null;

            Bitmap bmp = new Bitmap(img);

            // create a barcode reader instance
            var reader = new BarcodeReader();
            reader.Options.ReturnCodabarStartEnd = true;
            reader.Options.TryHarder = true;
            reader.Options.PossibleFormats = new BarcodeFormat[] { BarcodeFormat.CODE_128 };
            reader.AutoRotate = true;

            // Сейчас код ищет только один штрихкод.
            // Найти несколько штрихкодов можно с помощью метода DecodeMultiple, но по умолчанию возвращаются только разные штрихкоды.
            // Если мы хотим получить положение одного и того же штрихкода вставленного несколько раз, то нужно будет переписывать метод DecodeMultiple.
            // Есть ещё проблема с разноповернутыми штрихкодами если их несколько, то распознаются только однонаправленные,
            // возможно стоит вручную поворачивать и перезапускать алгоритм.
            //
            int scannerOrientation = 0;
            var result = reader.Decode(bmp);
            if (result != null)
            {
                barcodeResult = new BarcodeResult();
                barcodeResult.Code = result.Text;
                barcodeResult.Type = BarcodeType.ToString(result.BarcodeFormat);
                try
                {
                    var orient = result.ResultMetadata[ResultMetadataType.ORIENTATION];
                    scannerOrientation = (int)orient; // only vertical/horizontal
                }
                catch (Exception /*ex*/)
                {
                }

                if (result.ResultPoints.Length >= 2)
                {
                    //
                    var luminanceSource = new ZXing.BitmapLuminanceSource(bmp);
                    var binarizer = new ZXing.Common.HybridBinarizer(luminanceSource);
                    var bitMatrix = binarizer.BlackMatrix;

                    // result.ResultPoints - точкии линии сканера, для которой был распознан штрих код
                    // для barcode128 эти точки находятся внутри штрихкода видимо за start и stop элементами.
                    // нам нужно получить по этим точкам ограничивающий п/у для всего штрихкода
                    // но это точки в координатах повернутого листа на scannerOrientation градусов,
                    // поэтому преобразуем их в начальную и конечную координату сканера в координатак картинки img
                    Point[] scanLine = GetScanLine(result.ResultPoints, scannerOrientation, img.Width, img.Height, bitMatrix);

                    // Возьмем конечную точку скан линии
                    int x = scanLine[1].X;
                    int y = scanLine[1].Y;                    

                    // Вычислим магический белый квадрат - предположительное место штрихкода.
                    // Искать будет с точки x,y, возвращает 4 точки, типа ограничивающего штрих код белую п/у.
                    // Использующийся алгоритм не до конца понятен, возвращает п/у содержащий одну, несколько полосок, или весь штрих код.
                    // Поэтому единственная полезная информация которая может быть извлечена высота штрихкода (координаты одной вертикальной полосы)
                    var whiterectDetector = ZXing.Common.Detector.WhiteRectangleDetector.Create(bitMatrix, 1, x, y);
                    var whiteRect = whiterectDetector.detect();
                    if (whiteRect == null) return barcodeResult; // не удалось вычислить ограничивающий п/у.

                    // Посчитаем длину первой черной полоски после правой конечной точки scan линии.
                    int blackLineWidth = CalcBlackLineWidth(bitMatrix, scanLine);

                    // Вычислим по имеющимся данным ограничивающий п/у для штрихкода
                    int dx = blackLineWidth*3; // 3 полоски по бокам добавим (для 128 баркода)
                    ResultPoint[] barcodeRect = GetBarCodeRect(scanLine, whiteRect, dx);

                    // Вычислим п/у для создания штрихкода и его ориентацию(поворот).
                    // (п/у без поворота, от левой нижней точки и его угол поворота в градусах)
                    //barcodeResult.Left = (int)barcodeRect[1].X;
                    //barcodeResult.Bottom = (int)barcodeRect[1].Y;
                    //barcodeResult.Right = barcodeResult.Left + (int)ResultPoint.distance(barcodeRect[0], barcodeRect[2]);
                    //barcodeResult.Top = barcodeResult.Bottom - (int)ResultPoint.distance(barcodeRect[0], barcodeRect[1]);
                    barcodeResult.Left = (int)barcodeRect[0].X;
                    barcodeResult.Top = (int)barcodeRect[0].Y;
                    barcodeResult.Right = barcodeResult.Left + (int)ResultPoint.distance(barcodeRect[0], barcodeRect[2]);
                    barcodeResult.Bottom = barcodeResult.Top + (int)ResultPoint.distance(barcodeRect[0], barcodeRect[1]);
                    
                    /*                    }
                                                            else if (scannerOrientation == 180)
                                                            {
                                                                barcodeResult.Left = (int)barcodeRect[0].X;
                                                                barcodeResult.Bottom = (int)barcodeRect[0].Y;
                                                                barcodeResult.Right = barcodeResult.Left + (int)ResultPoint.distance(barcodeRect[0], barcodeRect[2]);
                                                                barcodeResult.Top = barcodeResult.Bottom - (int)ResultPoint.distance(barcodeRect[0], barcodeRect[1]);
                                                            }*/

                    //barcodeResult.Orientation = -Orientation(barcodeRect[0], barcodeRect[2]);
                    barcodeResult.Orientation = Orientation(barcodeRect[0], barcodeRect[2]);
                    
                    if (!string.IsNullOrEmpty(debugFileName))
                    {
                        // Закрасим область белым, чтобы можно было ещё искать barcode
                        var g = Graphics.FromImage(img);
                        g.FillPolygon(new SolidBrush(Color.Pink), ToDrawingRect(barcodeRect));

                        DebugSaveImage(debugFileName,
                            img,
                            barcodeRect,
                            whiteRect,
                            scanLine);
                    }
                }
            }
            return barcodeResult;
        }

        /// <summary>
        /// Сканер сканирует горизонтально поворачивая лист на 90 градусов и 
        /// проблема в том что он возвращает координаты в координатах повернутого листа.
        /// Поэтому преобразуем полученные координаты к координатам исходного листа в зависимости от поворота сканера.
        /// Причем особенность скан линии, что она направленная и начало штрих код у первой точки, а конец у второй.
        /// </summary>
        /// <param name="resultPoint">Координаты в повернутом на scannerOrientation градусов листе</param>
        /// <param name="scannerOrientation">На сколько повернут лист</param>
        /// <returns>Координаты в исходном листе до поворота</returns>
        private static Point[] GetScanLine(ResultPoint[] resultPoints, int scannerOrientation, int width, int height, BitMatrix bitMatrix)
        {
            var scanLine = new Point[2];
            if (scannerOrientation == 90 || scannerOrientation == 270)
            {
                scanLine[0] = new Point(width - (int)resultPoints[0].Y, (int)resultPoints[0].X);
                scanLine[1] = new Point(width - (int)resultPoints[1].Y, (int)resultPoints[1].X);
            }
            else
            {
                scanLine[0] = new Point((int)resultPoints[0].X, (int)resultPoints[0].Y);
                scanLine[1] = new Point((int)resultPoints[1].X, (int)resultPoints[1].Y);
            }

            // resultPoints во float формате и они не всегда попадают на линию штрих кода, могут быть очень рядом со штрихкодом
            // исправим этот недостаток притянув точку к ближайшей черной линии(штрих коду)
            // TODO: (пока по простому по диагонали ищем и по прямым) сделать по спирали
            int x = scanLine[1].X;
            int y = scanLine[1].Y;
            if (bitMatrix[x, y] == false) 
            {
                for (int dx = 0; dx < 10; dx++)
                {
                    if(bitMatrix[x + dx, y]) {
                        scanLine[1].X += dx;
                        scanLine[0].X += dx;
                        break;
                    }
                    else if(bitMatrix[x - dx, y]) {
                        scanLine[1].X -= dx;
                        scanLine[0].X -= dx;
                        break;
                    }
                    else if(bitMatrix[x, y + dx]) {
                        scanLine[1].Y += dx;
                        scanLine[0].Y += dx;
                        break;
                    }
                    else if(bitMatrix[x, y - dx]) {
                        scanLine[1].Y -= dx;
                        scanLine[0].Y -= dx;
                        break;
                    }
                    else if(bitMatrix[x + dx, y + dx]) {
                        scanLine[1].X += dx;
                        scanLine[0].X += dx;
                        scanLine[1].Y += dx;
                        scanLine[0].Y += dx;
                        break;
                    }
                    else if(bitMatrix[x - dx, y - dx]) {
                        scanLine[1].X -= dx;
                        scanLine[0].X -= dx;
                        scanLine[1].Y -= dx;
                        scanLine[0].Y -= dx;
                        break;
                    }
                    else if(bitMatrix[x + dx, y - dx]) {
                        scanLine[1].X += dx;
                        scanLine[0].X += dx;
                        scanLine[1].Y -= dx;
                        scanLine[0].Y -= dx;
                        break;
                    }
                    else if(bitMatrix[x - dx, y + dx]) {
                        scanLine[1].X -= dx;
                        scanLine[0].X -= dx;
                        scanLine[1].Y += dx;
                        scanLine[0].Y += dx; 
                        break;
                    }
                }
            }
           
            return scanLine;
        }

        private static int CalcBlackLineWidth(ZXing.Common.BitMatrix bitMatrix, Point[] scanLine)
        {
            int blackLineWidth = 0;
            int i = scanLine[1].X;
            int j = scanLine[1].Y;

            Point scanLineVector = GetScanLineVector(scanLine);

            // черную посчитаем
            for (; i >= 0 && i < bitMatrix.Width && j >= 0 && j < bitMatrix.Height; i += scanLineVector.X, j += scanLineVector.Y)
            {
                if (bitMatrix[i, j] == false) break;
                blackLineWidth++;
            }
            return blackLineWidth;
        }

        private static Point GetScanLineVector(Point[] scanLine)
        {
            ResultPoint vScan = new ResultPoint(scanLine[1].X - scanLine[0].X, scanLine[1].Y - scanLine[0].Y);
            int dx = Math.Sign(vScan.X);
            int dy = Math.Sign(vScan.Y);
            if (dx != 0 && dy != 0)
            {
                // Линия диагональная, такого быть не должно, но так как координаты float всё может быть,
                // тогда чья линия длиннее тот и прав
                if (Math.Abs(vScan.X) < Math.Abs(vScan.Y)) dx = 0;
                else dy = 0;
            }
            return new Point(dx, dy);
        }

        /// <summary>
        /// Разворачивает п/у по направлению скан линии
        /// </summary>
        private static ResultPoint[] RotateRect(ResultPoint[] whiteRect, Point[] scanLine)
        {
            Point scanVector = GetScanLineVector(scanLine);
            if (scanVector.Y < 0) // orientation 270
            {
                return new ResultPoint[] {
                    new ResultPoint(whiteRect[1].X, whiteRect[1].Y),
                    new ResultPoint(whiteRect[3].X, whiteRect[3].Y),
                    new ResultPoint(whiteRect[0].X, whiteRect[0].Y),
                    new ResultPoint(whiteRect[2].X, whiteRect[2].Y),
                };
            }
            else if (scanVector.X < 0) // orientation 180
            {
                return new ResultPoint[] {
                    new ResultPoint(whiteRect[3].X, whiteRect[3].Y),
                    new ResultPoint(whiteRect[2].X, whiteRect[2].Y),
                    new ResultPoint(whiteRect[1].X, whiteRect[1].Y),
                    new ResultPoint(whiteRect[0].X, whiteRect[0].Y),
                };
            }
            else if (scanVector.Y > 0) // orientation 90
            {
                return new ResultPoint[] {
                    new ResultPoint(whiteRect[2].X, whiteRect[2].Y),
                    new ResultPoint(whiteRect[0].X, whiteRect[0].Y),
                    new ResultPoint(whiteRect[3].X, whiteRect[3].Y),
                    new ResultPoint(whiteRect[1].X, whiteRect[1].Y),
                };
            }
            else // orientation 0
            {
                return whiteRect;
            }
        }

        private static ResultPoint[] GetBarCodeRect(Point[] scanLine, ResultPoint[] whiterect, float dx) 
        {
            ResultPoint pt0 = new ResultPoint(scanLine[0].X, scanLine[0].Y); // Начало линии сканирования (заканчивается за 3 штрихполоски до начала штрихкода)
            ResultPoint pt1 = new ResultPoint(scanLine[1].X, scanLine[1].Y); // Конец линии сканирования (заканчивается за 3 штрихполоски до конца штрихкода)

            // whiterect п/у ограничивающий одну или несколько полосок штрих кода,
            // сделаем его направленным, повернем вдоль скан линии, чтобы можно было ширину расширить
            whiterect = RotateRect(whiterect, scanLine);

            // Проекция точки x,y на найденную полоску штрихкода.
            ResultPoint proj = Projection(pt1, whiterect[1], whiterect[0]);
            float vLen = ResultPoint.distance(pt1, proj);
            ResultPoint v = new ResultPoint(pt1.X - proj.X, pt1.Y - proj.Y);

            // Вектор b перпендикулярный вектору a против часовой стрелки равен  b = (-ay, ax).
            // Найдем правую нормаль к вертикальной полоске
            float height = ResultPoint.distance(whiterect[1], whiterect[0]);
            ResultPoint vOrto = new ResultPoint(-(whiterect[1].Y - whiterect[0].Y) / height, (whiterect[1].X - whiterect[0].X) / height);

            var lastBarCodeLine = new ResultPoint[2] { 
                new ResultPoint(whiterect[0].X + v.X - vOrto.X*dx, whiterect[0].Y + v.Y - vOrto.Y*dx),
                new ResultPoint(whiterect[1].X + v.X - vOrto.X*dx, whiterect[1].Y + v.Y - vOrto.Y*dx)
            };

            ResultPoint proj0 = Projection(pt0, whiterect[1], whiterect[0]);
            ResultPoint v0 = new ResultPoint(pt0.X - proj0.X, pt0.Y - proj0.Y);
            var firstBarCodeLine = new ResultPoint[2] { 
                new ResultPoint(whiterect[0].X + v0.X + vOrto.X*dx, whiterect[0].Y + v0.Y + vOrto.Y*dx),
                new ResultPoint(whiterect[1].X + v0.X + vOrto.X*dx, whiterect[1].Y + v0.Y + vOrto.Y*dx)
            };

            return new ResultPoint[] {
                firstBarCodeLine[0],
                firstBarCodeLine[1],
                lastBarCodeLine[0],
                lastBarCodeLine[1]
            };
        }

        /// <summary>
        /// Угол поворота вектора p1,p2 в градусах
        /// </summary>
        /// <param name="p1">Начало вектора</param>
        /// <param name="p2">Конец вектора</param>
        /// <returns></returns>
        private static float Orientation(ResultPoint p1, ResultPoint p2)
        {
            var v = new ResultPoint(p2.X - p1.X, p2.Y - p1.Y);
            double angle = Math.Atan2(v.Y, v.X)*180/Math.PI;
            return (float)angle;
        }

        /// <summary>
        /// Проекция точки s на прямую ab
        /// </summary>
        private static ResultPoint Projection(ResultPoint p, ResultPoint a, ResultPoint b)
        {
            double k = ((p.X - a.X) * (b.X - a.X) + (p.Y - a.Y) * (b.Y - a.Y)) / ((b.X - a.X) * (b.X - a.X) + (b.Y - a.Y) * (b.Y - a.Y));
            return new ResultPoint((float)(a.X + (b.X - a.X) * k), (float)(a.Y + (b.Y - a.Y) * k));
        }
        
        /// <summary>
        /// Сохранить отладочную информацию об ограничивающих п/у в графический файл
        /// </summary>
        private static void DebugSaveImage(string filename, Image img, ResultPoint[] bounds, ResultPoint[] whiteBar, Point[] scanLine) 
        {
            var g = Graphics.FromImage(img);

            PointF[] points;

            //points = bounds.Select<ResultPoint, PointF>(p => { return new PointF(p.X, p.Y); }).ToArray();
            points = ToDrawingRect(bounds);
            g.DrawLines(new Pen(Color.Red, 3), points);

            points = ToDrawingRect(whiteBar);
            g.DrawLines(new Pen(Color.Blue, 3), points);            

            g.DrawLines(new Pen(Color.Green, 3), scanLine);

            img.Save(filename);
        }

        /// <summary>
        /// Поменяем местами вершины и замкнем п/у
        /// </summary>
        /// <param name="pts"></param>
        /// <returns></returns>
        private static PointF[] ToDrawingRect(ResultPoint[] pts)
        {
            var points = new List<PointF>();
            points.Add(new PointF(pts[0].X, pts[0].Y));
            points.Add(new PointF(pts[1].X, pts[1].Y));
            points.Add(new PointF(pts[3].X, pts[3].Y));
            points.Add(new PointF(pts[2].X, pts[2].Y));
            points.Add(new PointF(pts[0].X, pts[0].Y));
            return points.ToArray();
        }
    }
}
