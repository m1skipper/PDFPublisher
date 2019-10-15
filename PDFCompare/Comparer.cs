using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Globalization;


namespace PDFCompare
{
	class Comparer
	{
		const float INSERT_PAGE_FACTOR = 0.75f;
		public Settings Settings { get; set; }
		//--------------------------------------------------------------------------------------------------
		public Comparer()
		{
			Settings = new Settings();
		}
		//--------------------------------------------------------------------------------------------------
		public int Compare(string stPDF1, string stPDF2, string stPDFResult)
		{
			int i;
			try
			{
				PdfReader reader1 = new PdfReader(stPDF1);
				reader1.ConsolidateNamedDestinations();
				//ExtractionStrategy strategy = new ExtractionStrategy();
				LocationExtractionStrategy strategy = new LocationExtractionStrategy();
				for(i = 1; i <= reader1.NumberOfPages; i++)
				{
					strategy.Page = i;
					PdfTextExtractor.GetTextFromPage(reader1, i, strategy);
				}
				List<TextItem> items1 = strategy.GetTextItems();
				strategy.Reset();
				PdfReader reader2 = new PdfReader(stPDF2);
				reader2.ConsolidateNamedDestinations();
				for(i = 1; i <= reader2.NumberOfPages; i++)
				{
					strategy.Page = i;
					PdfTextExtractor.GetTextFromPage(reader2, i, strategy);
				}
				List<TextItem> items2 = strategy.GetTextItems();
				//...
				StringBuilder sb1 = new StringBuilder();
				for(i = 0; i < items1.Count; i++)
				{
					if(i != 0)
						sb1.Append('\n');
					sb1.Append(items1[i].Text);
				}
				StringBuilder sb2 = new StringBuilder();
				for(i = 0; i < items2.Count; i++)
				{
					if(i != 0)
						sb2.Append('\n');
					sb2.Append(items2[i].Text);
				}
				Diff diff = new Diff();
				Diff.Item[] diffitems = diff.DiffText(sb1.ToString(), sb2.ToString());
				//...
				MemoryStream ms1 = new MemoryStream();
				MemoryStream ms2 = new MemoryStream();
				using(PdfStamper stamper1 = new PdfStamper(reader1, ms1, '\0', true))
				{
					using(PdfStamper stamper2 = new PdfStamper(reader2, ms2, '\0', true))
					{
						AddHeader(stamper1, reader1, string.Format("Отличий обнаружено: {0}", diffitems.Length));
						//...
						int iPageAdded1 = 0;
						int iPageAdded2 = 0;
						for(i = 0; i < diffitems.Length; i++)
						{
							Diff.Item diffitem = diffitems[i];
							int iColor = diffitem.deletedA == 0 ? Settings.ColorAdded : (diffitem.insertedB == 0 ? Settings.ColorDeleted : Settings.ColorChanged);
							BaseColor color = new BaseColor(iColor & 0x000000FF, (iColor & 0x0000FF00) >> 8, (iColor & 0x00FF0000) >> 16);
							string stComment = string.Format("Отличие {0} из {1}", i + 1, diffitems.Length);
							MarkText(items1, diffitem.StartA, diffitem.deletedA, stamper1, color, iPageAdded1, stComment);
							MarkText(items2, diffitem.StartB, diffitem.insertedB, stamper2, color, iPageAdded2, stComment);
							// Если фрагменты начинаются на разных страницах, то выравниваем
							while(IntexOrLast(items1, diffitem.StartA).Page + iPageAdded1 < IntexOrLast(items2, diffitem.StartB).Page + iPageAdded2)
							{
								if((IntexOrLast(items2, diffitem.StartB).Page + iPageAdded2) - (IntexOrLast(items1, diffitem.StartA).Page + iPageAdded1) == 1)
								{
									Rectangle rectPage1 = reader1.GetPageSizeWithRotation(IntexOrLast(items1, diffitem.StartA).Page + iPageAdded1);
									Rectangle rectPage2 = reader2.GetPageSizeWithRotation(IntexOrLast(items2, diffitem.StartB).Page + iPageAdded2);
									if(INSERT_PAGE_FACTOR > IntexOrLast(items1, diffitem.StartA).MaxY / rectPage1.Height + ((rectPage2.Height - IntexOrLast(items2, diffitem.StartB).MaxY)) / rectPage2.Height)
										break;
								}
								int iNewPage = IntexOrLast(items1, diffitem.StartA).Page;
								if(iNewPage == reader1.NumberOfPages)
									iNewPage++;
								InsertPage(stamper1, iNewPage, reader1.GetPageSizeWithRotation(IntexOrLast(items1, diffitem.StartA).Page + iPageAdded1));
								iPageAdded1++;
							}
							while(IntexOrLast(items1, diffitem.StartA).Page + iPageAdded1 > IntexOrLast(items2, diffitem.StartB).Page + iPageAdded2)
							{
								if((IntexOrLast(items1, diffitem.StartA).Page + iPageAdded1) - (IntexOrLast(items2, diffitem.StartB).Page + iPageAdded2) == 1)
								{
									Rectangle rectPage1 = reader1.GetPageSizeWithRotation(IntexOrLast(items1, diffitem.StartA).Page + iPageAdded1);
									Rectangle rectPage2 = reader2.GetPageSizeWithRotation(IntexOrLast(items2, diffitem.StartB).Page + iPageAdded2);
									if(INSERT_PAGE_FACTOR > ((rectPage1.Height - IntexOrLast(items1, diffitem.StartA).MaxY) / rectPage1.Height) + IntexOrLast(items2, diffitem.StartB).MaxY / rectPage2.Height)
										break;
								}
								int iNewPage = IntexOrLast(items2, diffitem.StartB).Page;
								if(iNewPage == reader2.NumberOfPages)
									iNewPage++;
								InsertPage(stamper2, IntexOrLast(items2, diffitem.StartB).Page, reader2.GetPageSizeWithRotation(IntexOrLast(items2, diffitem.StartB).Page + iPageAdded2));
								iPageAdded2++;
							}
						}
						// Выравниваем количество страниц
						while(reader1.NumberOfPages < reader2.NumberOfPages)
							InsertPage(stamper1, reader1.NumberOfPages + 1, reader1.GetPageSizeWithRotation(reader1.NumberOfPages - 1));
						while(reader1.NumberOfPages > reader2.NumberOfPages)
							InsertPage(stamper2, reader2.NumberOfPages + 1, reader2.GetPageSizeWithRotation(reader2.NumberOfPages - 1));
					}
				}
				reader1 = new PdfReader(ms1.ToArray());
				reader2 = new PdfReader(ms2.ToArray());
				//...
				using(FileStream stream = new FileStream(stPDFResult, FileMode.Create))
				{
					Document document = new Document();
					PdfCopy pdf = new PdfCopy(document, stream);
					document.Open();
					for(i = 1; i <= reader1.NumberOfPages; i++)
					{
						PdfImportedPage page = pdf.GetImportedPage(reader1, i);
						pdf.AddPage(page);
						page = pdf.GetImportedPage(reader2, i);
						pdf.AddPage(page);
					}
					reader1.Close();
					reader2.Close();
					document.Close();
				}
			}
			catch(Exception /*e*/)
			{
				return 1;
			}
			return 0;
		}
		//--------------------------------------------------------------------------------------------------
		void MarkText(List<TextItem> items, int iPosition, int iCount, PdfStamper stamper, BaseColor color, int iPageAdd = 0, string stComment = null)
		{
			const string stTittle = "PDFCompare";
			if(iCount == 0)
			{
				float rWidth, rPosition;
				if(iPosition >= items.Count)
				{
					iPosition = items.Count - 1;
					rWidth = (items[iPosition].MaxY - items[iPosition].MinY) / 5;
					rPosition = items[iPosition].MaxX;
				}
				else
				{
					rWidth = (items[iPosition].MaxY - items[iPosition].MinY) / 5;
					rPosition = items[iPosition].MinX - rWidth;
				}
				Rectangle rect = new Rectangle(rPosition, items[iPosition].MaxY, rPosition + rWidth, items[iPosition].MinY);
				float[] quad = { rect.Left, rect.Bottom, rect.Right, rect.Bottom, rect.Left, rect.Top, rect.Right, rect.Top };
				PdfAnnotation highlight = PdfAnnotation.CreateMarkup(stamper.Writer, rect, stComment, PdfAnnotation.MARKUP_HIGHLIGHT, quad);
				highlight.Color = color;
				highlight.Title = stTittle;
				stamper.AddAnnotation(highlight, items[iPosition].Page + iPageAdd);
			}
			else
			{
				while(iCount > 0 && items[iPosition].IsImage)
				{
					TextItem item = items[iPosition];
					Rectangle rect = new Rectangle(item.MinX, item.MaxY, item.MaxX, item.MinY);
					PdfArray vertices = new PdfArray();
					float rHalfWidth = 4;
					vertices.Add(new PdfNumber(item.MinX - rHalfWidth));
					vertices.Add(new PdfNumber(item.MinY - rHalfWidth));
					vertices.Add(new PdfNumber(item.MaxX + rHalfWidth));
					vertices.Add(new PdfNumber(item.MinY - rHalfWidth));
					vertices.Add(new PdfNumber(item.MaxX + rHalfWidth));
					vertices.Add(new PdfNumber(item.MaxY + rHalfWidth));
					vertices.Add(new PdfNumber(item.MinX - rHalfWidth));
					vertices.Add(new PdfNumber(item.MaxY + rHalfWidth));
					vertices.Add(new PdfNumber(item.MinX - rHalfWidth));
					vertices.Add(new PdfNumber(item.MinY - rHalfWidth));
					PdfAnnotation highlight = PdfAnnotation.CreatePolygonPolyline(stamper.Writer, rect, stComment, true, vertices);
					highlight.Color = color;
					highlight.Title = stTittle;
					highlight.Border = new PdfBorderArray(0, 0, rHalfWidth * 2);
					stamper.AddAnnotation(highlight, item.Page + iPageAdd);
					iPosition++;
					iCount--;
				}
				if(iCount == 0)
					return;
				int i;
				for(i = iPosition; i < iPosition + iCount; i++)
				{
					if(items[i].IsImage)
					{
						MarkText(items, iPosition, i - iPosition, stamper, color, iPageAdd, stComment);
						MarkText(items, i, 1, stamper, color, iPageAdd, stComment);
						if(i + 1 >= iPosition + iCount)
							return;
						iCount -= (i - iPosition) + 1;
						iPosition = i + 1;
					}
				}
				if(!items[iPosition].IsNewLine)
				{
					i = iPosition;
					while(i < iPosition + iCount - 1 && !items[i + 1].IsNewLine)
						i++;
					Rectangle rect = new Rectangle(items[iPosition].MinX, items[iPosition].MaxY, items[i].MaxX, items[iPosition].MinY);
					float[] quad = { rect.Left, rect.Bottom, rect.Right, rect.Bottom, rect.Left, rect.Top, rect.Right, rect.Top };
					PdfAnnotation highlight = PdfAnnotation.CreateMarkup(stamper.Writer, rect, stComment, PdfAnnotation.MARKUP_HIGHLIGHT, quad);
					highlight.Color = color;
					highlight.Title = stTittle;
					stamper.AddAnnotation(highlight, items[iPosition].Page + iPageAdd);
					iCount -= (i - iPosition) + 1;
					iPosition = i + 1;
				}
				if(iCount == 0)
					return;
				if(items.Count > iPosition + iCount && !items[iPosition + iCount].IsNewLine)
				{
					i = iPosition + iCount - 1;
					while(!items[i].IsNewLine)
						i--;
					Rectangle rect = new Rectangle(items[i].MinX, items[i].MaxY, items[i = iPosition + iCount - 1].MaxX, items[i].MinY);
					float[] quad = { rect.Left, rect.Bottom, rect.Right, rect.Bottom, rect.Left, rect.Top, rect.Right, rect.Top };
					PdfAnnotation highlight = PdfAnnotation.CreateMarkup(stamper.Writer, rect, stComment, PdfAnnotation.MARKUP_HIGHLIGHT, quad);
					highlight.Color = color;
					highlight.Title = stTittle;
					stamper.AddAnnotation(highlight, items[iPosition].Page + iPageAdd);
					iCount -= iPosition + iCount - i;
				}
				if(iCount == 0)
					return;
				i = iPosition;
				float rMinX = items[iPosition].MinX;
				float rMaxX = items[iPosition].MaxX;
				while(i < iPosition + iCount)
				{
					i++;
					if(i == iPosition + iCount || items[iPosition].Page != items[i].Page)
					{
						Rectangle rect = new Rectangle(rMinX, items[iPosition].MaxY, rMaxX, items[i - 1].MinY);
						float[] quad = { rect.Left, rect.Bottom, rect.Right, rect.Bottom, rect.Left, rect.Top, rect.Right, rect.Top };
						PdfAnnotation highlight = PdfAnnotation.CreateMarkup(stamper.Writer, rect, stComment, PdfAnnotation.MARKUP_HIGHLIGHT, quad);
						highlight.Color = color;
						highlight.Title = stTittle;
						stamper.AddAnnotation(highlight, items[iPosition].Page + iPageAdd);
						iCount -= i - iPosition;
						iPosition = i;
						if(iCount == 0)
							return;
						rMinX = items[iPosition].MinX;
						rMaxX = items[iPosition].MaxX;
					}
					if(items[i].MinX < rMinX)
						rMinX = items[i].MinX;
					if(items[i].MaxX > rMaxX)
						rMaxX = items[i].MaxX;
				}
				/*
				for(i = iPosition; i < iPosition + iCount; i++)
				{
					TextItem item = items[i];
					Rectangle rect = new Rectangle(item.MinX, item.MaxY, item.MaxX, item.MinY);
					float[] quad = { rect.Left, rect.Bottom, rect.Right, rect.Bottom, rect.Left, rect.Top, rect.Right, rect.Top };
					PdfAnnotation highlight = PdfAnnotation.CreateMarkup(stamper.Writer, rect, stComment, PdfAnnotation.MARKUP_HIGHLIGHT, quad);
					highlight.Color = color;
					highlight.Title = stTittle;
					stamper.AddAnnotation(highlight, item.Page + iPageAdd);
				}
				*/
			}
		}
		//--------------------------------------------------------------------------------------------------
		void InsertPage(PdfStamper stamper, int iPosition, Rectangle rect)
		{
			stamper.InsertPage(iPosition, rect);
			PdfContentByte canvas = stamper.GetOverContent(iPosition);
			canvas.BeginText();
			canvas.SetFontAndSize(GetDefaultFont(), 12);
			canvas.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "<Страница отсутствует>", (rect.Left + rect.Right) / 2, (rect.Bottom + rect.Top) / 2, 0);
			canvas.EndText();
		}
		//--------------------------------------------------------------------------------------------------
		void AddHeader(PdfStamper stamper, PdfReader reader, string stComment)
		{
			PdfContentByte canvas = stamper.GetOverContent(1);
			canvas.BeginText();
			canvas.SetFontAndSize(GetDefaultFont(), 12);
			Rectangle rect = reader.GetPageSizeWithRotation(1);
			canvas.ShowTextAligned(PdfContentByte.ALIGN_CENTER, stComment, (rect.Left + rect.Right) / 2, rect.Top - 12, 0);
			canvas.EndText();
			rect.Bottom = rect.Top - 18;
			float[] quad = { rect.Left, rect.Bottom, rect.Right, rect.Bottom, rect.Left, rect.Top, rect.Right, rect.Top };
			PdfAnnotation highlight = PdfAnnotation.CreateMarkup(stamper.Writer, rect, stComment, PdfAnnotation.MARKUP_HIGHLIGHT, quad);
			highlight.Color = BaseColor.YELLOW;
			highlight.Title = "PDFCompare";
			stamper.AddAnnotation(highlight, 1);
		}
		//--------------------------------------------------------------------------------------------------
		BaseFont GetDefaultFont()
		{
			return BaseFont.CreateFont(System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "ARIAL.TTF"), BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
		}
		//--------------------------------------------------------------------------------------------------
		BaseColor GetBaseColor(int iColor)
		{
			return new BaseColor(iColor & 0x000000FF, (iColor & 0x0000FF00) >> 8, (iColor & 0x00FF0000) >> 16);
		}
		//--------------------------------------------------------------------------------------------------
		// При удалении фрагмента в самом конце индекс будет указывать за границы массива
		// В этом случае нам нужен последний элемент
		TextItem IntexOrLast(List<TextItem> items, int iIndex)
		{
			return iIndex < items.Count ? items[iIndex] : items.Last();
		}
		////////////////////////////////////////////////////////////////////////////////////////////////////
		//--------------------------------------------------------------------------------------------------
		public class TextItem
		{
			public TextItem()
			{
				MinX = 1;// Invalidate bound
			}
			public TextItem(string stText, int iPage, bool isNewLine)
			{
				MinX = 1;// Invalidate bound
				Text = stText;
				Page = iPage;
				IsNewLine = isNewLine;
			}
			public string Text { get; set; }
			public int Page { get; set; }
			public float MinX { get; set; }
			public float MinY { get; set; }
			public float MaxX { get; set; }
            public float MaxY { get; set; }

            public Vector OrientationVector { get; set; }
            public Vector StartPoint { get; set; }
            public Vector EndPoint { get; set; }

            public LineSegment AscentLine { get; set; }
            public LineSegment DescentLine { get; set; }

			public bool IsImage { get; set; }
			public bool IsNewLine { get; set; }
			public void BoundAppend(LineSegment line)
			{
				BoundAppend(line.GetStartPoint());
				BoundAppend(line.GetEndPoint());

                EndPoint = line.GetEndPoint();
			}
			public void BoundAppend(Vector vector)
			{
				if(MinX > MaxX)
				{
					MinX = MaxX = vector[0];
					MinY = MaxY = vector[1];
				}
				else
				{
					if(vector[0] < MinX)
						MinX = vector[0];
					if(vector[0] > MaxX)
						MaxX = vector[0];
					if(vector[1] < MinY)
						MinY = vector[1];
					if(vector[1] > MaxY)
						MaxY = vector[1];
				}
			}
		}
		////////////////////////////////////////////////////////////////////////////////////////////////////
		//--------------------------------------------------------------------------------------------------
		public class LocationExtractionStrategy : ITextExtractionStrategy
		{
			private List<TextChunk> m_LocationalResult;
			private List<TextItem> m_TextItems;
			private Dictionary<string, ImageData> m_Images = new Dictionary<string, ImageData>();
			private int m_iPage;
			public int Page
			{
				get { return m_iPage; }
				set
				{
					GetTextItems();
					m_iPage = value;
				}
			}

			//--------------------------------------------------------------------------------------------------
			public LocationExtractionStrategy()
			{
				Reset();
			}
			//--------------------------------------------------------------------------------------------------
			public void Reset()
			{
				m_iPage = 0;
				m_TextItems = new List<TextItem>();
				m_LocationalResult = new List<TextChunk>();
			}
			//--------------------------------------------------------------------------------------------------
			public List<TextItem> GetTextItems()
			{
				if(m_LocationalResult.Count != 0)
				{
					m_LocationalResult.Sort();
					bool isNewLine = true;
					TextItem curItem = null;
					TextChunk lastChunk = null;
					foreach(TextChunk chunk in m_LocationalResult)
					{
						bool bStartNewItem = false;
						if(lastChunk != null)
						{
							if(chunk.SameLine(lastChunk))
							{
								float dist = chunk.DistanceFromEndOf(lastChunk);
								if(dist < -chunk.charSpaceWidth)
									bStartNewItem = true;
								// we only insert a blank space if the trailing character of the previous string wasn't a space, and the leading character of the current string isn't a space
								else if(dist > chunk.charSpaceWidth / 2.0f && chunk.text[0] != ' ' && lastChunk.text[lastChunk.text.Length - 1] != ' ')
									bStartNewItem = true;
							}
							else
							{
								bStartNewItem = true;
								isNewLine = true;
							}
						}
						lastChunk = chunk;
						if(bStartNewItem && curItem != null)
						{
							m_TextItems.Add(curItem);
							curItem = null;
						}
						if(chunk.isImage)
						{
							if(curItem != null)
								m_TextItems.Add(curItem);
							curItem = new TextItem(chunk.text, chunk.iPage, isNewLine);
							isNewLine = false;

                            curItem.OrientationVector = chunk.orientationVector;

                            curItem.StartPoint = chunk.startLocation;
                            curItem.EndPoint = chunk.endLocation;
                            curItem.MinX = chunk.tfmImage[Matrix.I31];
							curItem.MinY = chunk.tfmImage[Matrix.I32];
							curItem.MaxX = chunk.tfmImage[Matrix.I11] + curItem.MinX;
							curItem.MaxY = chunk.tfmImage[Matrix.I22] + curItem.MinY;
							curItem.IsImage = true;
							m_TextItems.Add(curItem);
							curItem = null;
						}
						else
						{
							string st = chunk.text;
                            int stLength = new StringInfo(st).LengthInTextElements;
							int iWordStart = 0;
                            while(iWordStart < stLength)
							{
								if(st[iWordStart] == ' ')
								{
									if(curItem != null)
									{
										m_TextItems.Add(curItem);
										curItem = null;
									}
									iWordStart++;
									continue;
								}
								int iWordEnd = iWordStart;
                                while(iWordEnd < stLength && st[iWordEnd] != ' ')
									iWordEnd++;
                                if (curItem == null)
                                {
                                    curItem = new TextItem(st.Substring(iWordStart, iWordEnd - iWordStart), chunk.iPage, isNewLine);
                                    curItem.OrientationVector = chunk.orientationVector;
                                    curItem.StartPoint = chunk.startLocation;
                                    curItem.EndPoint = chunk.endLocation;

                                    curItem.AscentLine = new LineSegment(chunk.AscentLines[iWordStart].GetStartPoint(), chunk.AscentLines[iWordEnd - 1].GetEndPoint());
                                    curItem.DescentLine = new LineSegment(chunk.DescentLines[iWordStart].GetStartPoint(), chunk.DescentLines[iWordEnd - 1].GetEndPoint());

                                    isNewLine = false;
                                }
                                else
                                {
                                    curItem.Text += st.Substring(iWordStart, iWordEnd - iWordStart);

                                    // Передвинем только задний конец, передний оставим как есть.
                                    curItem.AscentLine = new LineSegment(curItem.AscentLine.GetStartPoint(), chunk.AscentLines[iWordEnd - 1].GetEndPoint());
                                    curItem.DescentLine = new LineSegment(curItem.DescentLine.GetStartPoint(), chunk.DescentLines[iWordEnd - 1].GetEndPoint());
                                }
								for(int i = iWordStart; i < iWordEnd; i++)
								{
									curItem.BoundAppend(chunk.AscentLines[i]);
									curItem.BoundAppend(chunk.DescentLines[i]);
								}
								iWordStart = iWordEnd;
							}
						}
					}
					if(curItem != null)
						m_TextItems.Add(curItem);
					m_LocationalResult = new List<TextChunk>();
				}
				return m_TextItems;
			}
			//--------------------------------------------------------------------------------------------------
			public string GetResultantText() { return ""; }
			public void BeginTextBlock() { }
			public void EndTextBlock() { }
			//--------------------------------------------------------------------------------------------------
			public void RenderImage(ImageRenderInfo renderInfo)
			{
				string stImageName = null;
				ImageData image = new ImageData();
				image.data = renderInfo.GetImage().GetImageAsBytes();
				foreach(KeyValuePair<string, ImageData> entry in m_Images)
				{
					if(entry.Value.data.SequenceEqual(image.data))
					{
						stImageName = entry.Key;
						break;
					}
				}
				if(string.IsNullOrEmpty(stImageName))
				{
					stImageName = string.Format("PDFCompareTextItemImage{0}", m_Images.Count + 1);
					m_Images.Add(stImageName, image);
				}
				Matrix tfm = renderInfo.GetImageCTM();
				float rMinX = tfm[Matrix.I31];
				float rMinY = tfm[Matrix.I32];
				float rMaxX = tfm[Matrix.I11] + rMinX;
				float rMaxY = tfm[Matrix.I22] + rMinY;
				TextChunk location = new TextChunk(stImageName, new Vector(rMinX, rMaxY, 0), new Vector(rMaxX, rMaxY, 0), 0);
				location.iPage = Page;
				location.isImage = true;
				location.tfmImage = renderInfo.GetImageCTM();
				m_LocationalResult.Add(location);
			}
			//--------------------------------------------------------------------------------------------------
			public void RenderText(TextRenderInfo renderInfo)
			{
				LineSegment segment = renderInfo.GetBaseline();
				TextChunk location = new TextChunk(renderInfo.GetText(), segment.GetStartPoint(), segment.GetEndPoint(), renderInfo.GetSingleSpaceWidth());
				location.iPage = Page;

                int renderInfoTextLength = new StringInfo(renderInfo.GetText()).LengthInTextElements;
                if (renderInfoTextLength == 1)
				{
                    location.AscentLines.Add(renderInfo.GetAscentLine());
					location.DescentLines.Add(renderInfo.GetDescentLine());
				}
				else
				{
					IList<TextRenderInfo> infos = renderInfo.GetCharacterRenderInfos();
					System.Diagnostics.Debug.Assert(infos != null);
                    System.Diagnostics.Debug.Assert(renderInfoTextLength == infos.Count);
					foreach(TextRenderInfo info in infos)
					{
						location.AscentLines.Add(info.GetAscentLine());
						location.DescentLines.Add(info.GetDescentLine());
					}
				}
				m_LocationalResult.Add(location);
			}
			//--------------------------------------------------------------------------------------------------
			private class TextChunk : IComparable<TextChunk>
			{
				/** the text of the chunk */
				internal string text;
				/** the starting location of the chunk */
				internal Vector startLocation;
				/** the ending location of the chunk */
				internal Vector endLocation;
				/** unit vector in the orientation of the chunk */
				internal Vector orientationVector;
				/** the orientation as a scalar for quick sorting */
				internal int orientationMagnitude;
				/** perpendicular distance to the orientation unit vector (i.e. the Y position in an unrotated coordinate system)
				 * we round to the nearest integer to handle the fuzziness of comparing floats */
				internal float distPerpendicular;
				//internal int distPerpendicular;
				/** distance of the start of the chunk parallel to the orientation unit vector (i.e. the X position in an unrotated coordinate system) */
				internal float distParallelStart;
				/** distance of the end of the chunk parallel to the orientation unit vector (i.e. the X position in an unrotated coordinate system) */
				internal float distParallelEnd;
				/** the width of a single space character in the font of the chunk */
				internal float charSpaceWidth;
				internal List<LineSegment> AscentLines = new List<LineSegment>();
				internal List<LineSegment> DescentLines = new List<LineSegment>();
				internal int iPage = 0;
				internal bool isImage = false;
				internal Matrix tfmImage;

				public TextChunk(String str, Vector startLocation, Vector endLocation, float charSpaceWidth)
				{
					this.text = str;
					this.startLocation = startLocation;
					this.endLocation = endLocation;
					this.charSpaceWidth = charSpaceWidth;

					orientationVector = endLocation.Subtract(startLocation).Normalize();
					orientationMagnitude = (int)(Math.Atan2(orientationVector[Vector.I2], orientationVector[Vector.I1]) * 1000);

					// see http://mathworld.wolfram.com/Point-LineDistance2-Dimensional.html
					// the two vectors we are crossing are in the same plane, so the result will be purely
					// in the z-axis (out of plane) direction, so we just take the I3 component of the result
					Vector origin = new Vector(0, 0, 1);
					distPerpendicular = (startLocation.Subtract(origin)).Cross(orientationVector)[Vector.I3];
					//distPerpendicular = (int)(startLocation.Subtract(origin)).Cross(orientationVector)[Vector.I3];

					distParallelStart = orientationVector.Dot(startLocation);
					distParallelEnd = orientationVector.Dot(endLocation);
				}
				/**
				 * @param as the location to compare to
				 * @return true is this location is on the the same line as the other
				 */
				public bool SameLine(TextChunk a)
				{
					if(orientationMagnitude != a.orientationMagnitude) return false;
					if(distPerpendicular != a.distPerpendicular) return false;
					return true;
				}
				/**
				 * Computes the distance between the end of 'other' and the beginning of this chunk
				 * in the direction of this chunk's orientation vector.  Note that it's a bad idea
				 * to call this for chunks that aren't on the same line and orientation, but we don't
				 * explicitly check for that condition for performance reasons.
				 * @param other
				 * @return the number of spaces between the end of 'other' and the beginning of this chunk
				 */
				public float DistanceFromEndOf(TextChunk other)
				{
					float distance = distParallelStart - other.distParallelEnd;
					return distance;
				}
				/**
				 * Compares based on orientation, perpendicular distance, then parallel distance
				 * @see java.lang.Comparable#compareTo(java.lang.Object)
				 */
				public int CompareTo(TextChunk rhs)
				{
					if(this == rhs) return 0; // not really needed, but just in case

					int rslt;
					rslt = CompareInts(orientationMagnitude, rhs.orientationMagnitude);
					if(rslt != 0) return rslt;

					rslt = CompareFloats(distPerpendicular, rhs.distPerpendicular);
					//rslt = CompareInts(distPerpendicular, rhs.distPerpendicular);
					if(rslt != 0) return rslt;

					// note: it's never safe to check floating point numbers for equality, and if two chunks
					// are truly right on top of each other, which one comes first or second just doesn't matter
					// so we arbitrarily choose this way.
					rslt = distParallelStart < rhs.distParallelStart ? -1 : 1;

					return rslt;
				}
				private static int CompareInts(int int1, int int2)
				{
					return int1 == int2 ? 0 : int1 < int2 ? -1 : 1;
				}
				private static int CompareFloats(float r1, float r2)
				{
					return (Math.Abs(r2 - r1) <= 1) ? 0 : r1 < r2 ? -1 : 1;
				}
			}
			//--------------------------------------------------------------------------------------------------
			private class ImageData
			{
				public byte[] data;
			}
		}
	}
}
