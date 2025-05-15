
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using WordDemo5.Models;
using ClosedXML.Excel;
using System.Text.Json;
using System.Text;


namespace WordDemo5.Controllers
{
    public class DocumentController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult SubmitForm(DocumentModel model) //Formdan gelen dinamik verileri almak için.
        {
            //Model üzerindeki verilerin kontrolünün sağlanması.
            if (model.ContentItems == null || !model.ContentItems.Any())
            {
                TempData["ErrorMessage"] = "En az bir içerik eklenmelidir.";
                return RedirectToAction("Index", "Home");
            }

            foreach (var item in model.ContentItems)
            {
                foreach (var block in item.ContentBlocks)
                {
                    if (block.ContentType == "Paragraph" && string.IsNullOrWhiteSpace(block.ParagraphText))
                    {
                        TempData["ErrorMessage"] = "Paragraf metni boş olamaz.";
                        return RedirectToAction("Index", "Home");
                    }
                    else if (block.ContentType == "Image" && block.ImageFile == null)
                    {
                        TempData["ErrorMessage"] = "Resim dosyası seçilmelidir.";
                        return RedirectToAction("Index", "Home");
                    }
                }
            }

            string fileName = "Document.docx";
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/downloads", fileName);

            try
            {
                // Word belgesi oluşturma
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = new Body();

                    // İşlemleri gerçekleştirecek olan metod
                    ProcessCommonSections(mainPart, filePath, body, model); //MainPart silinecek.

                    // Belgeyi kaydetme
                    mainPart.Document.Append(body);
                    mainPart.Document.Save();
                }

                TempData["Message"] = "Belge başarıyla oluşturuldu!";
                TempData["FilePath"] = $"/downloads/{fileName}";
            }
            catch (Exception ex)
            {
                TempData["ErrorMessage"] = $"Belge oluşturulurken bir hata oluştu: {ex.Message}";
            }

            // İşlem sonrası Home/Index görünümüne yönlendir
            return RedirectToAction("Index", "Home");
        }

        private void ProcessCommonSections(MainDocumentPart mainPart, string filePath, Body body, DocumentModel model) //Formdan gelen dinamik verileri yakalamak için.
        {
            // Kapak Sayfası - Section 1
            AddCoverPage(body, model); //Model den gelen verileri alacağı için model'ı ekledim.

            // Kapak sayfasını ayrı bir bölüm olarak tanımla
            SectionProperties coverSectionProperties = new SectionProperties(
                new PageSize() { Width = 11906, Height = 16838 }, // A4 boyutunda sayfa
                new PageMargin() { Top = 1138, Bottom = 1138, Left = 1138, Right = 1138 } // Kenar boşlukları
            );
            body.Append(coverSectionProperties);

            // Kapak Sayfası ve İkinci Bölüm Arasında Bir Section Break Ekleyin
            Paragraph sectionBreakParagraph = new Paragraph(
                new ParagraphProperties(new SectionProperties())); // Bölüm ayracı ekle
            body.Append(sectionBreakParagraph);

            // Yeni Section: Dikeyde İkiye Bölünmüş Sayfa
            SectionProperties splitSectionProperties = new SectionProperties(
                new PageSize() { Width = 11906, Height = 16838 }, // A4 boyutunda sayfa
                new PageMargin() { Top = 1138, Bottom = 1138, Left = 1138, Right = 1138 }, // Kenar boşlukları
                new Columns() { ColumnCount = 2, EqualWidth = true, Space = "453" } // 2 sütun, eşit genişlikte
            );
            // Sütun genişliklerini ayarlama
            Column column1 = new Column() { Width = "4596" }; // 8.1 cm genişlik (1 cm = 567 twips, 8.1 cm = 4596.7 twips)
            Column column2 = new Column() { Width = "4596" }; // 8.1 cm genişlik
            Columns columns = new Columns();
            columns.Append(column1);
            columns.Append(column2);
            splitSectionProperties.Append(columns);

            // Yeni Section - section 2 --> Giriş kısmından sonraası buraya eklencek

            AddContentSection2(mainPart, filePath, body, model); //Başlıklar, paragraflar, resimler ..... ekleneyen method.

            body.Append(splitSectionProperties);
        }

        //GİRİŞ SECTİON ALANININ OLUŞTURULMASINI SAĞLAYAN METHOD
        private void AddContentSection2(MainDocumentPart mainPart, string filePath, Body body, DocumentModel model)
        {
            int titleIndex = 1;
            int imageCounter = 0;
            int tableCounter = 0;
            int subTitleIndex = 0;//Yeni başlık eklendikten sonra bu sıfırlanacak.Çünkü Her başlığın alt başlıgı 1 den başlamalı.
            int subSubTitleIndex = 0; //Alt başlığın alt başlıgını temsil eden index.

            foreach (var item in model.ContentItems)
            {
                CreateParagraphTitle(titleIndex, body, item.Title);

                //Paragraf ve resimlerin sıralı bir şekilde eklenmesi işlemi
                foreach (var block in item.ContentBlocks)
                {
                    if (block.ContentType == "Paragraph")
                    {
                        CreateParagraph(body, block.ParagraphText);
                    }
                    else if (block.ContentType == "Image" && block.ImageFile != null)
                    {
                        imageCounter++;
                        CreateImage(mainPart, body, block.ImageFile);
                    }
                    else if (block.ContentType == "ImageCaption")
                    {
                        // Mevcut görsel numarasını (imageCounter) ile açıklama ekle
                        CreateImageExplanation(body, block.ImageCaptionText, imageCounter);
                    }
                    else if (block.ContentType == "TableCaption")
                    {
                        //Ekleme işlemini tabloyla birlikte yaptım.
                        tableCounter++;
                        CreateTableExplanation(body, block.TableCaptionText, tableCounter);
                    }
                    else if (block.ContentType == "Table" && block.TableFile != null)
                    {
                        //Tablo sayısı bir arttırılacak ve tablo ekleme işlemi yapılacak.
                        var tableFile = ReadExcelToTable(block.TableFile); //Excel tablosunu oku
                        AddTableToWordBody(body, tableFile); //verileri tabloya ekle.
                    }
                    else if (block.ContentType == "SubTitle")
                    {
                        //Alt başlık ekle methodu oluştur.
                        subTitleIndex++;
                        string formattedText = "";
                        if (block.SubTitleText != null)
                        {
                            formattedText = ToTitleCase(block.SubTitleText);
                        }
                        //Her kelimenin baş harfini büyük hale getir.
                        AddSubTitle(body, formattedText, titleIndex, subTitleIndex);
                    }
                    else if (block.ContentType == "SubSubTitle")
                    {
                        //Alt başlığa alt başlık ekle methodu oluştur.
                        subSubTitleIndex++;
                        string formattedText = "";
                        if (block.SubSubTitleText != null)
                        {
                            formattedText = ToTitleCase(block.SubSubTitleText);
                        }
                        AddSubSubTitle(body, titleIndex, subTitleIndex, subSubTitleIndex, formattedText);
                    }


                }
                titleIndex++; //Bir sonraki başlığa geçilmesini sağlamak için.
                //Yeni eklenen başlığın altında bunların yeniden 1 olarak başlaması gerekiyor bu nedenle default değerlere döndürdüm.
                subTitleIndex = 0;
                subSubTitleIndex = 0;

            }
        }
        /// <summary>
        /// Alt başlık için alt başlık ekleme işlemi.
        /// </summary>
        /// <param name="body"></param>
        /// <param name="subSubTitleText"></param>
        /// <param name="titleIndex"></param>
        /// <param name="subTitleIndex"></param>
        /// <param name="subSubTitleIndex"></param>
        /// <param name="text"></param>
        private void AddSubSubTitle(Body body, int titleIndex, int subTitleIndex, int subSubTitleIndex, string text) //1.1.1. Text, 1.1.2. Text2 şeklinde olacak.
        {
            // Alt Başlık için alt başlık ekleme
            RunProperties runPropertiesForTitle = new RunProperties(
                new FontSize() { Val = "18" }, // 9 Punto
                new Bold(), // Kalın
                new RunFonts() { Ascii = "Cambria Bold" } // Yazı tipi
            );


            ParagraphProperties paragrapPropertiesForTitle = new ParagraphProperties(
                new SpacingBetweenLines() { Before = "0", After = "0" }, // Yazdıktan sonra boşluk
                new Justification() { Val = JustificationValues.Left } // Sola hizalama
            );

            Run runForTitle = new Run(new Text($"{titleIndex}.{subTitleIndex}.{subSubTitleIndex} {text}")) //Başlığın 1.1.1. Text şeklinde eklenmesi için
            {
                RunProperties = runPropertiesForTitle
            };
            Paragraph paragraphForSubSubTitle = new Paragraph();
            paragraphForSubSubTitle.Append(paragrapPropertiesForTitle.CloneNode(true)); //CloneNode her seferinde aynı özelliği her yeni gelen veriye eklemesi için.
            paragraphForSubSubTitle.Append(runForTitle);
            body.Append(paragraphForSubSubTitle);
        }
        /// <summary>
        /// Alt başlık ekleme işlemi
        /// </summary>
        /// <param name="body"></param>
        /// <param name="text"></param>
        /// <param name="titleIndex"></param>
        /// <param name="subTitleIndex"></param>
        private void AddSubTitle(Body body, string text, int titleIndex, int subTitleIndex)
        {
            // Alt Başlık ekleme
            RunProperties runPropertiesForTitle = new RunProperties(
                new FontSize() { Val = "18" }, // 9 Punto
                new Bold(), // Kalın
                new RunFonts() { Ascii = "Cambria Bold" } // Yazı tipi
            );

            ParagraphProperties paragrapPropertiesForTitle = new ParagraphProperties(
                new SpacingBetweenLines() { Before = "0", After = "0" }, // Yazdıktan sonra boşluk
                new Justification() { Val = JustificationValues.Left } // Sola hizalama
            );

            Run runForTitle = new Run(new Text($"{titleIndex}.{subTitleIndex}. {text}")) //Başlığın 1.1.Text şeklinde eklenmesi için
            {
                RunProperties = runPropertiesForTitle
            };
            Paragraph paragraphForSubTitle = new Paragraph();
            paragraphForSubTitle.Append(paragrapPropertiesForTitle.CloneNode(true)); //CloneNode her seferinde aynı özelliği her yeni gelen veriye eklemesi için.
            paragraphForSubTitle.Append(runForTitle);
            body.Append(paragraphForSubTitle);

        }
        /// <summary>
        /// Girilen tablonun açıklama satırının eklenmesi.
        /// </summary>
        /// <param name="body"></param>
        /// <param name="explanationText"></param>
        /// <param name="currentTableNumber"></param>
        private void CreateTableExplanation(Body body, string explanationText, int currentTableNumber)
        {
            if (string.IsNullOrEmpty(explanationText))
                explanationText = string.Empty;

            // Başlık için RunProperties
            RunProperties runPropertiesForTableExpTitle = new RunProperties(
                new FontSize() { Val = "18" },
                new Bold(),
                new RunFonts() { Ascii = "Cambria" }
            );

            // Açıklama metni için RunProperties
            RunProperties runPropertiesForTableExp = new RunProperties(
                new FontSize() { Val = "18" },
                new RunFonts() { Ascii = "Cambria" }
            );

            // Paragraf özellikleri
            ParagraphProperties contentParagraphProperties = new ParagraphProperties(
                new Justification() { Val = JustificationValues.Left },
                new SpacingBetweenLines() { Before = "300", After = "0" }
            );

            // "Tablo X." başlığı (Bold)
            Run TableExpTitle = new Run(new Text($"Tablo {currentTableNumber}. ")
            {
                Space = SpaceProcessingModeValues.Preserve // Boşluğu koru
            })
            {
                RunProperties = runPropertiesForTableExpTitle
            };

            // Kullanıcı metni (Normal)
            Run TableExpRun = new Run(new Text(" " + explanationText))
            {
                RunProperties = runPropertiesForTableExp
            };

            // Paragrafı oluştur ve öğeleri ekle
            Paragraph contentParagraph = new Paragraph();
            contentParagraph.Append(contentParagraphProperties.CloneNode(true));
            contentParagraph.Append(TableExpTitle);
            contentParagraph.Append(TableExpRun);

            body.Append(contentParagraph);

        }
        /// <summary>
        /// Kullanıcıdan gelen excel tablosunun okunması.
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        private List<List<string>> ReadExcelToTable(IFormFile file)
        {
            var tableData = new List<List<string>>();

            using (var stream = new MemoryStream())
            {
                file.CopyTo(stream);
                using (var workbook = new XLWorkbook(stream))
                {
                    var worksheet = workbook.Worksheets.First();
                    var rowCount = worksheet.LastRowUsed().RowNumber();
                    var colCount = worksheet.LastColumnUsed().ColumnNumber();

                    for (int row = 1; row <= rowCount; row++)
                    {
                        var rowData = new List<string>();
                        for (int col = 1; col <= colCount; col++)
                        {
                            var cellValue = worksheet.Cell(row, col).GetValue<string>();
                            rowData.Add(cellValue);
                        }
                        tableData.Add(rowData);
                    }
                }
            }

            return tableData;
        }
        /// <summary>
        /// Okunan tablo verilerinin word belgesine tablo olarak eklenmesi.
        /// </summary>
        /// <param name="body"></param>
        /// <param name="tableData"></param>
        public void AddTableToWordBody(Body body, List<List<string>> tableData)
        {
            Table table = new Table();

            // Tablo genişlik ve kenarlık ayarları
            TableProperties tblProp = new TableProperties(
                new TableWidth()
                {
                    Width = "5000", // Genişlik değeri (tam genişlik için 100% kullanmak adına "5000")
                    Type = TableWidthUnitValues.Pct // Yüzdelik olarak ayarlanıyor
                },
                new TableBorders(
                    new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                    new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                    new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                    new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                    new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                    new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 }
                ),
                new TableJustification() { Val = TableRowAlignmentValues.Center },
                new TableLook()
                {
                    FirstRow = true,
                    NoHorizontalBand = false,
                    NoVerticalBand = true
                }
            );
            table.AppendChild(tblProp);
            foreach (var rowData in tableData)
            {
                TableRow tr = new TableRow();
                foreach (var cellData in rowData)
                {
                    // Yazı tipi ve boyutu ayarları
                    RunProperties runProperties = new RunProperties(
                        new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri" }, // Yazı tipi
                        new FontSize() { Val = "20" }, // 10 punto = 20 half-point
                        new FontSizeComplexScript() { Val = "20" } // Karmaşık yazı sistemleri için de aynı boyut
                    );

                    Run run = new Run();
                    run.Append(runProperties);
                    run.Append(new Text(cellData ?? ""));

                    ParagraphProperties paragraphCellProperties = new ParagraphProperties(
                        new Justification() { Val = JustificationValues.Center }
                    );

                    Paragraph paragraph = new Paragraph();
                    paragraph.Append(paragraphCellProperties);
                    paragraph.Append(run);

                    TableCell tc = new TableCell(paragraph);

                    // Hücreyi dikeyde ortalama
                    TableCellProperties cellProps = new TableCellProperties(
                        new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
                    );
                    tc.Append(cellProps);

                    tr.Append(tc);
                }
                table.Append(tr);
            }

            body.Append(table);
            // Tablo sonrasına boşluk bırakmak için paragraf ekle
            Paragraph spacerParagraph = new Paragraph();
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            SpacingBetweenLines spacing = new SpacingBetweenLines() { After = "50" };
            paragraphProperties.Append(spacing);
            spacerParagraph.Append(paragraphProperties);

            body.Append(spacerParagraph);
        }
        /// <summary>
        /// Paragaraf Başlıklarının eklenmesi işlemi.
        /// </summary>
        /// <param name="titleIndex"></param>
        /// <param name="body"></param>
        /// <param name="text"></param>
        private void CreateParagraphTitle(int titleIndex, Body body, string text)
        {
            // Başlık ekleme
            RunProperties runPropertiesForTitle = new RunProperties(
                new FontSize() { Val = "20" }, // 10 Punto
                new Bold(), // Kalın
                new RunFonts() { Ascii = "Cambria Bold" } // Yazı tipi
            );

            ParagraphProperties paragrapPropertiesForTitle = new ParagraphProperties(
                new SpacingBetweenLines() { Before = "0", After = "0" }, // Yazdıktan sonra boşluk
                new Justification() { Val = JustificationValues.Left } // Sola hizalama
            );

            Run runForTitle = new Run(new Text($"{titleIndex}. {text}")) //Başlığın 1. 2. şeklinde eklenmesi için
            {
                RunProperties = runPropertiesForTitle
            };
            Paragraph paragraphForTitle = new Paragraph();
            paragraphForTitle.Append(paragrapPropertiesForTitle.CloneNode(true)); //CloneNode her seferinde aynı özelliği her yeni gelen veriye eklemesi için.
            paragraphForTitle.Append(runForTitle);
            body.Append(paragraphForTitle);

        }
        /// <summary>
        /// Resim açıklaması eklenmesi işlemi için.
        /// </summary>
        /// <param name="body">Üzerine eklenecek belge </param>
        /// <param name="explanationText">Görsel açıklası.</param>
        private void CreateImageExplanation(Body body, string explanationText, int currentImageNumber)
        {
            if (string.IsNullOrEmpty(explanationText))
                explanationText = string.Empty;

            // Başlık için RunProperties
            RunProperties runPropertiesForImageExpTitle = new RunProperties(
                new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "18" },
                new Bold(),
                new RunFonts() { Ascii = "Cambria" }
            );

            // Açıklama metni için RunProperties
            RunProperties runPropertiesForImageExp = new RunProperties(
                new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "18" },
                new RunFonts() { Ascii = "Cambria" }
            );

            // Paragraf özellikleri
            ParagraphProperties contentParagraphProperties = new ParagraphProperties(
                new Justification() { Val = JustificationValues.Left },
                new SpacingBetweenLines() { Before = "0", After = "300" }
            );

            // "Şekil X." başlığı (Bold)
            Run ImageExpTitle = new Run(new Text($"Şekil {currentImageNumber}. ")
            {
                Space = SpaceProcessingModeValues.Preserve // Boşluğu koru
            })
            {
                RunProperties = runPropertiesForImageExpTitle
            };

            // Kullanıcı metni (Normal)
            Run ImageExpRun = new Run(new Text(" " + explanationText))
            {
                RunProperties = runPropertiesForImageExp
            };

            // Paragrafı oluştur ve öğeleri ekle
            Paragraph contentParagraph = new Paragraph();
            contentParagraph.Append(contentParagraphProperties.CloneNode(true));
            contentParagraph.Append(ImageExpTitle);
            contentParagraph.Append(ImageExpRun);

            body.Append(contentParagraph);

        }
        /// <summary>
        /// Görsel oluşturulması işlemini yapan methot.
        /// </summary>
        /// <param name="mainPart">Word sayfası</param>
        /// <param name="body">Sayfa içeriğinin ekleneceği alan</param>
        /// <param name="imageFile">Eklenecek görsel</param>
        private void CreateImage(MainDocumentPart mainPart, Body body, IFormFile imageFile)
        {
            string uploadDir = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");
            if (!Directory.Exists(uploadDir))
            {
                Directory.CreateDirectory(uploadDir);
            }
            //Yüklenen resimleri kaydet
            List<string> imageFiles = new List<string>();
            string imageFilePath = Path.Combine(uploadDir, imageFile.FileName);
            using (var stream = new FileStream(imageFilePath, FileMode.Create))
            {
                imageFile.CopyTo(stream);
            }
            imageFiles.Add(imageFilePath);
            AddImagesToDocument(mainPart, imageFiles, body);
        }

        /// <summary>
        /// DocumenModel dan gelen paragrafların sayfa da olşturulamsını sağlana methot.
        /// </summary>
        /// <param name="body">Eklenecek sayfa</param>
        /// <param name="text">Eklenecek Text(Paragraf)</param>
        private void CreateParagraph(Body body, string text)
        {
            RunProperties runPropertiesForContent = new RunProperties(
                 new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "20" }, // 10 punto
                new RunFonts() { Ascii = "Cambria" } // Yazı tipi
            );
            ParagraphProperties contentParagraphProperties = new ParagraphProperties(
                new Justification() { Val = JustificationValues.Both }, // İki yana yasla
                new SpacingBetweenLines() { Before = "250", After = "250" } // Paragraftan sonra boşluk
            );

            Run contentRun = new Run(new Text(text)) //Paragrafı ekle.
            {
                RunProperties = runPropertiesForContent
            };
            Paragraph contentParagraph = new Paragraph();
            contentParagraph.Append(contentParagraphProperties.CloneNode(true)); // Yeni özellikler ekleniyor
            contentParagraph.Append(contentRun);
            body.Append(contentParagraph);
        }

        /// <summary>
        /// Görselin bir kopyasını oluşturup belgeye eklenme işlemine yönlendirir.
        /// </summary>
        /// <param name="mainPart"></param>
        /// <param name="imageFiles"></param>
        /// <param name="body"></param>
        private void AddImagesToDocument(MainDocumentPart mainPart, List<string> imageFiles, Body body)
        {
            foreach (var imagePath in imageFiles)
            {
                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                using (FileStream stream = new FileStream(imagePath, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }
                AddImageToBody(mainPart, mainPart.GetIdOfPart(imagePart), body);
            }
        }
        /// <summary>
        /// Görseli belgeye isteniilen şekilde yerleştirir.
        /// </summary>
        /// <param name="mainPart"></param>
        /// <param name="relationshipId"></param>
        /// <param name="body"></param>
        private void AddImageToBody(MainDocumentPart mainPart, string relationshipId, Body body)
        {
            if (mainPart.Document.Body == null)
            {
                mainPart.Document.AppendChild(new Body());
            }

            // Resim partını al
            ImagePart imagePart = mainPart.GetPartById(relationshipId) as ImagePart;

            // Orijinal resim boyutlarını oku
            double originalWidth, originalHeight;
            using (Stream stream = imagePart.GetStream())
            {
                using (System.Drawing.Image image = System.Drawing.Image.FromStream(stream)) //System.Drawing.Common kütüphanesini indirdim.
                {
                    originalWidth = image.Width;
                    originalHeight = image.Height;
                }
            }

            // EMU cinsinden istenen genişlik (8cm) --> sayfa ikiye bölündüğünden bir kısmın genişliği bu kadar, boyutlandırma oranı buna göre yapılıyorr.
            long widthEMU = 2880000L;

            // En-boy oranına göre yüksekliği hesaplama işlemini yapıyor
            double aspectRatio = originalHeight / originalWidth;
            long heightEMU = (long)(widthEMU * aspectRatio);

            var element =
    new Paragraph(
        // Paragraf aralığını sıfırla
        new ParagraphProperties(
            new Justification() { Val = JustificationValues.Center }, //Resmi bulunduğu alanda ortalamak için.
            new SpacingBetweenLines() { Before = "0", After = "0" } // Resim ile açıklama arası boşluk kalmayacak

        ),
        new Run(
            new Drawing(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = widthEMU, Cy = heightEMU },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent()
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties()
                    {
                        Id = (UInt32Value)1U,
                        Name = "Picture"
                    },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                        new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks() { NoChangeAspect = true }),
                    new DocumentFormat.OpenXml.Drawing.Graphic(
                        new DocumentFormat.OpenXml.Drawing.GraphicData(
                            new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties()
                                    {
                                        Id = (UInt32Value)0U,
                                        Name = "New Bitmap Image.jpg"
                                    },
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
                                new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                    new DocumentFormat.OpenXml.Drawing.Blip()
                                    {
                                        Embed = relationshipId
                                    },
                                    new DocumentFormat.OpenXml.Drawing.Stretch(
                                        new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                                new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                    new DocumentFormat.OpenXml.Drawing.Transform2D(
                                        new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                                        new DocumentFormat.OpenXml.Drawing.Extents() { Cx = widthEMU, Cy = heightEMU }),
                                    new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                        new DocumentFormat.OpenXml.Drawing.AdjustValueList()
                                    )
                                    { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }))
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                )
                { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U }
            )
        )
    );

            body.AppendChild(element);
        }

        /// <summary>
        /// Kapak sayfası üzerindeki inputların belgeye eklendiği method.
        /// </summary>
        /// <param name="body"></param>
        private void AddCoverPage(Body body, DocumentModel model)
        {
            //Kapak sayfasındaki bütün alanlara eklemeler buradan yapılacak.

            //Formdan gelen alanlar
            //Sol ve sağ üstte eklenecek alanlar, api sorgusu ile ing çevirilerek ing kısmına da eklenecek.
            string university = ToTitleCase(model.coverPage.University);
            string placeOfPublication = ToTitleCase(model.coverPage.PlaceofPublication);
            string leftinfo = ToTitleCase(model.coverPage.Info); //(Cilt, Sayı No vb.) sol üstte üçüncü satır.

            //Doi numarası
            string doiNumber = model.coverPage.DoiNumber;

            //Her kelimenin baş harfi büyük hale getirilmesi için..
            string formattedArticleTitle = ToTitleCase(model.coverPage.ArticleTitle);

            string[] authors = model.coverPage.Authors.Split(",", StringSplitOptions.RemoveEmptyEntries); //Yazarlar virgünden sonra split yapılacak ve soyadları büyük hale getirilecek.
            List<string> formattedAuthors = new List<string>(); //Düzenlenmiş isimlerin ekleneceği liste.
            foreach (var author in authors)
            {
                formattedAuthors.Add(ToFormattedName(author));
            }

            string faculty = ToTitleCase(model.coverPage.Faculty);

            DateTime receivedDate = model.coverPage.ReceivedDate;
            DateTime accepteDate = model.coverPage.AcceptedDate;
            DateTime publishedOnlineDate = model.coverPage.PublishedOnlineDate;
            string fullOfDate = $"(Alınış / Received: {receivedDate.Date.ToString("d")}, Kabul / Accepted: {accepteDate.Date.ToString("d")}, Online Yayınlama / Published:{publishedOnlineDate.Date.ToString("d")})"; //Yalnızca gün/ay/yıl

            string[] keywords = model.coverPage.Keywords.Split(",", StringSplitOptions.RemoveEmptyEntries);
            List<string> formattedKeywords = new List<string>();
            foreach (var item in keywords)
            {
                //Hepsinin ilk harfi büyük yapılacak ve listeye eklenecek.
                formattedKeywords.Add(ToTitleCase(item));
            }

            string summary = model.coverPage.Abstract;



            //İngilizce olan kısımlar bunları çeviriden yapacağım.

            string universityEn = TranslateText(university);
            string placeOfPublicationEn = TranslateText(placeOfPublication);
            string rightinfo = TranslateText(leftinfo);

            string formattedArticleTitleEn = TranslateText(formattedArticleTitle);

            List<string> formattedKeywordsEn = new List<string>();
            foreach (var item in keywords)
            {
                formattedKeywordsEn.Add(TranslateText(item));
            }

            string summaryEn = TranslateText(summary);

            // Header tablosunu oluştur
            Table headerTable = CreateHeaderTable(university, placeOfPublication, leftinfo, universityEn, placeOfPublicationEn, rightinfo); //Formdan gelen veriler ile değiştirildi.

            // Word dokümanına tabloyu ekle
            body.Append(headerTable);

            AddDoiSection(body, doiNumber); //formdan gelen veriler ile değiştirildi.

            Table mainTitleTable = CreateTableForHeader(formattedArticleTitle, formattedAuthors, faculty);//Formdan gelen veri ile değiştirdi.
            body.Append(mainTitleTable);

            //İki tablo arasında boşluk olamlı bunu da info ile sağladım.
            Paragraph datesParagraph = AddInfoLine(fullOfDate, 18, "Cambria Bold", false);//Formdan gelen veri ile değiştirildi.
            body.Append(datesParagraph);

            Table summaryTable = CreateTableForSummary("Anahtar Kelimeler ", "Öz ", formattedKeywords, summary); //Formdan gelen veri ile değiştirildi.
            body.Append(summaryTable);

            Paragraph engTitle = EnglishTitle(formattedArticleTitleEn, 24, "Cambria Bold", true);
            body.Append(engTitle);

            //İngilizce özeet alanının eklenmesi.
            Table summaryTableEng = CreateTableForSummary("Keywords ", "Abstract ", formattedKeywordsEn, summaryEn);
            body.Append(summaryTableEng);
        }
        static string TranslateText(string text, string sourceLang = "tr", string targetLang = "en")
        {
            string url = $"https://api.mymemory.translated.net/get?q={Uri.EscapeDataString(text)}&langpair={sourceLang}|{targetLang}";

            using (HttpClient client = new HttpClient())
            {
                var response = client.GetAsync(url).GetAwaiter().GetResult();

                if (response.IsSuccessStatusCode)
                {
                    var json = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                    using JsonDocument doc = JsonDocument.Parse(json);
                    var translatedText = doc.RootElement
                                            .GetProperty("responseData")
                                            .GetProperty("translatedText")
                                            .GetString();
                    return translatedText;
                }
                else
                {
                    return $"Hata: {response.StatusCode}";
                }
            }
        }
        /// <summary>
        /// Girilen başlık alanının baş harflerini büyük hale getirmek için kullandım aynı zamanda anahtar kelimelerin baş harflerini de büyütmek için kullandım.
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string ToTitleCase(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return "";

            var words = input.Split(' ', StringSplitOptions.RemoveEmptyEntries);

            var capitalizedWords = words.Select(word =>
                char.ToUpper(word[0]) + word.Substring(1).ToLower());

            return string.Join(" ", capitalizedWords);
        }

        /// <summary>
        /// Yazar isimlerini uygun formaat getirme işlemi.
        /// </summary>
        /// <param name="fullName"></param>
        /// <returns></returns>
        public static string ToFormattedName(string fullName)
        {
            if (string.IsNullOrWhiteSpace(fullName))
                return "";
            //İsim ve soyisim birbirinden ayırdım.
            var parts = fullName.Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries);

            if (parts.Length == 0)
                return "";

            // İsim: ilk harfi büyük, diğerleri küçük
            string formattedFirstName = char.ToUpper(parts[0][0]) + parts[0].Substring(1).ToLower();

            // Soyisim: tamamen büyük (son kelime)
            string formattedLastName = parts.Length > 1
                ? parts[^1].ToUpper()
                : "";

            // Eğer 2'den fazla parça varsa (örneğin "mehmet ali taş"), ortadaki isimler de düzeltilmeli
            string middleNames = parts.Length > 2
                ? string.Join(" ", parts.Skip(1).Take(parts.Length - 2).Select(name =>
                    char.ToUpper(name[0]) + name.Substring(1).ToLower()))
                : "";

            return string.Join(" ", new[] { formattedFirstName, middleNames, formattedLastName }.Where(x => !string.IsNullOrEmpty(x)));
        }

        /// <summary>
        /// Kapak sayfasında sağ ve sol üstte bulunan bilgilerin eklendiği tablo.
        /// </summary>
        /// <param name="leftInfo1"></param>
        /// <param name="leftInfo2"></param>
        /// <param name="leftInfo3"></param>
        /// <param name="rightInfo1"></param>
        /// <param name="rightInfo2"></param>
        /// <param name="rightInfo3"></param>
        /// <returns></returns>
        private Table CreateHeaderTable(string university, string placeOfPublised, string leftinfo,
                                       string universityEn, string placeOfPublisedEn, string rightinfo)
        {
            // Yeni bir tablo oluştur
            Table headerTable = new Table();

            // Tablo özelliklerini ayarla
            TableProperties tableProperties = new TableProperties(
                new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" }, // %100 genişlik
                new TableBorders(
                    new TopBorder() { Val = BorderValues.None },
                    new LeftBorder() { Val = BorderValues.None },
                    new BottomBorder() { Val = BorderValues.None },
                    new RightBorder() { Val = BorderValues.None },
                    new InsideHorizontalBorder() { Val = BorderValues.None },
                    new InsideVerticalBorder() { Val = BorderValues.None }
                )
            );

            headerTable.AppendChild(tableProperties);

            // Tabloya bir satır ekle
            TableRow tableRow = new TableRow();

            // Sol hücre (Left Info)
            TableCell leftCell = new TableCell();
            leftCell.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Pct, Width = "2500" })); // %50 genişlik
            leftCell.Append(CreateParagraphSides(university));
            leftCell.Append(CreateParagraphSides(placeOfPublised));
            leftCell.Append(CreateParagraphSides(leftinfo));

            //Bunlar ing çevirilip eklencek tekrar farklı değişken alınmayacak.

            // Sağ hücre (Right Info)
            TableCell rightCell = new TableCell();
            rightCell.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Pct, Width = "2500" })); // %50 genişlik
            rightCell.Append(CreateParagraphSides(universityEn, true)); // Sağdan hizalanmış
            rightCell.Append(CreateParagraphSides(placeOfPublisedEn, true));
            rightCell.Append(CreateParagraphSides(rightinfo, true));

            // Satıra hücreleri ekle
            tableRow.Append(leftCell, rightCell);

            // Tabloya satırı ekle
            headerTable.Append(tableRow);

            return headerTable;
        }


        // Paragraf oluşturma fonksiyonu (Satır aralıklarını azaltma)
        private Paragraph CreateParagraphSides(string text, bool isRightAligned = false)
        {
            RunProperties runProperties = new RunProperties(
                new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "16" }, // 7 punto = 14 (OpenXML'de 2 ile çarpılır)
                new RunFonts() { Ascii = "Cambria" } // Yazı tipi Arial
            );

            Run run = new Run(new Text(text)) { RunProperties = runProperties };

            // Satır aralıklarını azalt
            SpacingBetweenLines spacing = new SpacingBetweenLines() { Before = "0", After = "0", Line = "200", LineRule = LineSpacingRuleValues.Auto };

            ParagraphProperties paragraphProperties = new ParagraphProperties(spacing);
            if (isRightAligned)
            {
                paragraphProperties.Justification = new Justification() { Val = JustificationValues.Right };
            }
            else
            {
                paragraphProperties.Justification = new Justification() { Val = JustificationValues.Left };
            }

            Paragraph paragraph = new Paragraph();
            paragraph.Append(paragraphProperties);
            paragraph.Append(run);

            return paragraph;
        }


        //DOİ bilgisi ekle
        private void AddDoiSection(Body body, string doiInfo)
        {
            // DOI paragrafını oluştur
            Paragraph doiParagraph = CreateDoiParagraph(doiInfo);
            body.Append(doiParagraph);
        }

        // DOI paragrafını oluşturan yardımcı metot
        private Paragraph CreateDoiParagraph(string doiInfo)
        {
            RunProperties runProperties = new RunProperties(
                new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "18" }, // 9 punto = 18 (OpenXML'de 2 ile çarpılır)
                new RunFonts() { Ascii = "Cambria" } // Yazı tipi Arial
            );

            Run run = new Run(new Text(doiInfo)) { RunProperties = runProperties };
            ParagraphProperties paragraphProperties = new ParagraphProperties(
                new SpacingBetweenLines() { Before = "300", After = "0" }, // Önce boşluk
                new Justification() { Val = JustificationValues.Right } // Sağa hizalama
            );

            Paragraph doiParagraph = new Paragraph();
            doiParagraph.Append(paragraphProperties);
            doiParagraph.Append(run);

            return doiParagraph;
        }

        //Başlık va yazar bilgisleri için tablo oluşturdum.
        private Table CreateTableForHeader(string title, List<string> authors, string faculty)
        {
            // Yeni bir tablo oluştur
            Table headerTable = new Table();

            // Tablo özelliklerini ayarla
            TableProperties tableProperties = new TableProperties(
                new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" }, // %100 genişlik
                new TableBorders(
                    new TopBorder() { Val = BorderValues.Single, Size = 8 },
                    new LeftBorder() { Val = BorderValues.None },
                    new BottomBorder() { Val = BorderValues.None },
                    new RightBorder() { Val = BorderValues.None },
                    new InsideHorizontalBorder() { Val = BorderValues.None },
                    new InsideVerticalBorder() { Val = BorderValues.None }
                )
            );

            headerTable.AppendChild(tableProperties);


            // Ortalamak için bir yardımcı fonksiyon oluştur
            TableCell CreateCenteredCell(string text, string fontSize = "24", bool isBold = false, string fontName = "Arial")
            {
                Run run = new Run(new Text(text));

                // Yazı tipi ve özellikleri ayarla
                RunProperties runProperties = new RunProperties(
                    new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = fontSize }, // Yazı boyutu
                    new RunFonts() { Ascii = fontName, HighAnsi = fontName, EastAsia = fontName, ComplexScript = fontName } // Yazı tipi
                );

                if (isBold)
                {
                    runProperties.AppendChild(new Bold()); // Kalın yazı
                }

                run.RunProperties = runProperties;

                Paragraph paragraph = new Paragraph(run)
                {
                    ParagraphProperties = new ParagraphProperties(new Justification() { Val = JustificationValues.Center }) // Ortalama
                };

                return new TableCell(paragraph)
                {
                    TableCellProperties = new TableCellProperties()
                };
            }

            // Başlık satırını ekle
            TableRow titleRow = new TableRow();
            titleRow.AppendChild(CreateCenteredCell(title, "24", true, "Cambria Bold")); // Büyük, kalın başlık ve özel font
            headerTable.AppendChild(titleRow);

            // Yazar bilgilerini ekle
            TableRow writesRow = new TableRow();
            writesRow.AppendChild(CreateCenteredCell(string.Join(", ", authors), "24", true, "Cambria Bold")); // Yazar bilgileri
            headerTable.AppendChild(writesRow);

            // Fakülte bilgilerini ekle
            TableRow facultyRow = new TableRow();
            facultyRow.AppendChild(CreateCenteredCell(faculty, "22", false, "Cambria Bold")); // Fakülte bilgileri
            headerTable.AppendChild(facultyRow);

            //Normalde info kısmı da buradaydı ancak daha sonra infoyu arada boşluk bırakabilmesi için ayrı olarak ekleme kararı aldım.
            return headerTable;
        }

        //İnfo line eklenmesi için oluşturdum.
        private Paragraph AddInfoLine(string infoText, double fontSize, string fontName, bool isBold)
        {

            // info metnini içeren bir Run nesnesi oluşturuluyor
            Run run = new Run(new Text(infoText));
            // Yazı tipi ve özelliklerini ayarlıyoruz
            RunProperties runProperties = new RunProperties(
                new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = fontSize.ToString() }, // Yazı boyutu
                new RunFonts() { Ascii = fontName, HighAnsi = fontName, EastAsia = fontName, ComplexScript = fontName } // Yazı tipi
            );
            if (isBold)
            {
                runProperties.AppendChild(new Bold()); // Yazıyı bold yap.
            }
            // Run nesnesinin özelliklerini ayarlıyoruz
            run.RunProperties = runProperties;

            // Üst ve alttan boşluk bırakıp yazıyı ortala.
            Paragraph paragraph = new Paragraph(run)
            {
                ParagraphProperties = new ParagraphProperties(new Justification() { Val = JustificationValues.Center }, // Ortalama
                new SpacingBetweenLines() { Before = "0", After = "200" })// Üstten ve alttan boşluk
            };
            return paragraph;


        }

        //özet alanları için tablo ing türkçe beraber kullanılacak. Parametrelere göre.
        private Table CreateTableForSummary(string titleKey, string titleSum, List<string> keywords, string summaryText)
        {
            Table summaryTable = new Table();

            TableProperties tableProperties = new TableProperties(
                new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" }, // %100 genişlik
                new TableBorders(
                    new TopBorder() { Val = BorderValues.Single, Size = 8 },
                    new LeftBorder() { Val = BorderValues.None },
                    new BottomBorder() { Val = BorderValues.Single, Size = 8 },
                    new RightBorder() { Val = BorderValues.None },
                    new InsideHorizontalBorder() { Val = BorderValues.None },
                    new InsideVerticalBorder() { Val = BorderValues.None }
                )
            );

            summaryTable.AppendChild(tableProperties);

            // Hücre oluşturucu (Bold ve FontSize ekleyen)
            TableCell CreateCell(string title, string content, string width, string titleFontSize = "18", string contentFontSize = "18")
            {
                // Başlık (Bold ve büyük font)
                RunProperties titleRunProperties = new RunProperties(new Bold(), new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = titleFontSize });
                Run titleRun = new Run(new Text(title)) { RunProperties = titleRunProperties };

                // İçerik (Normal font)
                RunProperties contentRunProperties = new RunProperties(new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = contentFontSize });
                Run contentRun = new Run(new Text(content)) { RunProperties = contentRunProperties };

                Paragraph paragraph = new Paragraph(titleRun, contentRun)
                {
                    ParagraphProperties = new ParagraphProperties(new Justification() { Val = JustificationValues.Both })
                };

                TableCellProperties cellProperties = new TableCellProperties(
                    new TableCellWidth() { Type = TableWidthUnitValues.Pct, Width = width }
                );

                return new TableCell(paragraph) { TableCellProperties = cellProperties };
            }
            // Ortak paragraf oluşturma metodu- özellikleri bir alanda topladım kod tekrarından kurtulmak için.
            Paragraph CreateStyledParagraph(string text, bool isBold = false)
            {
                RunProperties runProperties = new RunProperties(
                    new RunFonts() { Ascii = "Cambria", HighAnsi = "Cambria", EastAsia = "Cambria", ComplexScript = "Cambria" },
                    new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "18" }
                );

                if (isBold)
                {
                    runProperties.Append(new Bold());
                }

                Run run = new Run(new Text(text)) { RunProperties = runProperties };

                Paragraph paragraph = new Paragraph(run)
                {
                    ParagraphProperties = new ParagraphProperties(
                        new SpacingBetweenLines() { Before = "0", After = "0", Line = "200", LineRule = LineSpacingRuleValues.Auto },
                        new Justification() { Val = JustificationValues.Left }
                    )
                };

                return paragraph;
            }

            // Satır oluştur ve iki hücre ekle
            TableRow row = new TableRow();
            TableCell leftCell = new TableCell();
            leftCell.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Pct, Width = "1250" }));
            // Başlık ekle (titleKey)
            leftCell.Append(CreateStyledParagraph(titleKey + ":", true));

            // Anahtar kelimeler ekleniyor
            if (keywords.Count > 0)
            {
                for (int i = 0; i < keywords.Count; i++)
                {
                    string keywordText = (i < keywords.Count - 1) ? keywords[i] + ", " : keywords[i];
                    leftCell.Append(CreateStyledParagraph(keywordText));
                }
            }
            else
            {
                leftCell.Append(CreateStyledParagraph("N/A"));
            }


            // Sağ hücre (Özet içeriği)
            TableCell rightCell = CreateCell(titleSum + ": ", summaryText, "3750", "18", "18");

            row.Append(leftCell, rightCell);
            summaryTable.AppendChild(row);

            return summaryTable;
        }

        private Paragraph EnglishTitle(string titleText, double fontSize, string fontName, bool isBold)
        {
            // Başlık metnini içeren bir Run nesnesi oluşturuluyor
            Run run = new Run(new Text(titleText));

            // Yazı tipi ve özelliklerini ayarlıyoruz
            RunProperties runProperties = new RunProperties(
                new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = fontSize.ToString() }, // Yazı boyutu
                new RunFonts() { Ascii = fontName, HighAnsi = fontName, EastAsia = fontName, ComplexScript = fontName } // Yazı tipi
            );

            // Eğer kalın yazı isteniyorsa, kalın yazı özelliği ekliyoruz
            if (isBold)
            {
                runProperties.AppendChild(new Bold()); // Kalın yazı
            }

            // Run nesnesinin özelliklerini ayarlıyoruz
            run.RunProperties = runProperties;

            // Son olarak, başlık içeren bir paragraf oluşturuyoruz ve ortalıyoruz
            Paragraph paragraph = new Paragraph(run)
            {
                ParagraphProperties = new ParagraphProperties(new Justification() { Val = JustificationValues.Center }, // Ortalama
                new SpacingBetweenLines() { Before = "200", After = "200" }) // Daha dar boşluk // Üstten ve alttan boşluk
            };
            return paragraph;
        }



    }
}
