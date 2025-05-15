using DocumentFormat.OpenXml.Presentation;
using System.ComponentModel.DataAnnotations;

namespace WordDemo5.Models
{
    public class DocumentModel //Başlıkları alacak olacak döküman modeli
    {
        //    /*Sections 
        //      * Section1 --> cover page
        //        Section2 --> introduction page
        //     */
        public CoverPage coverPage { get; set; }= new CoverPage();
        public List<ContentItem> ContentItems { get; set; } = new List<ContentItem>();
    }
    public class CoverPage
    {
        public string University { get; set; } //Üni adı
        public string PlaceofPublication { get; set; } //Yayın yeri
        public string Info { get; set; }//Bilgi (Cilt, Sayı No vb.)
        public string DoiNumber { get; set; }//DOI bilgisi
        public string ArticleTitle { get; set; }//Makale başlığı
        public string Authors { get; set; }//Yazarlar
        public string Faculty { get; set; }//Fakülte adı
        public DateTime ReceivedDate { get; set; }//Alınış tarihi
        public DateTime AcceptedDate { get; set; }//Kabul edilme tarihi
        public DateTime PublishedOnlineDate { get; set; }//Online Yayınlanma tarihi
        public string Keywords { get; set; }//Anahtar kelimeler
        public string Abstract { get; set; }//Özet metni
    }
    public class ContentItem
    {
        [Required(ErrorMessage = "Başlık zorunludur")]
        public string Title { get; set; }
        public List<ContentBlock> ContentBlocks { get; set; } = new List<ContentBlock>();
    }
    public class ContentBlock
    {
        public string ContentType { get; set; } // "Paragraph", "Image" ,"ImageExplanation".....
        [Required(ErrorMessage = "Paragraf metni boş olamaz.")]
        public string ParagraphText { get; set; }
        public IFormFile? ImageFile { get; set; } //? --> Null olabilir.
        public string? ImageCaptionText { get; set; }
        public string? TableCaptionText { get; set; } // ? Boş bırakılabilir.
        public IFormFile? TableFile { get; set; } //Tablo --> Null olabilir.
        public string? SubTitleText { get; set; } //Alt başlık --> isteğe bağlı.
        public string? SubSubTitleText { get; set; } //Alt başlık için alt başlık ekle --> isteğe bağlı.

    }



}
