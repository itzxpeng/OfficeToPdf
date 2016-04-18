using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Convertion
{
    public class ConvertToImages
    {

        public int StartIndex { get; set; }

        public int EndIndex { get; set; }

        public int DPI { get; set; }

        public ConvertToImages()
        {
            DPI = 96;
            StartIndex = int.MinValue;
            EndIndex = int.MaxValue;
        }

        public string Generate(string outputpath, string format, Presentation presentation)
        {
            if (presentation != null)
            {
                var width = GetSlideWidth(presentation);
                var height = GetSlideHeight(presentation);
                var count = StartIndex == int.MinValue ? 0 : StartIndex;
                foreach (Slide slide in presentation.Slides.OfType<Slide>())
                {
                    //if (count < StartIndex || count > EndIndex)
                    //{
                    //    count += 1;
                    //    continue;
                    //}

                    var imagepath = System.IO.Path.Combine(outputpath, GetImageName(slide, format));
                    slide.Export(imagepath, format, width, height);

                    var info = new PreviewImageInfo
                    {
                        OrignalIndex = count,
                        UniqeId = slide.SlideID.ToString(),
                        ImageDatas = File.ReadAllBytes(imagepath)
                    };
                    imageinfos.Add(info);
                    count += 1;
                }
            }
            return outputpath;
        }

        private List<PreviewImageInfo> imageinfos = new List<PreviewImageInfo>();

        public ICollection<PreviewImageInfo> GeneratedImagesInfo
        {
            get { return imageinfos; }
        }

        private string GetImageName(Slide slide, string format)
        {
            return string.Format("Slide{0}.{1}", slide.SlideIndex, format);
        }

        private int GetSlideWidth(Presentation presentation)
        {
            return (int)(this.DPI / 96.0 / 0.75 * presentation.PageSetup.SlideWidth);
        }

        private int GetSlideHeight(Presentation presentation)
        {
            return (int)(this.DPI / 96.0 / 0.75 * presentation.PageSetup.SlideHeight);
        }
    }
}
