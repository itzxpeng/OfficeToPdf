/* ==============================================================================
 * Description：ConvertToPdf  
 * Author：peytonzhang
 * Created：3/28/2016 1:24:10 PM
 * ==============================================================================*/

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Convertion
{
    public class ConvertToPdf
    {
        public string Generate(string outPutPath, Presentation preso)
        {

            PpPrintRangeType printRangeType = PpPrintRangeType.ppPrintAll;
            PrintRange printRange = null;
            preso.ExportAsFixedFormat(
                outPutPath,
                PpFixedFormatType.ppFixedFormatTypePDF,
                PpFixedFormatIntent.ppFixedFormatIntentPrint,
                MsoTriState.msoFalse,
                PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
                PpPrintOutputType.ppPrintOutputSlides,
                MsoTriState.msoFalse,
                printRange,
                printRangeType,
                null,
                false,
                true,
                false, //bool DocStructureTags = true
                true,
                false,
                Type.Missing);

            return outPutPath;
        }
    }
}
