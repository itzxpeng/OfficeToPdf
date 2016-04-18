/* ==============================================================================
 * Description：ConvertFromPPTX  
 * Author：peytonzhang
 * Created：3/28/2016 11:13:09 AM
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
    public class ConvertFromPPTX
    {
        Application _pptApplication = new Application();
        public bool ConvertToImages(string tempPresentationFullPath, string outpath)
        {
            
            try
            {
                Microsoft.Office.Interop.PowerPoint.Presentation preso = _pptApplication.Presentations.Open2007(tempPresentationFullPath, MsoTriState.msoFalse,
                            MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoTrue);
                ConvertToImages cti = new ConvertToImages();
                //string outpath = AppDomain.CurrentDomain.BaseDirectory + "Test1";
                if (!System.IO.Directory.Exists(outpath))
                    System.IO.Directory.CreateDirectory(outpath);
                string result = cti.Generate(outpath, "jpg", preso);
                if (!String.IsNullOrEmpty(result))
                    return true;
                else
                {
                    return false;
                }
            }catch(Exception ex)
            {
                throw ex.InnerException;
            }
        }

        public bool ConvertToPdf(string tempPresentationFullPath,string outpath)
        {
            try
            {
                Microsoft.Office.Interop.PowerPoint.Presentation preso = null;
                preso = _pptApplication.Presentations.Open2007(tempPresentationFullPath, MsoTriState.msoFalse,
                        MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoTrue);
                string name = System.IO.Path.GetFileNameWithoutExtension(tempPresentationFullPath);
                if (!System.IO.Directory.Exists(outpath))
                    System.IO.Directory.CreateDirectory(outpath);

                string outPutPath = System.IO.Path.Combine(outpath, name + ".pdf");

                ConvertToPdf pdf = new ConvertToPdf();
                string opath = pdf.Generate(outPutPath, preso);
                if (!String.IsNullOrEmpty(opath))
                    return true;
                else
                    return false;
            }catch(Exception ex)
            {
                throw ex.InnerException;
            }
        }
    }
}
