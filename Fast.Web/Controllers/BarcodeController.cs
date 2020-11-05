using System.Drawing;
using System.IO;
using System.Web.Mvc;
using ZXing;
using System.Drawing.Imaging;

namespace Fast.Web.Controllers
{
    public class BarcodeController : Controller
    {
        // GET: Barcode
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult Content()
        {
            return View();
        }
        public ActionResult RenderBarcode(string sample)
        {
            Image img = null;
            using (var x = new MemoryStream())
            {
                var res = new BarcodeWriter() { Format = BarcodeFormat.CODE_128 };

                res.Options.Height = 80;
                res.Options.Width = 280;
                res.Options.PureBarcode = false;

                img = res.Write(sample);
                img.Save(x, System.Drawing.Imaging.ImageFormat.Jpeg);

                return File(x.ToArray(), "image/jpeg");
            }
        }

        public ActionResult RenderBarcodeBig(string sample)
        {
            Image img = null;
            using (var x = new MemoryStream())
            {
                var res = new BarcodeWriter() { Format = BarcodeFormat.CODE_128 };

                res.Options.Height = 150;
                res.Options.Width = 400;
                res.Options.PureBarcode = false;

                img = res.Write(sample);
                img.Save(x, System.Drawing.Imaging.ImageFormat.Jpeg);

                return File(x.ToArray(), "image/jpeg");
            }

        }
        public ActionResult DrawString(string text)
        {
            using (var x = new MemoryStream())
            {
                //first, create a dummy bitmap just to get a graphics object
                Image img = new Bitmap(100, 100);
                Graphics drawing = Graphics.FromImage(img);

                //measure the string to see how big the image needs to be
                SizeF textSize = drawing.MeasureString(text, SystemFonts.DefaultFont);

                //free up the dummy image and old graphics object
                img.Dispose();
                drawing.Dispose();

                //create a new image of the right size
                img = new Bitmap((int)textSize.Width, (int)textSize.Height);

                drawing = Graphics.FromImage(img);

                //paint the background
                drawing.Clear(Color.White);

                //create a brush for the text
                Brush textBrush = new SolidBrush(Color.Black);

                drawing.DrawString(text, SystemFonts.DefaultFont, textBrush, 0, 0);

                img.Save(x, System.Drawing.Imaging.ImageFormat.Jpeg);

                textBrush.Dispose();
                drawing.Dispose();

                return File(x.ToArray(), "image/jpeg");
            }
        }
    }
}