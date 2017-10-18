using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Excel.AddIn.Shape2Image
{
    class ShapeBuffer
    {
        public Shape Shape { get; private set; }
        public ImageFormat PictureFormat { get; set; }

        public ShapeBuffer(Shape shape)
        {
            Shape = shape;
            PictureFormat = ImageFormat.Jpeg;
        }

        public Image ReadClipboard()
        {
            // クリップボードにあるデータを取得する
            var iData = Clipboard.GetDataObject();

            // クリップボードにBitmapファイルがあれば
            if (iData.GetDataPresent(DataFormats.Bitmap))
            {
                return iData.GetData(DataFormats.Bitmap) as Image;
            }
            return null;
        }

        public void WriteClipboard()
        {
            // EXCELの「図形のコピー」でビットマップ形式でコピーする
            Shape.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap);
        }

        public void SaveAsPicture(string file)
        {
            try
            {
                WriteClipboard();
                var img = ReadClipboard();
                if (img == null)
                {
                    return;
                }
                img.Save(file, PictureFormat);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                Clipboard.Clear();
            }
        }
    }
}
