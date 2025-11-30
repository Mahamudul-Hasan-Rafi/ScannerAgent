using System;
using System.IO;
using WIA;
using ScannerAgent.Model;

namespace ScannerAgent.Utils
{
    public class ScannerUtils
    {

        public static ScannedImage ConvertImageToBase64(ImageFile imageFile, int pageNumber)
        {
            //string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".tmp");

            try
            {
                //imageFile.SaveFile(tempPath);
                //byte[] imageBytes = File.ReadAllBytes(tempPath);

                // Use the get_BinaryData() accessor as required by CS1545
                var binaryData = imageFile.FileData.get_BinaryData();
                // The returned object is typically a byte[] or a COM object that can be cast to MemoryStream
                byte[] imageBytes;
                if (binaryData is byte[] bytes)
                {
                    imageBytes = bytes;
                }
                else if (binaryData is MemoryStream ms)
                {
                    imageBytes = ms.ToArray();
                }
                else
                {
                    // Fallback: try to convert to byte[]
                    imageBytes = (byte[])System.Runtime.InteropServices.Marshal.GetObjectForIUnknown(
                        System.Runtime.InteropServices.Marshal.GetIUnknownForObject(binaryData));
                }
               
                string base64 = Convert.ToBase64String(imageBytes);

                string mimeType = "image/jpeg";
                if (imageFile.FormatID == WiaFormatID.wiaFormatPNG) mimeType = "image/png";
                else if (imageFile.FormatID == WiaFormatID.wiaFormatTIFF) mimeType = "image/tiff";
                else if (imageFile.FormatID == WiaFormatID.wiaFormatBMP) mimeType = "image/bmp";

                return new ScannedImage
                {
                    PageNumber = pageNumber,
                    Base64Data = base64, // $"data:{mimeType};base64,{base64}",
                    Size = imageBytes.Length,
                    Format = mimeType
                };
            }
            finally
            {
                //try { File.Delete(tempPath); } catch { }
            }
        }


        public static string GetWiaFormat(string format)
        {
            switch (format.ToLower())
            {
                case "png": return WiaFormatID.wiaFormatPNG;
                case "tiff": return WiaFormatID.wiaFormatTIFF;
                case "bmp": return WiaFormatID.wiaFormatBMP;
                default: return WiaFormatID.wiaFormatJPEG;
            }
        }


        public static string EscapeJsonString(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            return s.Replace("\\", "\\\\")
                    .Replace("\"", "\\\"")
                    .Replace("\n", "\\n")
                    .Replace("\r", "\\r");
        }


        public static void ReleaseCom(object comObj)
        {
            if (comObj == null) return;
            try
            {
                while (System.Runtime.InteropServices.Marshal.ReleaseComObject(comObj) > 0) { }
            }
            catch { }
        }

    }
}
