using System;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace DrawioPpt.Core.Services
{
    public static class CompressionUtility
    {
        public static string CompressToBase64(string content)
        {
            if (string.IsNullOrEmpty(content))
            {
                return string.Empty;
            }

            byte[] rawBytes = Encoding.UTF8.GetBytes(content);
            using (MemoryStream output = new MemoryStream())
            {
                using (GZipStream gzip = new GZipStream(output, CompressionMode.Compress, true))
                {
                    gzip.Write(rawBytes, 0, rawBytes.Length);
                }

                return Convert.ToBase64String(output.ToArray());
            }
        }

        public static string DecompressFromBase64(string compressedContent)
        {
            if (string.IsNullOrEmpty(compressedContent))
            {
                return string.Empty;
            }

            byte[] compressedBytes = Convert.FromBase64String(compressedContent);
            using (MemoryStream input = new MemoryStream(compressedBytes))
            using (GZipStream gzip = new GZipStream(input, CompressionMode.Decompress))
            using (MemoryStream output = new MemoryStream())
            {
                gzip.CopyTo(output);
                return Encoding.UTF8.GetString(output.ToArray());
            }
        }
    }
}

