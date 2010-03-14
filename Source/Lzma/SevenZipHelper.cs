// This file is from Peter Bromberg - 'Using Compressed Resources in .NET with LZMA (7-Zip) Compression'

using System;
using System.IO;
using System.Reflection;
namespace SevenZip.Compression.LZMA
{
    public static class SevenZipHelper
    {
       static int dictionary = 1 << 23; // 1 << 23;
      // static Int32 posStateBits = 2;
     // static  Int32 litContextBits = 3; // for normal files
        // UInt32 litContextBits = 0; // for 32-bit data
     // static  Int32 litPosBits = 0;
        // UInt32 litPosBits = 2; // for 32-bit data
    // static   Int32 algorithm = 2;
    // static    Int32 numFastBytes = 128;

        /*
         <Properties>
  
   dictionary - [0,28], default: 23 (2^23 = 8MB)

  numFastBytes: set number of fast bytes - [5, 255], default: 128
          Usually big number gives a little bit better compression ratio 
          and slower compression process.

   listContextBits: set number of literal context bits - [0, 8], default: 3
          Sometimes lc=4 gives gain for big files.

  litPosBits: set number of literal pos bits - [0, 4], default: 0
          lp switch is intended for periodical data when period is 
          equal 2^value (where lp=value). For example, for 32-bit (4 bytes) 
          periodical data you can use lp=2. Often it's better to set lc=0, 
          if you change lp switch.

  -pb{N}: set number of pos bits - [0, 4], default: 2
          pb switch is intended for periodical data 
          when period is equal 2^value (where lp=value).

  -eos:   write End Of Stream marker
         */



     static   bool eos = false;
     static   CoderPropID[] propIDs = 
				{
					CoderPropID.DictionarySize,
					CoderPropID.PosStateBits,
					CoderPropID.LitContextBits,
					CoderPropID.LitPosBits,
					CoderPropID.Algorithm,
					CoderPropID.NumFastBytes,
					CoderPropID.MatchFinder,
					CoderPropID.EndMarker
				};

        // these are the default properties, keeping it simple for now:
     static   object[] properties = 
				{
					(Int32)(dictionary),
					(Int32)(2), /* PosStateBits 2 */
					(Int32)(3), /* LitContextBits 3 */
					(Int32)(0), /* LitPosBits 0 */
					(Int32)(2), /*Algorithm  2 */
					(Int32)(128), /* NumFastBytes 128 */
					"bt4", /* MatchFinder "bt4" */
					eos   /* endMarker  eos */
				};

        public static byte[] Compress(byte[] inputBytes)
        {
            MemoryStream inStream = new MemoryStream(inputBytes);
            MemoryStream outStream = new MemoryStream();
            Encoder encoder = new Encoder();
            encoder.SetCoderProperties(propIDs, properties);
            encoder.WriteCoderProperties(outStream);
            long fileSize = inStream.Length;
            for (int i = 0; i < 8; i++)
                outStream.WriteByte((Byte)(fileSize >> (8 * i)));
            encoder.Code(inStream, outStream, -1, -1, null);
            return outStream.ToArray();
        }


	   // public static byte[] GetDecompressedResourceFromAssembly(Assembly assembly, string resourceName)
	   // {
	   //     // Get the resource
	   //     Stream str = assembly.GetManifestResourceStream(resourceName);
	   //     byte[] b = new byte[(int) str.Length];
	   //     str.Read(b, 0, b.Length);
	   //     // decompress the resource
	   //     byte[] b2 = Decompress(b);
	   //    return b2;
	   //}


        public static byte[] Decompress(byte[] inputBytes)
        {
            MemoryStream newInStream = new MemoryStream(inputBytes);
            Decoder decoder = new Decoder();            
            newInStream.Seek(0, 0);
            MemoryStream newOutStream = new MemoryStream();
            byte[] properties2 = new byte[5];
            if (newInStream.Read(properties2, 0, 5) != 5)
                throw (new Exception("input .lzma is too short"));
            long outSize = 0;
            for (int i = 0; i < 8; i++)
            {
                int v = newInStream.ReadByte();
                if (v < 0)
                    throw (new Exception("Can't Read 1"));
                outSize |= ((long)(byte)v) << (8 * i);
            }
            decoder.SetDecoderProperties(properties2);
            long compressedSize = newInStream.Length - newInStream.Position;
            decoder.Code(newInStream, newOutStream, compressedSize, outSize, null);
            byte[] b = newOutStream.ToArray();
            return b;
        }
    }
}