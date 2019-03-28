using System;
using System.Diagnostics.Contracts;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Collections;
using System.Diagnostics;
using System.Text;

namespace Kavod.Vba.Compression
{
    /// <summary>
    /// A CompressedChunk is a record that encodes all data from a DecompressedChunk (section 
    /// 2.4.1.1.3) in compressed format. A CompressedChunk has two parts: a CompressedChunkHeader 
    /// (section 2.4.1.1.5) followed by a CompressedChunkData (section 2.4.1.1.6). The number of bytes 
    /// in a CompressedChunk MUST be greater than or equal to 3. The number of bytes in a 
    /// CompressedChunk MUST be less than or equal to 4098.
    /// </summary>
    /// <remarks></remarks>
    internal class CompressedChunk
    {
        internal CompressedChunk(DecompressedChunk decompressedChunk)
        {
//            Contract.Requires<ArgumentNullException>(decompressedChunk != null);
//            Contract.Ensures(Header != null);
//            Contract.Ensures(ChunkData != null);

            ChunkData = new CompressedChunkData(decompressedChunk);
            if (ChunkData.Size >= Globals.MaxBytesPerChunk)
            {
                ChunkData = new RawChunk(decompressedChunk.Data);
            }
            Header = new CompressedChunkHeader(ChunkData);
        }

        internal CompressedChunk(BinaryReader dataReader)
        {
//            Contract.Requires<ArgumentNullException>(dataReader != null);
//            Contract.Ensures(Header != null);
//            Contract.Ensures(ChunkData != null);

            Header = new CompressedChunkHeader(dataReader);
            if (Header.IsCompressed)
            {
                ChunkData = new CompressedChunkData(dataReader, Header.CompressedChunkDataSize);
            }
            else
            {
                ChunkData = new RawChunk(dataReader.ReadBytes(Header.CompressedChunkDataSize));
            }
        }

        internal CompressedChunkHeader Header { get; }

        internal IChunkData ChunkData { get; }

        internal byte[] SerializeData()
        {
            var serializedHeader = Header.SerializeData();
            var serializedChunkData = ChunkData.SerializeData();

            var data = serializedHeader.Concat(serializedChunkData);
            if (!Header.IsCompressed)
            {
                var dataLength = serializedHeader.LongLength + serializedChunkData.LongLength;
                var paddingLength = Globals.NumberOfChunkHeaderBytes
                                    + Globals.MaxBytesPerChunk
                                    - dataLength;
                var padding = Enumerable.Repeat(Globals.PaddingByte, (int)paddingLength);
                data = data.Concat(padding);
            }
            return data.ToArray();
        }
    }

    /// <summary>
    /// If CompressedChunkHeader.CompressedChunkFlag (section 2.4.1.1.5) is 0b0, CompressedChunkData 
    /// contains an array of CompressedChunkHeader.CompressedChunkSize elements plus 3 bytes of 
    /// uncompressed data.  If CompressedChunkHeader CompressedChunkFlag is 0b1, CompressedChunkData 
    /// contains an array of TokenSequence (section 2.4.1.1.7) elements.
    /// </summary>
    /// <remarks></remarks>
    internal class CompressedChunkData : IChunkData
    {
        private readonly List<TokenSequence> _tokensequences = new List<TokenSequence>();

        internal CompressedChunkData(DecompressedChunk chunk)
        {
//            Contract.Requires<ArgumentNullException>(chunk != null);
            
            var tokens = Tokenizer.TokenizeUncompressedData(chunk.Data);
            _tokensequences.AddRange(tokens.ToTokenSequences());
        }

        internal CompressedChunkData(BinaryReader dataReader, UInt16 compressedChunkDataSize)
        {
            var data = dataReader.ReadBytes(compressedChunkDataSize);

            using (var reader = new BinaryReader(new MemoryStream(data)))
            {
                var position = 0;
                while (reader.BaseStream.Position < reader.BaseStream.Length)
                {
                    var sequence = TokenSequence.GetFromCompressedData(reader, position);
                    _tokensequences.Add(sequence);
                    position += (int)sequence.Tokens.Sum(t => t.Length);
                }
            }
        }

        internal IEnumerable<TokenSequence> TokenSequences => _tokensequences;

        public byte[] SerializeData()
        {
            // get data from TokenSequences.
            var data = from t in _tokensequences
                       from d in t.SerializeData()
                       select d;
            return data.ToArray();
        }

        // TODO this is probably really inefficient.
        public int Size => SerializeData().Length;
    }

    internal class CompressedChunkHeader
    {
        internal CompressedChunkHeader(IChunkData chunkData)
        {
            IsCompressed = chunkData is CompressedChunkData;
            CompressedChunkSize = (ushort) (chunkData.Size + 2);
        }

        internal CompressedChunkHeader(UInt16 header)
        {
            DecodeHeader(header);
        }

        internal CompressedChunkHeader(BinaryReader dataReader)
        {
            var header = dataReader.ReadUInt16();
            DecodeHeader(header);
        }

        private void DecodeHeader(UInt16 header)
        {
            var temp = (UInt16)(header & 0xf000);
            switch (temp)
            {
                case 0xb000:
                    IsCompressed = true;
                    break;

                case 0x3000:
                    IsCompressed = false;
                    break;

                default:
                    throw new Exception();
            }

            // 2.4.1.3.12 Extract CompressedChunkSize
            // SET temp TO Header BITWISE AND 0x0FFF
            // SET Size TO temp PLUS 3
            CompressedChunkSize = (UInt16)((header & 0xfff) + 3);

            ValidateChunkSizeAndCompressedFlag();
        }

        internal bool IsCompressed { get; private set; }

        internal UInt16 CompressedChunkSize { get; private set; }

        internal UInt16 CompressedChunkDataSize => (UInt16)(CompressedChunkSize - 2);

        internal byte[] SerializeData()
        {
            ValidateChunkSizeAndCompressedFlag();

            UInt16 header;
            if (IsCompressed)
            {
                header = (UInt16)(0xb000 | (CompressedChunkSize - 3));
            }
            else
            {
                header = (UInt16)(0x3000 | (CompressedChunkSize - 3));
            }
            return BitConverter.GetBytes(header);
        }

        private void ValidateChunkSizeAndCompressedFlag()
        {
            if (IsCompressed 
                && CompressedChunkSize > 4098)
            {
                throw new Exception();
            }
            if (!IsCompressed 
                && CompressedChunkSize != 4098)
            {
                throw new Exception();
            }
        }
    }

    /// <summary>
    /// A CompressedContainer is an array of bytes holding the compressed data. The Decompression 
    /// algorithm (section 2.4.1.3.1) processes a CompressedContainer to populate a DecompressedBuffer. 
    /// The Compression algorithm (section 2.4.1.3.6) processes a DecompressedBuffer to produce a 
    /// CompressedContainer.  A CompressedContainer MUST be the last array of bytes in a stream (1). 
    /// On read, the end of stream (1) indicator determines when the entire CompressedContainer has 
    /// been read.  The CompressedContainer is a SignatureByte followed by array of CompressedChunk 
    /// (section 2.4.1.1.4) structures.
    /// </summary>
    /// <remarks></remarks>
    internal class CompressedContainer
    {
        private const byte SignatureByteSig = 0x1;

        private readonly List<CompressedChunk> _compressedChunks = new List<CompressedChunk>();
        
        internal CompressedContainer(byte[] compressedData)
        {
            var reader = new BinaryReader(new MemoryStream(compressedData));

            if (reader.ReadByte() != SignatureByteSig)
            {
                throw new Exception();
            }

            while (reader.BaseStream.Position < reader.BaseStream.Length)
            {
                _compressedChunks.Add(new CompressedChunk(reader));
            }
        }

        internal CompressedContainer(DecompressedBuffer buffer)
        {
            foreach (var chunk in buffer.DecompressedChunks)
            {
                _compressedChunks.Add(new CompressedChunk(chunk));
            }
        }

        internal IEnumerable<CompressedChunk> CompressedChunks => _compressedChunks;

        internal byte[] SerializeData()
        {
            using (var writer = new BinaryWriter(new MemoryStream()))
            {
                writer.Write(SignatureByteSig);

                foreach (var chunk in CompressedChunks)
                {
                    writer.Write(chunk.SerializeData());
                }

                using (var reader = new BinaryReader(writer.BaseStream))
                {
                    reader.BaseStream.Position = 0;
                    return reader.ReadBytes((int) reader.BaseStream.Length);
                }
            }
        }
    }

    /// <summary>
    /// CopyToken is a two-byte record interpreted as an unsigned 16-bit integer in little-endian 
    /// order. A CopyToken is a compressed encoding of an array of bytes from a DecompressedChunk 
    /// (section 2.4.1.1.3). The byte array encoded by a CopyToken is a byte-for-byte copy of a byte 
    /// array elsewhere in the same DecompressedChunk, called a CopySequence (section 2.4.1.3.19).  
    /// 
    /// The starting location, in a DecompressedChunk, is determined by the Compressing a Token 
    /// (section 2.4.1.3.9) and the Decompressing a Token (section 2.4.1.3.5) algorithms. Packed into 
    /// the CopyToken is the Offset, the distance, in byte count, to the beginning of the CopySequence. 
    /// Also packed into the CopyToken is the Length, the number of bytes encoded in the CopyToken. 
    /// Length also specifies the count of bytes in the CopySequence. The values encoded in Offset and 
    /// Length are computed by the Matching (section 2.4.1.3.19.4) algorithm.
    /// </summary>
    /// <remarks></remarks>
    internal class CopyToken : IToken, IEquatable<CopyToken>
    {
        private readonly UInt16 _tokenOffset;
        private readonly UInt16 _tokenLength;

        /// <summary>
        /// Constructor used to create a CopyToken when compressing a DecompressedChunk.
        /// </summary>
        /// <param name="tokenPosition">
        /// The start position of the CopyToken decompressed data in the current DecompressedChunk.
        /// </param>
        /// <param name="tokenOffset">
        /// The offset in bytes from the start position in the current DecompressedChunk from which to 
        /// start copying.
        /// </param>
        /// <param name="tokenLength">The number of bytes to copy from the offset.</param>
        /// <remarks></remarks>

        internal CopyToken(long tokenPosition, UInt16 tokenOffset, UInt16 tokenLength)
        {
            Position = tokenPosition;
            _tokenOffset = tokenOffset;
            _tokenLength = tokenLength;
        }

        /// <summary>
        /// Constructor used to create CopyToken instance when reading compressed token from a stream.
        /// </summary>
        /// <param name="dataReader">
        /// A BinaryReader object where the position is located at an encoded CopyToken.
        /// </param>
        /// <remarks></remarks>
        internal CopyToken(BinaryReader dataReader, long position)
        {
            Position = position;
            CopyToken.UnPack(dataReader.ReadUInt16(), Position, out _tokenOffset, out _tokenLength);
        }

        public long Length => _tokenLength;

        internal UInt16 Offset => _tokenOffset;

        internal long Position { get; }

        internal static UInt16 Pack(long position, UInt16 offset, UInt16 length)
        {
            // 2.4.1.3.19.3 Pack CopyToken
            var result = CopyTokenHelp(position);

            if (length > result.MaximumLength)
                throw new Exception();

            //SET temp1 TO Offset MINUS 1
            var temp1 = (UInt16)(offset - 1);

            //SET temp2 TO 16 MINUS BitCount
            var temp2 = (UInt16)(16 - result.BitCount);

            //SET temp3 TO Length MINUS 3
            var temp3 = (UInt16)(length - 3);

            //SET Token TO (temp1 LEFT SHIFT BY temp2) BITWISE OR temp3
            return (UInt16)((temp1 << temp2) | temp3);
        }

        public void DecompressToken(BinaryWriter writer)
        {
            // It is possible that the length is greater than the offset which means we would need to
            // read more bytes than are available.  To handle this we need to read the bytes available
            // (ie Offset amount) and then pad the remaining length with copies of the data read from 
            // the beginning of the buffer.

            var streamPosition = writer.BaseStream.Position;
            var reader = new BinaryReader(writer.BaseStream, Encoding.Unicode, true);
            reader.BaseStream.Position = streamPosition - _tokenOffset;
            var copySequence = reader.ReadBytes(Math.Min(_tokenOffset, _tokenLength));

            Array.Resize(ref copySequence, _tokenLength);

            for (int i = _tokenOffset; i <= _tokenLength - 1; i++)
            {
                var copyByte = copySequence[i % _tokenOffset];
                copySequence[i] = copyByte;
            }

            // Move the position of the underlying stream back to the original position and write the
            // CopySequence.
            writer.BaseStream.Position = streamPosition;
            writer.Write(copySequence);
        }

        internal static void UnPack(UInt16 packedToken, long position, out UInt16 unpackedOffset, out UInt16 unpackedLength)
        {
            // CALL CopyToken Help (section 2.4.1.3.19.1) returning LengthMask, OffsetMask, and BitCount.
            var result = CopyToken.CopyTokenHelp(position);

            // SET Length TO (Token BITWISE AND LengthMask) PLUS 3.
            unpackedLength = (UInt16)((packedToken & result.LengthMask) + 3);

            // SET temp1 TO Token BITWISE AND OffsetMask.
            var temp1 = (UInt16)(packedToken & result.OffsetMask);

            // SET temp2 TO 16 MINUS BitCount.
            var temp2 = (UInt16)(16 - result.BitCount);

            // SET Offset TO (temp1 RIGHT SHIFT BY temp2) PLUS 1.
            unpackedOffset = (UInt16)((temp1 >> temp2) + 1);
        }

        /// <summary>
        /// CopyToken Help derived bit masks are used by the Unpack CopyToken (section 2.4.1.3.19.2) 
        /// and the Pack CopyToken (section 2.4.1.3.19.3) algorithms. CopyToken Help also derives the 
        /// maximum length for a CopySequence (section 2.4.1.3.19) which is used by the Matching 
        /// algorithm (section 2.4.1.3.19.4).
        /// The pseudocode uses the state variables described in State Variables (section 2.4.1.2): 
        /// DecompressedCurrent and DecompressedChunkStart.
        /// </summary>
        internal static CopyTokenHelpResult CopyTokenHelp(long difference)
        {
            var result = new CopyTokenHelpResult();

            // SET BitCount TO the smallest integer that is GREATER THAN OR EQUAL TO LOGARITHM base 2 
            // of difference
            result.BitCount = 0;
            while ((1 << result.BitCount) < difference)
            {
                result.BitCount += 1;
            }

            // The number of bits used to encode Length MUST be greater than or equal to four. The 
            // number of bits used to encode Length MUST be less than or equal to 12
            // SET BitCount TO the maximum of BitCount and 4
            if (result.BitCount < 4)
                result.BitCount = 4;
            if (result.BitCount > 12)
                throw new Exception();

            // SET LengthMask TO 0xFFFF RIGHT SHIFT BY BitCount
            result.LengthMask = (UInt16)(0xffff >> result.BitCount);

            // SET OffsetMask TO BITWISE NOT LengthMask
            result.OffsetMask = (UInt16)(~result.LengthMask);

            // SET MaximumLength TO (0xFFFF RIGHT SHIFT BY BitCount) PLUS 3
            result.MaximumLength = (UInt16)((0xffff >> result.BitCount) + 3);

            return result;
        }

        public byte[] SerializeData()
        {
            var packedData = Pack(Position, _tokenOffset, _tokenLength);
            return BitConverter.GetBytes(packedData);
        }

        #region Nested Classes

        internal struct CopyTokenHelpResult
        {
            internal UInt16 LengthMask { get; set; }
            internal UInt16 OffsetMask { get; set; }
            internal UInt16 BitCount { get; set; }  // offset bit count.
            internal UInt16 MaximumLength { get; set; }
            internal UInt16 LengthBitCount => (UInt16)(16 - BitCount);
        }

        #endregion

        #region IEquatable
        public static bool operator !=(CopyToken first, CopyToken second)
        {
            return !(first == second);
        }

        public static bool operator ==(CopyToken first, CopyToken second)
        {
            return Equals(first, second);
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as CopyToken);
        }

        public bool Equals(IToken other)
        {
            return Equals(other as CopyToken);
        }

        public bool Equals(CopyToken other)
        {
            if (ReferenceEquals(other, null))
            {
                return false;
            }
            return other.Position == Position
                   && other.Length == Length
                   && other.Offset == Offset;
        }

        public override int GetHashCode()
        {
            return Position.GetHashCode() ^ Length.GetHashCode() ^ Offset.GetHashCode();
        }
        #endregion
    }

    /// <summary>
    /// The DecompressedBuffer is a resizable array of bytes that contains the same data as the 
    /// CompressedContainer (section 2.4.1.1.1), but the data is in an uncompressed format.
    /// </summary>
    /// <remarks></remarks>
    internal class DecompressedBuffer
    {
        internal DecompressedBuffer(byte[] uncompressedData)
        {
            using (var reader = new BinaryReader(new MemoryStream(uncompressedData)))
            {
                while (reader.BaseStream.Position < reader.BaseStream.Length)
                {
                    var chunk = new DecompressedChunk(reader);
                    DecompressedChunks.Add(chunk);
                }
            }
        }

        internal DecompressedBuffer(CompressedContainer container)
        {
            foreach (var chunk in container.CompressedChunks)
            {
                DecompressedChunks.Add(new DecompressedChunk(chunk));
            }
        }

        internal List<DecompressedChunk> DecompressedChunks { get; } = new List<DecompressedChunk>();

        internal byte[] Data
        {
            get
            {
                using (var writer = new BinaryWriter(new MemoryStream()))
                {
                    foreach (var chunk in DecompressedChunks)
                    {
                        writer.Write(chunk.Data);
                    }

                    using (var reader = new BinaryReader(writer.BaseStream))
                    {
                        reader.BaseStream.Position = 0;

                        return reader.ReadBytes((int) reader.BaseStream.Length);
                    }
                }
            }
        }
    }

    /// <summary>
    /// A DecompressedChunk is a resizable array of bytes in the DecompressedBuffer 
    /// (section 2.4.1.1.2). The byte array is the data from a CompressedChunk (section 2.4.1.1.4) in 
    /// uncompressed format.
    /// </summary>
    /// <remarks></remarks>
    internal class DecompressedChunk
    {
        internal DecompressedChunk(CompressedChunk compressedChunk)
        {
            if (compressedChunk.Header.IsCompressed)
            {
                // Loop through all the data, get TokenSequences and decompress them.
                using (var writer = new BinaryWriter(new MemoryStream()))
                {
                    var tokens = ((CompressedChunkData)compressedChunk.ChunkData).TokenSequences;
                    foreach (var sequence in tokens)
                    {
                        sequence.Tokens.DecompressTokenSequence(writer);
                    }

                    var stream = (MemoryStream)writer.BaseStream;
                    var decompressedData = stream.GetBuffer();
                    Array.Resize(ref decompressedData, (int)stream.Length);

                    Data = decompressedData;
                }
            }
            else
            {
                Data = compressedChunk.ChunkData.SerializeData();
            }
        }

        internal DecompressedChunk(BinaryReader reader)
        {
            var bytesToRead = reader.BaseStream.Length - reader.BaseStream.Position;

            if (bytesToRead > Globals.MaxBytesPerChunk)
                bytesToRead = Globals.MaxBytesPerChunk;

            Data = reader.ReadBytes((int) bytesToRead);
        }

        internal byte[] Data { get; }
    }

    internal static class Extensions
    {
        [DebuggerStepThrough]
        internal static byte[] ToMcbsBytes(this string textToConvert, UInt16 codePage)
        {
            return Encoding.GetEncoding(codePage).GetBytes(textToConvert);
        }

        // http://stackoverflow.com/questions/321370/convert-hex-string-to-byte-array
        internal static byte[] StringToByteArray(string hex)
        {
            return Enumerable.Range(0, hex.Length)
                             .Where(x => x % 2 == 0)
                             .Select(x => Convert.ToByte(hex.Substring(x, 2), 16))
                             .ToArray();
        }
    }

    internal static class Globals
    {
        internal const int MaxBytesPerChunk = 4096;
        internal const int NumberOfChunkHeaderBytes = 2;
        internal const byte PaddingByte = 0x0;
    }

    internal interface IChunkData
    {
        byte[] SerializeData();

        int Size { get; }
    }

    internal interface IToken : IEquatable<IToken>
    {
        void DecompressToken(BinaryWriter writer);

        byte[] SerializeData();

        long Length { get; }
    }

    /// <summary>
    /// A LiteralToken is a copy of one byte, in uncompressed format, from the DecompressedBuffer 
    /// (section 2.4.1.1.2).
    /// </summary>
    /// <remarks></remarks>
    internal class LiteralToken : IToken, IEquatable<LiteralToken>
    {
        private readonly byte[] _data;

        internal LiteralToken(BinaryReader dataReader)
        {
            _data = dataReader.ReadBytes(1);
        }

        internal LiteralToken(byte data)
        {
            _data = new [] { data };
        }

        public void DecompressToken(BinaryWriter writer)
        {
            writer.Write(_data);
            writer.Flush();
        }

        public byte[] SerializeData()
        {
            return _data;
        }

        public long Length => 1L;

        #region IEquatable
        public static bool operator !=(LiteralToken first, LiteralToken second)
        {
            return !(first == second);
        }

        public static bool operator ==(LiteralToken first, LiteralToken second)
        {
            return Equals(first, second);
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as LiteralToken);
        }

        public bool Equals(IToken other)
        {
            return Equals(other as LiteralToken);
        }

        public bool Equals(LiteralToken other)
        {
            if (ReferenceEquals(other, null))
            {
                return false;
            }
            return other._data.SequenceEqual(_data);
        }

        public override int GetHashCode()
        {
            return _data.GetHashCode();
        }
        #endregion
    }

    internal class RawChunk : IChunkData
    {
        private readonly byte[] _data;

        public RawChunk(byte[] data)
        {
            _data = data;
        }

        public byte[] SerializeData()
        {
            return _data;
        }

        public int Size => _data.Length;
    }

    internal static class Tokenizer
    {
        internal static IEnumerable<TokenSequence> ToTokenSequences(this IEnumerable<IToken> tokens)
        {
            var accumulatedTokens = new List<IToken>();
            foreach (var t in tokens)
            {
                if (accumulatedTokens.Count == 8)
                {
                    yield return new TokenSequence(accumulatedTokens);
                    accumulatedTokens.Clear();
                }
                accumulatedTokens.Add(t);
            }
            if (accumulatedTokens.Count != 0)
            {
                yield return new TokenSequence(accumulatedTokens);
            }
        }

        internal static void DecompressTokenSequence(this IEnumerable<IToken> tokens, BinaryWriter writer)
        {
            foreach (var token in tokens)
            {
                token.DecompressToken(writer);
            }
        }

        internal static bool OverlapsWith(this CopyToken thisToken, CopyToken otherToken)
        {
            var firstToken = thisToken;
            var secondToken = otherToken;
            if (thisToken.Position > otherToken.Position)
            {
                firstToken = otherToken;
                secondToken = thisToken;
            }
//            Contract.Assert(firstToken.Position <= secondToken.Position);

            return firstToken.Position + firstToken.Length > secondToken.Position;
        }

        internal static bool Contains(this CopyToken thisToken, CopyToken otherToken)
        {
            var otherTokenStartsAfterThisToken = thisToken.Position <= otherToken.Position;
            var otherTokenEndsBeforeThisToken = thisToken.Position + thisToken.Length >=
                                                otherToken.Position + otherToken.Length;
            return otherTokenStartsAfterThisToken && otherTokenEndsBeforeThisToken;
        }

        internal static IEnumerable<IToken> TokenizeUncompressedData(byte[] uncompressedData)
        {
            // The commented code is alternative to the specification for the compression.
            //var possibleCopyTokens = AllPossibleCopyTokens(uncompressedData);
            //var normalCopyTokens = NormalizeCopyTokens(possibleCopyTokens);
            //var allTokens = WeaveTokens(normalCopyTokens, uncompressedData);

            var copyTokens = GetSpecificationCopyTokens(uncompressedData);
            var allTokens = WeaveTokens(copyTokens, uncompressedData);
            foreach (var t in allTokens)
            {
                yield return t;
            }
        }

        private static IEnumerable<CopyToken> GetSpecificationCopyTokens(byte[] uncompressedData)
        {
            var position = 0L;
            while (position < uncompressedData.Length)
            {
                UInt16 offset = 0;
                UInt16 length = 0;
                Match(uncompressedData, position, out offset, out length);

                if (length > 0)
                {
                    yield return new CopyToken(position, offset, length);
                    position += length;
                }
                else
                {
                    position++;
                }
            }
        }

        private static IEnumerable<CopyToken> AllPossibleCopyTokens(byte[] uncompressedData)
        {
            var position = 0L;
            while (position < uncompressedData.Length)
            {
                UInt16 offset = 0;
                UInt16 length = 0;
                Match(uncompressedData, position, out offset, out length);
                
                if (length > 0)
                {
                    yield return new CopyToken(position, offset, length);
                }
                position++;
            }
        }

        private static IEnumerable<CopyToken> NormalizeCopyTokens(IEnumerable<CopyToken> copyTokens)
        {
            var remainingTokens = RemoveRedundantTokens(copyTokens).ToList();

            remainingTokens = RemoveOverlappingTokens(remainingTokens).ToList();

            return remainingTokens;
        }

        private static IEnumerable<CopyToken> RemoveRedundantTokens(IEnumerable<CopyToken> tokens)
        {
            CopyToken previous = null;
            foreach (var next in tokens)
            {
                if (previous == null)
                {
                    previous = next;
                    continue;
                }
                if (previous.OverlapsWith(next))
                {
                    //figure out which one to keep.  There can only be one!
                    if (previous.Length >= next.Length)
                    {
                        yield return previous;
                        // can't return next.
                    }
                    else
                    {
                        yield return next;
                    }
                }
                else
                {
                    yield return previous;
                    previous = next;
                }
            }
        }

        private static IEnumerable<CopyToken> RemoveOverlappingTokens(IEnumerable<CopyToken> tokens)
        {
            // create a list of the current tokens.
            Node list = null;
            foreach (var t in tokens.Reverse())
            {
                list = new Node(t, list);
            }
//            Contract.Assert(list != null);

            return FindBestPath(list);
        }

        private static Node FindBestPath(Node node)
        {
//            Contract.Requires<ArgumentNullException>(node != null);

            // find any overlapping tokens
            Node bestPath = null;
            foreach (var overlappingNode in GetOverlappingNodes(node))
            {
                var currentPath = new Node(overlappingNode.Value, null);

                // find the next non-overlapping node.
                var nonOverlappingNode = GetNextNonOverlappingNode(overlappingNode);
                if (nonOverlappingNode != null)
                {
                    currentPath.NextNode = FindBestPath(nonOverlappingNode);
                }

                if (bestPath == null 
                    || bestPath.Length < currentPath.Length)
                {
                    bestPath = currentPath;
                }
            }
            return bestPath;
        }

        private static IEnumerable<Node> GetOverlappingNodes(Node node)
        {
//            Contract.Requires<ArgumentNullException>(node != null);

            var firstNode = node;

            while (node != null 
                && firstNode.Value.OverlapsWith(node.Value))
            {
                yield return node;
                node = node.NextNode;
            }
        }

        private static Node GetNextNonOverlappingNode(Node node)
        {
//            Contract.Requires<ArgumentNullException>(node != null);

            var firstNode = node;

            while (node != null 
                && firstNode.Value.OverlapsWith(node.Value))
            {
                node = node.NextNode;
            }
            return node;
        }

        private static IEnumerable<IToken> WeaveTokens(IEnumerable<CopyToken> copyTokens, byte[] uncompressedData)
        {
            var position = 0L;
            foreach (var currentCopyToken in copyTokens)
            {
                while (position < currentCopyToken.Position)
                {
                    yield return new LiteralToken(uncompressedData[position]);
                    position++;
                }
                yield return currentCopyToken;
                position += currentCopyToken.Length;
            }
            while (position < uncompressedData.Length)
            {
                yield return new LiteralToken(uncompressedData[position]);
                position++;
            }
        }

        internal static void Match(byte[] uncompressedData, long position, out UInt16 matchedOffset, out UInt16 matchedLength)
        {
            var decompressedCurrent = position;
            var decompressedEnd = uncompressedData.Length;
            const long decompressedChunkStart = 0;

            // SET Candidate TO DecompressedCurrent MINUS 1
            var candidate = decompressedCurrent - 1L;
            // SET BestLength TO 0
            var bestLength = 0L;
            var bestCandidate = 0L;

            // WHILE Candidate is GREATER THAN OR EQUAL TO DecompressedChunkStart
            while (candidate >= decompressedChunkStart)
            {
                // SET C TO Candidate
                var c = candidate;
                // SET D TO DecompressedCurrent
                var d = decompressedCurrent;
                // SET Len TO 0
                var len = 0;

                // WHILE (D is LESS THAN DecompressedEnd)
                // and (the byte at D EQUALS the byte at C)
                while (d < decompressedEnd
                       && uncompressedData[d] == uncompressedData[c])
                {
                    // INCREMENT Len
                    len++;
                    // INCREMENT C
                    c++;
                    // INCREMENT D
                    d++;
                } // END WHILE

                // IF Len is GREATER THAN BestLength THEN
                if (len > bestLength)
                {
                    // SET BestLength TO Len
                    bestLength = len;
                    // SET BestCandidate TO Candidate
                    bestCandidate = candidate;
                } // ENDIF

                // DECREMENT Candidate
                candidate--;
            } // END WHILE

            // IF BestLength is GREATER THAN OR EQUAL TO 3 THEN
            if (bestLength >= 3)
            {
                // CALL CopyToken Help (section 2.4.1.3.19.1) returning MaximumLength
                var result = CopyToken.CopyTokenHelp(decompressedCurrent);

                // SET Length TO the MINIMUM of BestLength and MaximumLength
                matchedLength = (UInt16)bestLength;
                if (bestLength > result.MaximumLength)
                    matchedLength = result.MaximumLength;

                // SET Offset TO DecompressedCurrent MINUS BestCandidate
                matchedOffset = (UInt16)(decompressedCurrent - bestCandidate);
            }
            else // ELSE
            {
                // SET Length TO 0
                matchedLength = 0;
                // SET Offset TO 0
                matchedOffset = 0;
            } // ENDIF
        }

        //region Private Classes

        private class Node : IEnumerable<CopyToken>
        {
            public Node(CopyToken value, Node nextNode)
            {
//                Contract.Requires<ArgumentNullException>(value != null);

                Value = value;
                NextNode = nextNode;
            }

            internal CopyToken Value { get; }

            internal Node NextNode { get; set; }

            internal long Length
            {
                get
                {
                    if (NextNode != null)
                    {
                        return Value.Length + NextNode.Length;
                    }
                    return Value.Length;
                }
            }

            public IEnumerator<CopyToken> GetEnumerator()
            {
                return new NodeEnumerator(this);
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }

        private class NodeEnumerator : IEnumerator<CopyToken>
        {
            private Node _currentNode;
            private Node _nextNode;

            public NodeEnumerator(Node node)
            {
                _nextNode = node;
            }

            public void Dispose()
            {
            }

            public bool MoveNext()
            {
                if (_nextNode == null)
                {
                    return false;
                }
                _currentNode = _nextNode;
                _nextNode = _nextNode.NextNode;
                return true;
            }

            public void Reset()
            {
                throw new NotSupportedException();
            }

            public CopyToken Current => _currentNode.Value;

            object IEnumerator.Current => Current;
        }

        //endregion
    }

    /// <summary>
    /// A TokenSequence is a FlagByte followed by an array of Tokens. The number of Tokens in the final 
    /// TokenSequence MUST be greater than or equal to 1. The number of Tokens in the final 
    /// TokenSequence MUST less than or equal to eight. All other TokenSequences in the 
    /// CompressedChunkData MUST contain eight Tokens.
    /// </summary>
    /// <remarks></remarks>
    internal class TokenSequence
    {
        private byte _flagByte;
        private readonly List<IToken> _tokens = new List<IToken>();

        public TokenSequence(IEnumerable<IToken> enumerable) : this()
        {
            _tokens.AddRange(enumerable);
            
//            Contract.Assert(_tokens.Count > 0);
//            Contract.Assert(_tokens.Count <= 8);
            
            // set the flag byte.
            for (var i = 0; i < _tokens.Count; i++)
            {
                if (_tokens[i] is CopyToken)
                {
                    SetIsCopyToken(i, true);
                }
            }
        }

        private TokenSequence()
        { }

        internal long Length => Tokens.Sum(t => t.Length);

        internal IReadOnlyList<IToken> Tokens => _tokens;

        internal static TokenSequence GetFromCompressedData(BinaryReader reader, long position)
        {
            var sequence = new TokenSequence
            {
                _flagByte = reader.ReadByte()
            };

            for (var i = 0; i <= 7; i++)
            {
                if (sequence.GetIsCopyToken(i))
                {
                    var token = new CopyToken(reader, position);
                    sequence._tokens.Add(token);
                    position += Convert.ToInt64(token.Length);
                }
                else
                {
                    sequence._tokens.Add(new LiteralToken(reader));
                    position += 1;
                }
            }
            return sequence;
        }

        private void SetIsCopyToken(int index, bool value)
        {
            var setByte = (byte)Math.Pow(2, index);
            _flagByte = (byte)(_flagByte | setByte);
        }

        private bool GetIsCopyToken(int index)
        {
            var compareByte = (byte)Math.Pow(2, index);
            return (compareByte & _flagByte) != 0x0;
        }

        internal byte[] SerializeData()
        {
            var data = Enumerable.Repeat(_flagByte, 1);
            foreach (var token in Tokens)
            {
                data = data.Concat(token.SerializeData());
            }
            return data.ToArray();
        }
    }

    public static class VbaCompression
    {
        public static byte[] Compress(byte[] data)
        {
            var buffer = new DecompressedBuffer(data);
            var container = new CompressedContainer(buffer);
            return container.SerializeData();
        }

        public static byte[] Decompress(byte[] data)
        {
            var container = new CompressedContainer(data);
            var buffer = new DecompressedBuffer(container);
            return buffer.Data;
        }
    }
}