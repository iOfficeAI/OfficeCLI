// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Buffers.Binary;
using System.Text;

namespace OfficeCli.Core;

/// <summary>
/// Self-contained reader/writer for the Microsoft Compound File Binary
/// (CFB / OLE structured storage, the <c>D0 CF 11 E0</c> container) format,
/// scoped to exactly what OLE embedding needs: a single named stream inside
/// the root storage.
///
/// This replaces the former third-party OpenMcdf dependency so the OLE
/// wrap/unwrap path is fully owned in-tree. The implementation follows
/// [MS-CFB]:
///
/// <list type="bullet">
///   <item>Writer emits a V3 container (512-byte sectors). Streams smaller
///   than the 4096-byte mini-stream cutoff are stored in the mini stream
///   via the mini FAT; larger streams go in regular FAT sectors. Both
///   paths are implemented because embedded payloads vary in size and a
///   spec-compliant reader (real Office, the former OpenMcdf) decides which
///   region to read from purely off the recorded stream size.</item>
///   <item>Reader parses V3 <i>and</i> V4 containers (sector size taken from
///   the header sector shift) defensively, since the bytes it unwraps may
///   originate from any tool. Every offset is bounds-checked and chain
///   traversal is iteration-capped; malformed input yields <c>null</c>
///   rather than throwing, so callers fall back to the raw bytes.</item>
/// </list>
/// </summary>
internal static class CompoundFile
{
    private const uint FREESECT = 0xFFFFFFFF;
    private const uint ENDOFCHAIN = 0xFFFFFFFE;
    private const uint FATSECT = 0xFFFFFFFD;
    private const uint NOSTREAM = 0xFFFFFFFF;

    private const int MiniSectorSize = 64;
    private const int MiniStreamCutoff = 4096;
    private const int DirEntrySize = 128;
    private const int HeaderDifatCount = 109; // FAT sector slots stored in the header

    private static readonly byte[] Magic =
        { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

    // ==================== Writer ====================

    /// <summary>
    /// Build a minimal V3 CFB byte array whose root storage contains a single
    /// stream named <paramref name="streamName"/> holding <paramref name="data"/>.
    /// </summary>
    public static byte[] WriteSingleStream(string streamName, byte[] data)
    {
        if (streamName == null) throw new ArgumentNullException(nameof(streamName));
        data ??= Array.Empty<byte>();

        const int sectorSize = 512;
        var sectors = new List<byte[]>();   // regular sector payloads, id == index
        var fat = new List<uint>();         // FAT entry per sector (parallel to sectors)

        // Append a fresh chain of regular sectors holding blob; return start id.
        int WriteChain(byte[] blob)
        {
            int n = Math.Max(1, (blob.Length + sectorSize - 1) / sectorSize);
            int start = sectors.Count;
            for (int i = 0; i < n; i++)
            {
                var s = new byte[sectorSize];
                int off = i * sectorSize;
                int len = Math.Min(sectorSize, blob.Length - off);
                if (len > 0) Array.Copy(blob, off, s, 0, len);
                sectors.Add(s);
                fat.Add(0);
            }
            for (int i = 0; i < n; i++)
                fat[start + i] = i == n - 1 ? ENDOFCHAIN : (uint)(start + i + 1);
            return start;
        }

        uint rootStart = ENDOFCHAIN;
        long rootSize = 0;
        uint streamStart;
        uint firstMiniFat = ENDOFCHAIN;
        uint numMiniFat = 0;

        if (data.Length == 0)
        {
            streamStart = ENDOFCHAIN;
        }
        else if (data.Length < MiniStreamCutoff)
        {
            // Small stream → mini stream + mini FAT.
            int numMini = (data.Length + MiniSectorSize - 1) / MiniSectorSize;
            int miniBytes = numMini * MiniSectorSize;

            var miniStream = new byte[miniBytes];
            Array.Copy(data, miniStream, data.Length);
            rootStart = (uint)WriteChain(miniStream);
            rootSize = miniBytes;

            int miniFatSectors = (numMini + 127) / 128;
            var miniFatBlob = new byte[miniFatSectors * sectorSize];
            // Default every slot to FREESECT, then lay down the 0→1→…→END chain.
            for (int i = 0; i < miniFatBlob.Length; i += 4)
                BinaryPrimitives.WriteUInt32LittleEndian(miniFatBlob.AsSpan(i), FREESECT);
            for (int i = 0; i < numMini; i++)
                BinaryPrimitives.WriteUInt32LittleEndian(
                    miniFatBlob.AsSpan(i * 4),
                    i == numMini - 1 ? ENDOFCHAIN : (uint)(i + 1));
            firstMiniFat = (uint)WriteChain(miniFatBlob);
            numMiniFat = (uint)miniFatSectors;

            streamStart = 0; // first mini-sector index
        }
        else
        {
            // Large stream → regular FAT sectors.
            streamStart = (uint)WriteChain(data);
        }

        // Directory: Root Entry + the stream entry, padded to a 512-byte sector.
        var dir = new byte[4 * DirEntrySize];
        WriteDirEntry(dir, 0 * DirEntrySize, "Root Entry", objType: 5,
            left: NOSTREAM, right: NOSTREAM, child: 1, start: rootStart, size: rootSize);
        WriteDirEntry(dir, 1 * DirEntrySize, streamName, objType: 2,
            left: NOSTREAM, right: NOSTREAM, child: NOSTREAM,
            start: streamStart, size: data.Length);
        // Entries 2 and 3 stay zeroed (objType 0 = unallocated).
        int dirStart = WriteChain(dir);

        // Allocate FAT sectors last: they must also be represented in the FAT.
        int dataSectors = sectors.Count;
        int numFat = Math.Max(1, (dataSectors + 126) / 127); // ceil(dataSectors/127)
        // Adding numFat sectors may push us over another FAT-sector boundary.
        while (dataSectors + numFat > numFat * 128) numFat++;
        if (numFat > HeaderDifatCount)
            throw new NotSupportedException(
                "OLE payload too large to wrap (would need DIFAT sectors).");

        var fatSectorIds = new int[numFat];
        for (int i = 0; i < numFat; i++)
        {
            fatSectorIds[i] = sectors.Count;
            sectors.Add(new byte[sectorSize]);
            fat.Add(FATSECT);
        }

        // Serialize the FAT array across the FAT sectors (tail = FREESECT).
        var fatBytes = new byte[numFat * sectorSize];
        for (int i = 0; i < fatBytes.Length; i += 4)
            BinaryPrimitives.WriteUInt32LittleEndian(fatBytes.AsSpan(i), FREESECT);
        for (int s = 0; s < fat.Count; s++)
            BinaryPrimitives.WriteUInt32LittleEndian(fatBytes.AsSpan(s * 4), fat[s]);
        for (int i = 0; i < numFat; i++)
            Array.Copy(fatBytes, i * sectorSize, sectors[fatSectorIds[i]], 0, sectorSize);

        // Header (512 bytes) + every sector in id order.
        var header = new byte[sectorSize];
        Array.Copy(Magic, header, Magic.Length);
        // CLSID (8..23) stays zero.
        BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(24), 0x003E); // minor version
        BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(26), 0x0003); // major version (V3)
        BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(28), 0xFFFE); // byte order
        BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(30), 0x0009); // sector shift → 512
        BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(32), 0x0006); // mini sector shift → 64
        // reserved (34..39) zero, numDirSectors (40) = 0 for V3.
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(44), (uint)numFat);
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(48), (uint)dirStart);
        // transaction signature (52) zero.
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(56), MiniStreamCutoff);
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(60), firstMiniFat);
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(64), numMiniFat);
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(68), ENDOFCHAIN); // first DIFAT sector
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(72), 0);          // num DIFAT sectors
        // DIFAT array at 76: FAT sector ids, rest FREESECT.
        for (int i = 0; i < HeaderDifatCount; i++)
            BinaryPrimitives.WriteUInt32LittleEndian(
                header.AsSpan(76 + i * 4),
                i < numFat ? (uint)fatSectorIds[i] : FREESECT);

        var output = new byte[sectorSize + sectors.Count * sectorSize];
        Array.Copy(header, 0, output, 0, sectorSize);
        for (int i = 0; i < sectors.Count; i++)
            Array.Copy(sectors[i], 0, output, sectorSize + i * sectorSize, sectorSize);
        return output;
    }

    private static void WriteDirEntry(byte[] buf, int offset, string name,
        byte objType, uint left, uint right, uint child, uint start, long size)
    {
        var nameBytes = Encoding.Unicode.GetBytes(name);
        int copyLen = Math.Min(nameBytes.Length, 62); // 31 UTF-16 chars + null = 64 bytes
        Array.Copy(nameBytes, 0, buf, offset, copyLen);
        BinaryPrimitives.WriteUInt16LittleEndian(buf.AsSpan(offset + 64), (ushort)(copyLen + 2));
        buf[offset + 66] = objType;
        buf[offset + 67] = 1; // colorFlag: black
        BinaryPrimitives.WriteUInt32LittleEndian(buf.AsSpan(offset + 68), left);
        BinaryPrimitives.WriteUInt32LittleEndian(buf.AsSpan(offset + 72), right);
        BinaryPrimitives.WriteUInt32LittleEndian(buf.AsSpan(offset + 76), child);
        // CLSID (80..95), stateBits (96..99), timestamps (100..115) stay zero.
        BinaryPrimitives.WriteUInt32LittleEndian(buf.AsSpan(offset + 116), start);
        BinaryPrimitives.WriteUInt64LittleEndian(buf.AsSpan(offset + 120), (ulong)size);
    }

    // ==================== Multi-stream / nested writer ====================

    private sealed class DirNode
    {
        public string Name = "";
        public bool IsStorage;
        public byte[] Data = Array.Empty<byte>();
        public readonly List<DirNode> Children = new();
        public int Index = -1;
        // Resolved directory pointers (entry indices, or NOSTREAM).
        public uint Left = NOSTREAM, Right = NOSTREAM, Child = NOSTREAM;
        // Stream placement.
        public uint Start = ENDOFCHAIN;
        public long Size;
    }

    /// <summary>
    /// Build a V3 CFB byte array containing an arbitrary set of named streams,
    /// optionally nested inside storages. Each entry's <c>path</c> is
    /// '/'-separated: all but the last segment are storages, the last is the
    /// stream (segment names may include the control-char prefixes Office uses,
    /// e.g. <c>"DataSpaces/Version"</c>).
    ///
    /// Used to assemble password-encrypted OOXML (EncryptionInfo +
    /// EncryptedPackage + the DataSpaces transform tree). Small streams
    /// (&lt; 4096 bytes) are consolidated into the root mini stream via the mini
    /// FAT; larger ones use regular FAT sectors — exactly like a real Office
    /// container, so spec-compliant readers (Excel, LibreOffice) accept it.
    /// Siblings within each storage form a balanced BST ordered by the
    /// [MS-CFB] §2.6.4 name comparison, so directory lookups resolve correctly.
    /// </summary>
    public static byte[] WriteStreams(IEnumerable<(string path, byte[] data)> entries)
    {
        const int sectorSize = 512;

        // ── 1. Build the storage/stream tree from the flat paths. ──────────
        var root = new DirNode { Name = "Root Entry", IsStorage = true };
        foreach (var (path, data) in entries)
        {
            var segs = path.Split('/');
            var cur = root;
            for (int i = 0; i < segs.Length; i++)
            {
                bool leaf = i == segs.Length - 1;
                var existing = cur.Children.Find(c => c.Name == segs[i]);
                if (existing == null)
                {
                    existing = new DirNode { Name = segs[i], IsStorage = !leaf };
                    if (leaf) existing.Data = data ?? Array.Empty<byte>();
                    cur.Children.Add(existing);
                }
                cur = existing;
            }
        }

        // ── 2. Assign directory indices (Root == 0, then pre-order). ───────
        var nodes = new List<DirNode>();
        void Index(DirNode n) { n.Index = nodes.Count; nodes.Add(n); foreach (var c in n.Children) Index(c); }
        Index(root);

        // ── 3. Per storage, order children + build a balanced sibling BST. ──
        foreach (var n in nodes)
        {
            if (!n.IsStorage || n.Children.Count == 0) continue;
            var kids = new List<DirNode>(n.Children);
            kids.Sort((a, b) => CfbNameCompare(a.Name, b.Name));
            n.Child = (uint)BuildBst(kids, 0, kids.Count - 1).Index;
        }

        // ── 4. Lay out streams: small → mini stream, large → regular FAT. ──
        var sectors = new List<byte[]>();
        var fat = new List<uint>();
        int WriteChain(byte[] blob)
        {
            int count = Math.Max(1, (blob.Length + sectorSize - 1) / sectorSize);
            int start = sectors.Count;
            for (int i = 0; i < count; i++)
            {
                var s = new byte[sectorSize];
                int off = i * sectorSize, len = Math.Min(sectorSize, blob.Length - off);
                if (len > 0) Array.Copy(blob, off, s, 0, len);
                sectors.Add(s);
                fat.Add(i == count - 1 ? ENDOFCHAIN : 0);
            }
            for (int i = 0; i < count - 1; i++) fat[start + i] = (uint)(start + i + 1);
            return start;
        }

        var miniStream = new List<byte>();
        var miniChainNext = new List<uint>();   // mini FAT (per 64-byte mini-sector)
        foreach (var n in nodes)
        {
            if (n.IsStorage) continue;
            n.Size = n.Data.Length;
            if (n.Data.Length == 0) { n.Start = ENDOFCHAIN; continue; }
            if (n.Data.Length < MiniStreamCutoff)
            {
                int firstMini = miniStream.Count / MiniSectorSize;
                int need = (n.Data.Length + MiniSectorSize - 1) / MiniSectorSize;
                miniStream.AddRange(n.Data);
                int pad = need * MiniSectorSize - n.Data.Length;
                for (int i = 0; i < pad; i++) miniStream.Add(0);
                for (int i = 0; i < need; i++)
                    miniChainNext.Add(i == need - 1 ? ENDOFCHAIN : (uint)(firstMini + i + 1));
                n.Start = (uint)firstMini;
            }
            else
            {
                n.Start = (uint)WriteChain(n.Data);
            }
        }

        // Mini stream is itself a regular stream owned by Root Entry.
        uint rootStart = ENDOFCHAIN; long rootSize = 0;
        uint firstMiniFat = ENDOFCHAIN; uint numMiniFat = 0;
        if (miniStream.Count > 0)
        {
            rootStart = (uint)WriteChain(miniStream.ToArray());
            rootSize = miniStream.Count;

            int miniFatSectors = Math.Max(1, (miniChainNext.Count + 127) / 128);
            var miniFatBlob = new byte[miniFatSectors * sectorSize];
            for (int i = 0; i < miniFatBlob.Length; i += 4)
                BinaryPrimitives.WriteUInt32LittleEndian(miniFatBlob.AsSpan(i), FREESECT);
            for (int i = 0; i < miniChainNext.Count; i++)
                BinaryPrimitives.WriteUInt32LittleEndian(miniFatBlob.AsSpan(i * 4), miniChainNext[i]);
            firstMiniFat = (uint)WriteChain(miniFatBlob);
            numMiniFat = (uint)miniFatSectors;
        }
        root.Start = rootStart;
        root.Size = rootSize;

        // ── 5. Directory stream: one 128-byte entry per node, sector-padded. ─
        int dirEntries = nodes.Count;
        int dirSectors = Math.Max(1, (dirEntries + 3) / 4);
        var dir = new byte[dirSectors * sectorSize];
        foreach (var n in nodes)
        {
            byte objType = n == root ? (byte)5 : n.IsStorage ? (byte)1 : (byte)2;
            WriteDirEntry(dir, n.Index * DirEntrySize, n.Name, objType,
                n.Left, n.Right, n.Child, n.Start, n.Size);
        }
        int dirStart = WriteChain(dir);

        // ── 6. Allocate FAT sectors last (they must self-represent). ────────
        int dataSectors = sectors.Count;
        int numFat = Math.Max(1, (dataSectors + 126) / 127);
        while (dataSectors + numFat > numFat * 128) numFat++;
        if (numFat > HeaderDifatCount)
            throw new NotSupportedException("Encrypted payload too large to wrap (would need DIFAT sectors).");

        var fatSectorIds = new int[numFat];
        for (int i = 0; i < numFat; i++)
        {
            fatSectorIds[i] = sectors.Count;
            sectors.Add(new byte[sectorSize]);
            fat.Add(FATSECT);
        }

        var fatBytes = new byte[numFat * sectorSize];
        for (int i = 0; i < fatBytes.Length; i += 4)
            BinaryPrimitives.WriteUInt32LittleEndian(fatBytes.AsSpan(i), FREESECT);
        for (int s = 0; s < fat.Count; s++)
            BinaryPrimitives.WriteUInt32LittleEndian(fatBytes.AsSpan(s * 4), fat[s]);
        for (int i = 0; i < numFat; i++)
            Array.Copy(fatBytes, i * sectorSize, sectors[fatSectorIds[i]], 0, sectorSize);

        // ── 7. Header + sectors. ───────────────────────────────────────────
        var header = new byte[sectorSize];
        Array.Copy(Magic, header, Magic.Length);
        BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(24), 0x003E);
        BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(26), 0x0003); // V3
        BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(28), 0xFFFE);
        BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(30), 0x0009); // 512-byte sectors
        BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(32), 0x0006); // 64-byte mini sectors
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(44), (uint)numFat);
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(48), (uint)dirStart);
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(56), MiniStreamCutoff);
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(60), firstMiniFat);
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(64), numMiniFat);
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(68), ENDOFCHAIN);
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(72), 0);
        for (int i = 0; i < HeaderDifatCount; i++)
            BinaryPrimitives.WriteUInt32LittleEndian(
                header.AsSpan(76 + i * 4), i < numFat ? (uint)fatSectorIds[i] : FREESECT);

        var output = new byte[sectorSize + sectors.Count * sectorSize];
        Array.Copy(header, 0, output, 0, sectorSize);
        for (int i = 0; i < sectors.Count; i++)
            Array.Copy(sectors[i], 0, output, sectorSize + i * sectorSize, sectorSize);
        return output;
    }

    /// <summary>Build a height-balanced BST from name-sorted children, wiring
    /// each node's Left/Right to its subtree roots; returns the subtree root.</summary>
    private static DirNode BuildBst(List<DirNode> sorted, int lo, int hi)
    {
        int mid = (lo + hi) / 2;
        var node = sorted[mid];
        node.Left = lo <= mid - 1 ? (uint)BuildBst(sorted, lo, mid - 1).Index : NOSTREAM;
        node.Right = mid + 1 <= hi ? (uint)BuildBst(sorted, mid + 1, hi).Index : NOSTREAM;
        return node;
    }

    /// <summary>[MS-CFB] §2.6.4 directory name order: by UTF-16 length first,
    /// then by uppercased UTF-16 code units.</summary>
    private static int CfbNameCompare(string a, string b)
    {
        if (a.Length != b.Length) return a.Length - b.Length;
        for (int i = 0; i < a.Length; i++)
        {
            int ca = char.ToUpperInvariant(a[i]), cb = char.ToUpperInvariant(b[i]);
            if (ca != cb) return ca - cb;
        }
        return 0;
    }

    // ==================== Reader ====================

    /// <summary>
    /// Extract the stream named <paramref name="streamName"/> from a CFB byte
    /// array. Returns the stream bytes, or <c>null</c> if the input is not a
    /// valid CFB, the stream is absent, or any structural inconsistency is hit.
    /// </summary>
    public static byte[]? ReadStream(byte[] cfb, string streamName)
    {
        try
        {
            if (cfb == null || cfb.Length < 512) return null;
            for (int i = 0; i < Magic.Length; i++)
                if (cfb[i] != Magic[i]) return null;

            int sectorShift = BinaryPrimitives.ReadUInt16LittleEndian(cfb.AsSpan(30));
            if (sectorShift < 7 || sectorShift > 20) return null;
            int sectorSize = 1 << sectorShift;
            if (sectorSize < 128) return null;

            uint numFat = BinaryPrimitives.ReadUInt32LittleEndian(cfb.AsSpan(44));
            uint firstDir = BinaryPrimitives.ReadUInt32LittleEndian(cfb.AsSpan(48));
            uint miniCutoff = BinaryPrimitives.ReadUInt32LittleEndian(cfb.AsSpan(56));
            uint firstMiniFat = BinaryPrimitives.ReadUInt32LittleEndian(cfb.AsSpan(60));
            uint firstDifat = BinaryPrimitives.ReadUInt32LittleEndian(cfb.AsSpan(68));
            uint numDifat = BinaryPrimitives.ReadUInt32LittleEndian(cfb.AsSpan(72));

            int totalSectors = (cfb.Length - sectorSize) / sectorSize;
            if (totalSectors <= 0) return null;

            int SectorOffset(uint id) => sectorSize + (int)id * sectorSize;
            bool ValidSector(uint id) => id < (uint)totalSectors;

            // Gather FAT sector ids: header DIFAT (109) then any DIFAT sectors.
            var fatSectorIds = new List<uint>();
            for (int i = 0; i < HeaderDifatCount && fatSectorIds.Count < numFat; i++)
            {
                uint id = BinaryPrimitives.ReadUInt32LittleEndian(cfb.AsSpan(76 + i * 4));
                if (id == FREESECT) break;
                fatSectorIds.Add(id);
            }
            uint difatCur = firstDifat;
            int difatGuard = 0;
            int perDifat = sectorSize / 4 - 1;
            while (difatCur != ENDOFCHAIN && difatCur != FREESECT &&
                   fatSectorIds.Count < numFat && difatGuard++ <= totalSectors + 1)
            {
                if (!ValidSector(difatCur)) return null;
                int baseOff = SectorOffset(difatCur);
                for (int i = 0; i < perDifat && fatSectorIds.Count < numFat; i++)
                {
                    uint id = BinaryPrimitives.ReadUInt32LittleEndian(cfb.AsSpan(baseOff + i * 4));
                    if (id == FREESECT) continue;
                    fatSectorIds.Add(id);
                }
                difatCur = BinaryPrimitives.ReadUInt32LittleEndian(cfb.AsSpan(baseOff + perDifat * 4));
            }
            if (numDifat > 0) { /* followed above; count is advisory */ }

            // Build the full FAT array.
            int fatEntries = fatSectorIds.Count * (sectorSize / 4);
            var fat = new uint[fatEntries];
            int w = 0;
            foreach (uint fid in fatSectorIds)
            {
                if (!ValidSector(fid)) return null;
                int baseOff = SectorOffset(fid);
                for (int i = 0; i < sectorSize / 4; i++)
                    fat[w++] = BinaryPrimitives.ReadUInt32LittleEndian(cfb.AsSpan(baseOff + i * 4));
            }

            // Follow a FAT chain; cap iterations to guard against cycles.
            List<uint>? Chain(uint start)
            {
                var ids = new List<uint>();
                uint cur = start;
                int guard = 0;
                while (cur != ENDOFCHAIN && cur != FREESECT)
                {
                    if (cur >= (uint)fat.Length || !ValidSector(cur)) return null;
                    if (guard++ > totalSectors + 1) return null;
                    ids.Add(cur);
                    cur = fat[cur];
                }
                return ids;
            }

            byte[]? ReadRegular(uint start, long size)
            {
                var ids = Chain(start);
                if (ids == null) return null;
                var buf = new byte[ids.Count * sectorSize];
                for (int i = 0; i < ids.Count; i++)
                    Array.Copy(cfb, SectorOffset(ids[i]), buf, i * sectorSize, sectorSize);
                if (size < 0 || size > buf.Length) return buf;
                var outBuf = new byte[size];
                Array.Copy(buf, outBuf, (int)size);
                return outBuf;
            }

            // Parse the directory and locate Root Entry + the target stream.
            var dirChain = Chain(firstDir);
            if (dirChain == null) return null;
            var target = Encoding.Unicode.GetBytes(streamName);
            uint rootStart = ENDOFCHAIN;
            long rootSize = 0;
            bool foundStream = false;
            uint streamStart = ENDOFCHAIN;
            long streamSize = 0;

            foreach (uint dsec in dirChain)
            {
                int baseOff = SectorOffset(dsec);
                for (int e = 0; e + DirEntrySize <= sectorSize; e += DirEntrySize)
                {
                    int off = baseOff + e;
                    byte objType = cfb[off + 66];
                    if (objType == 5) // root
                    {
                        rootStart = BinaryPrimitives.ReadUInt32LittleEndian(cfb.AsSpan(off + 116));
                        rootSize = (long)BinaryPrimitives.ReadUInt64LittleEndian(cfb.AsSpan(off + 120));
                    }
                    else if (objType == 2 && !foundStream) // stream
                    {
                        int nameLen = BinaryPrimitives.ReadUInt16LittleEndian(cfb.AsSpan(off + 64));
                        nameLen = Math.Clamp(nameLen, 0, 64);
                        int cmpLen = nameLen >= 2 ? nameLen - 2 : 0; // drop null terminator
                        if (cmpLen == target.Length && cfb.AsSpan(off, cmpLen).SequenceEqual(target))
                        {
                            streamStart = BinaryPrimitives.ReadUInt32LittleEndian(cfb.AsSpan(off + 116));
                            streamSize = (long)BinaryPrimitives.ReadUInt64LittleEndian(cfb.AsSpan(off + 120));
                            foundStream = true;
                        }
                    }
                }
            }

            if (!foundStream) return null;
            if (streamSize == 0) return Array.Empty<byte>();

            // Large stream: read straight from the regular FAT.
            if (miniCutoff == 0 || streamSize >= miniCutoff)
                return ReadRegular(streamStart, streamSize);

            // Small stream: read from the mini stream via the mini FAT.
            byte[]? miniStream = ReadRegular(rootStart, rootSize);
            if (miniStream == null) return null;

            var miniFatChain = Chain(firstMiniFat);
            if (miniFatChain == null) return null;
            int miniFatEntries = miniFatChain.Count * (sectorSize / 4);
            var miniFat = new uint[miniFatEntries];
            int mw = 0;
            foreach (uint mfid in miniFatChain)
            {
                int baseOff = SectorOffset(mfid);
                for (int i = 0; i < sectorSize / 4; i++)
                    miniFat[mw++] = BinaryPrimitives.ReadUInt32LittleEndian(cfb.AsSpan(baseOff + i * 4));
            }

            var result = new byte[streamSize];
            int written = 0;
            uint miniCur = streamStart;
            int miniGuard = 0;
            int totalMini = miniStream.Length / MiniSectorSize;
            while (miniCur != ENDOFCHAIN && miniCur != FREESECT && written < streamSize)
            {
                if (miniCur >= (uint)miniFat.Length || miniCur >= (uint)totalMini) return null;
                if (miniGuard++ > totalMini + 1) return null;
                int take = (int)Math.Min(MiniSectorSize, streamSize - written);
                Array.Copy(miniStream, (int)miniCur * MiniSectorSize, result, written, take);
                written += take;
                miniCur = miniFat[miniCur];
            }
            return written == streamSize ? result : null;
        }
        catch
        {
            return null;
        }
    }
}
