// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Buffers.Binary;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace OfficeCli.Core;

/// <summary>
/// Reads and writes password-encrypted OOXML (the <c>D0 CF 11 E0</c> CDFV2
/// container Office produces when you set a document password). Implements the
/// <b>Agile</b> encryption scheme from <c>[MS-OFFCRYPTO]</c> §2.3.4.10–2.3.4.15:
/// AES-CBC over a per-segment-IV'd <c>EncryptedPackage</c>, with the package key
/// wrapped by a password-derived key (iterated hash) and an HMAC
/// <c>dataIntegrity</c> block.
///
/// <para>This is a clean-room implementation written from the published spec and
/// validated byte-for-byte against an independent reference decryptor — it does
/// not port code from any (L)GPL/MPL office suite, keeping the file Apache-2.0
/// clean like the rest of the tree.</para>
///
/// <para>Scope: <b>Agile</b> only (the modern default since Office 2013 —
/// SHA-512/AES-256). The legacy "Standard" (ECMA-376 binary) and pre-2007 RC4
/// schemes are intentionally out of scope; <see cref="IsEncryptedOoxml"/>
/// recognises them so callers can give a clear "unsupported scheme" message
/// rather than a corrupt-file error.</para>
///
/// <para>The CFB container itself is read/written by <see cref="CompoundFile"/>;
/// this class only owns the cryptography and the <c>EncryptionInfo</c> XML.</para>
/// </summary>
internal static class MsOffCrypto
{
    // CDFV2 / OLE compound-file magic — shared with CompoundFile.
    private static readonly byte[] CfbMagic =
        { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

    // Block keys that salt the key-derivation for each encrypted blob
    // ([MS-OFFCRYPTO] 2.3.4.12–2.3.4.14). Constant across all files.
    private static readonly byte[] BkVerifierInput = { 0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79 };
    private static readonly byte[] BkVerifierValue = { 0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e };
    private static readonly byte[] BkKeyValue      = { 0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6 };
    private static readonly byte[] BkHmacKey        = { 0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6 };
    private static readonly byte[] BkHmacValue       = { 0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33 };

    private const int PackageSegment = 4096; // EncryptedPackage cipher segment size
    private static readonly XNamespace NsEnc =
        "http://schemas.microsoft.com/office/2006/encryption";
    private static readonly XNamespace NsPwd =
        "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

    /// <summary>True if the first bytes are the OLE/CFB magic — i.e. this is an
    /// encrypted (or otherwise OLE-wrapped) OOXML file, not a plain zip.</summary>
    public static bool IsCfb(byte[] bytes)
    {
        if (bytes is null || bytes.Length < CfbMagic.Length) return false;
        for (int i = 0; i < CfbMagic.Length; i++)
            if (bytes[i] != CfbMagic[i]) return false;
        return true;
    }

    /// <summary>Why an encrypted file could not be opened, for a precise message.</summary>
    public enum Scheme { NotEncrypted, Agile, StandardUnsupported }

    /// <summary>
    /// Classify a CFB blob: is it Agile-encrypted (supported), Standard/legacy
    /// (recognised but unsupported), or not an encrypted OOXML at all?
    /// </summary>
    public static Scheme Classify(byte[] cfb)
    {
        if (!IsCfb(cfb)) return Scheme.NotEncrypted;
        byte[]? info = CompoundFile.ReadStream(cfb, "EncryptionInfo");
        if (info is null || info.Length < 8) return Scheme.NotEncrypted;
        ushort vMajor = BinaryPrimitives.ReadUInt16LittleEndian(info.AsSpan(0));
        ushort vMinor = BinaryPrimitives.ReadUInt16LittleEndian(info.AsSpan(2));
        // Agile: version 4.4 with the AES/extensible flag; the rest is XML.
        if (vMajor == 4 && vMinor == 4) return Scheme.Agile;
        // Standard (2.2/3.2/4.2) and legacy RC4 — recognised, not supported.
        return Scheme.StandardUnsupported;
    }

    /// <summary>Convenience: is this an OOXML we can actually decrypt?</summary>
    public static bool IsEncryptedOoxml(byte[] cfb) => Classify(cfb) == Scheme.Agile;

    public sealed class WrongPasswordException : Exception
    {
        public WrongPasswordException() : base("Incorrect password for encrypted document.") { }
    }
    public sealed class UnsupportedSchemeException : Exception
    {
        public UnsupportedSchemeException(string m) : base(m) { }
    }

    // ==================== Decrypt ====================

    /// <summary>
    /// Decrypt an Agile-encrypted OOXML <paramref name="cfb"/> with
    /// <paramref name="password"/>, returning the inner plain zip bytes
    /// (PK\x03\x04…). Throws <see cref="WrongPasswordException"/> if the password
    /// fails the verifier, or <see cref="UnsupportedSchemeException"/> for
    /// non-Agile encryption.
    /// </summary>
    public static byte[] Decrypt(byte[] cfb, string password)
    {
        switch (Classify(cfb))
        {
            case Scheme.NotEncrypted:
                throw new InvalidOperationException("File is not an encrypted OOXML document.");
            case Scheme.StandardUnsupported:
                throw new UnsupportedSchemeException(
                    "This file uses the legacy 'Standard' encryption scheme, which officecli " +
                    "does not support (only modern 'Agile' encryption). Open it in LibreOffice/Excel instead.");
        }

        byte[] info = CompoundFile.ReadStream(cfb, "EncryptionInfo")
            ?? throw new InvalidOperationException("EncryptionInfo stream missing.");
        byte[] package = CompoundFile.ReadStream(cfb, "EncryptedPackage")
            ?? throw new InvalidOperationException("EncryptedPackage stream missing.");

        var d = AgileDescriptor.Parse(info);
        byte[] hSpin = SpinHash(d.HashAlg, d.PwSalt, password, d.SpinCount);

        // Verify the password via the verifier hash before doing real work.
        byte[] verifierInput = AesCbcDecrypt(
            DeriveKey(d.HashAlg, hSpin, BkVerifierInput, d.KeyBytes), Fit(d.PwSalt, d.BlockSize), d.EncVerifierHashInput);
        byte[] verifierHash = AesCbcDecrypt(
            DeriveKey(d.HashAlg, hSpin, BkVerifierValue, d.KeyBytes), Fit(d.PwSalt, d.BlockSize), d.EncVerifierHashValue);
        byte[] expected = Hash(d.HashAlg, verifierInput.AsSpan(0, d.SaltSize).ToArray());
        if (!verifierHash.AsSpan(0, expected.Length).SequenceEqual(expected))
            throw new WrongPasswordException();

        // Unwrap the package secret key, then decrypt the package per-segment.
        byte[] secret = AesCbcDecrypt(
            DeriveKey(d.HashAlg, hSpin, BkKeyValue, d.KeyBytes), Fit(d.PwSalt, d.BlockSize), d.EncKeyValue);
        secret = secret.AsSpan(0, d.KeyDataKeyBytes).ToArray();

        long total = BinaryPrimitives.ReadInt64LittleEndian(package.AsSpan(0));
        byte[] enc = package.AsSpan(8).ToArray();
        var outBuf = new byte[checked((int)((enc.Length + d.BlockSize - 1) / d.BlockSize) * d.BlockSize)];
        int written = 0;
        for (int i = 0; i < enc.Length; i += PackageSegment)
        {
            int len = Math.Min(PackageSegment, enc.Length - i);
            byte[] iv = Fit(Hash(d.KeyDataHashAlg, d.KeyDataSalt, IntLe(i / PackageSegment)), d.BlockSize);
            byte[] dec = AesCbcDecrypt(secret, iv, enc.AsSpan(i, len).ToArray());
            Array.Copy(dec, 0, outBuf, written, dec.Length);
            written += dec.Length;
        }
        if (total < 0 || total > outBuf.Length) total = outBuf.Length;
        return outBuf.AsSpan(0, (int)total).ToArray();
    }

    // ==================== Encrypt ====================

    /// <summary>
    /// Encrypt plain OOXML zip <paramref name="plain"/> under
    /// <paramref name="password"/> using Agile encryption (SHA-512 / AES-256),
    /// returning a complete CDFV2 file (EncryptionInfo + EncryptedPackage + the
    /// DataSpaces transform streams) ready to write to disk.
    /// </summary>
    public static byte[] Encrypt(byte[] plain, string password)
    {
        const string hashAlg = "SHA512";
        const int blockSize = 16, saltSize = 16, keyBytes = 32, hashSize = 64, spin = 100000;

        byte[] keyDataSalt = RandomNumberGenerator.GetBytes(saltSize);
        byte[] pwSalt = RandomNumberGenerator.GetBytes(saltSize);
        byte[] secret = RandomNumberGenerator.GetBytes(keyBytes);
        byte[] verifierInput = RandomNumberGenerator.GetBytes(saltSize);

        byte[] hSpin = SpinHash(hashAlg, pwSalt, password, spin);
        byte[] encVerifierHashInput = AesCbcEncrypt(
            DeriveKey(hashAlg, hSpin, BkVerifierInput, keyBytes), Fit(pwSalt, blockSize), PadBlock(verifierInput, blockSize));
        byte[] verifierHash = Hash(hashAlg, verifierInput);
        byte[] encVerifierHashValue = AesCbcEncrypt(
            DeriveKey(hashAlg, hSpin, BkVerifierValue, keyBytes), Fit(pwSalt, blockSize), PadBlock(verifierHash, blockSize));
        byte[] encKeyValue = AesCbcEncrypt(
            DeriveKey(hashAlg, hSpin, BkKeyValue, keyBytes), Fit(pwSalt, blockSize), secret);

        // EncryptedPackage: 8-byte LE size prefix, then per-4096-segment AES-CBC.
        var enc = new byte[checked(((plain.Length + blockSize - 1) / blockSize) * blockSize)];
        int written = 0;
        for (int i = 0; i < plain.Length; i += PackageSegment)
        {
            int len = Math.Min(PackageSegment, plain.Length - i);
            byte[] iv = Fit(Hash(hashAlg, keyDataSalt, IntLe(i / PackageSegment)), blockSize);
            byte[] seg = AesCbcEncrypt(secret, iv, PadBlock(plain.AsSpan(i, len).ToArray(), blockSize));
            Array.Copy(seg, 0, enc, written, seg.Length);
            written += seg.Length;
        }
        var package = new byte[8 + written];
        BinaryPrimitives.WriteInt64LittleEndian(package.AsSpan(0), plain.Length);
        Array.Copy(enc, 0, package, 8, written);

        // dataIntegrity: HMAC of the EncryptedPackage, both key and value wrapped
        // under the package secret with block-key-salted IVs.
        byte[] hmacKey = RandomNumberGenerator.GetBytes(hashSize);
        byte[] encHmacKey = AesCbcEncrypt(secret,
            Fit(Hash(hashAlg, keyDataSalt, BkHmacKey), blockSize), PadBlock(hmacKey, blockSize));
        byte[] hmacValue = HmacSha512(hmacKey, package);
        byte[] encHmacValue = AesCbcEncrypt(secret,
            Fit(Hash(hashAlg, keyDataSalt, BkHmacValue), blockSize), PadBlock(hmacValue, blockSize));

        string xml = BuildEncryptionInfoXml(
            blockSize, saltSize, keyBytes * 8, hashSize, hashAlg, keyDataSalt,
            encHmacKey, encHmacValue, spin, pwSalt, encVerifierHashInput, encVerifierHashValue, encKeyValue);
        var info = new byte[8 + Encoding.UTF8.GetByteCount(xml)];
        BinaryPrimitives.WriteUInt16LittleEndian(info.AsSpan(0), 4);     // version major
        BinaryPrimitives.WriteUInt16LittleEndian(info.AsSpan(2), 4);     // version minor
        BinaryPrimitives.WriteUInt32LittleEndian(info.AsSpan(4), 0x40);  // fAgileReserved
        Encoding.UTF8.GetBytes(xml, 0, xml.Length, info, 8);

        return CompoundFile.WriteStreams(new[]
        {
            ("EncryptionInfo", info),
            ("EncryptedPackage", package),
            // Constant DataSpaces transform tree Office expects on encrypted files.
            ("DataSpaces/Version", DataSpaces.Version),
            ("DataSpaces/DataSpaceMap", DataSpaces.DataSpaceMap),
            ("DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace", DataSpaces.StrongEncryptionDataSpace),
            ("DataSpaces/TransformInfo/StrongEncryptionTransform/Primary", DataSpaces.Primary),
        });
    }

    // ==================== Agile EncryptionInfo ====================

    private sealed class AgileDescriptor
    {
        public string HashAlg = "SHA512";
        public int BlockSize, SaltSize, KeyBytes, SpinCount;
        public byte[] PwSalt = Array.Empty<byte>();
        public byte[] EncVerifierHashInput = Array.Empty<byte>();
        public byte[] EncVerifierHashValue = Array.Empty<byte>();
        public byte[] EncKeyValue = Array.Empty<byte>();
        // keyData (package) parameters — can differ from the password keyEncryptor.
        public string KeyDataHashAlg = "SHA512";
        public int KeyDataBlockSize, KeyDataKeyBytes;
        public byte[] KeyDataSalt = Array.Empty<byte>();

        public static AgileDescriptor Parse(byte[] info)
        {
            // First 8 bytes are the version/flags header; the rest is the XML.
            var doc = XDocument.Parse(Encoding.UTF8.GetString(info, 8, info.Length - 8));
            XElement root = doc.Root ?? throw new InvalidOperationException("EncryptionInfo: no root.");
            XElement keyData = root.Element(NsEnc + "keyData")
                ?? throw new InvalidOperationException("EncryptionInfo: no keyData.");
            XElement encKey = root.Element(NsEnc + "keyEncryptors")?.Element(NsEnc + "keyEncryptor")
                ?.Element(NsPwd + "encryptedKey")
                ?? throw new InvalidOperationException("EncryptionInfo: no password encryptedKey.");

            int Int(XElement e, string a) => int.Parse(e.Attribute(a)?.Value
                ?? throw new InvalidOperationException($"EncryptionInfo: missing @{a}."));
            byte[] B64(XElement e, string a) => Convert.FromBase64String(e.Attribute(a)?.Value
                ?? throw new InvalidOperationException($"EncryptionInfo: missing @{a}."));

            return new AgileDescriptor
            {
                KeyDataHashAlg = keyData.Attribute("hashAlgorithm")?.Value ?? "SHA512",
                KeyDataBlockSize = Int(keyData, "blockSize"),
                KeyDataKeyBytes = Int(keyData, "keyBits") / 8,
                KeyDataSalt = B64(keyData, "saltValue"),
                HashAlg = encKey.Attribute("hashAlgorithm")?.Value ?? "SHA512",
                BlockSize = Int(encKey, "blockSize"),
                SaltSize = Int(encKey, "saltSize"),
                KeyBytes = Int(encKey, "keyBits") / 8,
                SpinCount = Int(encKey, "spinCount"),
                PwSalt = B64(encKey, "saltValue"),
                EncVerifierHashInput = B64(encKey, "encryptedVerifierHashInput"),
                EncVerifierHashValue = B64(encKey, "encryptedVerifierHashValue"),
                EncKeyValue = B64(encKey, "encryptedKeyValue"),
            };
        }
    }

    private static string BuildEncryptionInfoXml(
        int blockSize, int saltSize, int keyBits, int hashSize, string hashAlg, byte[] keyDataSalt,
        byte[] encHmacKey, byte[] encHmacValue, int spin, byte[] pwSalt,
        byte[] encVerifierHashInput, byte[] encVerifierHashValue, byte[] encKeyValue)
    {
        string B(byte[] b) => Convert.ToBase64String(b);
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
            "<encryption xmlns=\"http://schemas.microsoft.com/office/2006/encryption\" " +
            "xmlns:p=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\" " +
            "xmlns:c=\"http://schemas.microsoft.com/office/2006/keyEncryptor/certificate\">" +
            $"<keyData saltSize=\"{saltSize}\" blockSize=\"{blockSize}\" keyBits=\"{keyBits}\" " +
            $"hashSize=\"{hashSize}\" cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" " +
            $"hashAlgorithm=\"{hashAlg}\" saltValue=\"{B(keyDataSalt)}\"/>" +
            $"<dataIntegrity encryptedHmacKey=\"{B(encHmacKey)}\" encryptedHmacValue=\"{B(encHmacValue)}\"/>" +
            "<keyEncryptors><keyEncryptor uri=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\">" +
            $"<p:encryptedKey spinCount=\"{spin}\" saltSize=\"{saltSize}\" blockSize=\"{blockSize}\" " +
            $"keyBits=\"{keyBits}\" hashSize=\"{hashSize}\" cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" " +
            $"hashAlgorithm=\"{hashAlg}\" saltValue=\"{B(pwSalt)}\" " +
            $"encryptedVerifierHashInput=\"{B(encVerifierHashInput)}\" " +
            $"encryptedVerifierHashValue=\"{B(encVerifierHashValue)}\" " +
            $"encryptedKeyValue=\"{B(encKeyValue)}\"/></keyEncryptor></keyEncryptors></encryption>";
    }

    // ==================== Primitives ====================

    /// <summary>H_0 = Hash(salt ‖ UTF16LE(password)); then iterate
    /// H_{i+1} = Hash(LE32(i) ‖ H_i) spinCount times.</summary>
    private static byte[] SpinHash(string alg, byte[] salt, string password, int spin)
    {
        byte[] h = Hash(alg, salt, Encoding.Unicode.GetBytes(password));
        var buf = new byte[4 + h.Length];
        for (int i = 0; i < spin; i++)
        {
            BinaryPrimitives.WriteInt32LittleEndian(buf.AsSpan(0), i);
            Array.Copy(h, 0, buf, 4, h.Length);
            h = Hash(alg, buf);
        }
        return h;
    }

    /// <summary>key = Hash(hSpin ‖ blockKey), truncated/zero-padded to keyBytes.</summary>
    private static byte[] DeriveKey(string alg, byte[] hSpin, byte[] blockKey, int keyBytes)
        => Fit(Hash(alg, hSpin, blockKey), keyBytes);

    private static byte[] Hash(string alg, params byte[][] parts)
    {
        using var h = CreateHash(alg);
        foreach (var p in parts) h.AppendData(p);
        return h.GetHashAndReset();
    }

    private static IncrementalHash CreateHash(string alg) => alg.ToUpperInvariant() switch
    {
        "SHA512" => IncrementalHash.CreateHash(HashAlgorithmName.SHA512),
        "SHA384" => IncrementalHash.CreateHash(HashAlgorithmName.SHA384),
        "SHA256" => IncrementalHash.CreateHash(HashAlgorithmName.SHA256),
        "SHA1" => IncrementalHash.CreateHash(HashAlgorithmName.SHA1),
        _ => throw new UnsupportedSchemeException($"Unsupported hash algorithm '{alg}'."),
    };

    private static byte[] HmacSha512(byte[] key, byte[] data) => HMACSHA512.HashData(key, data);

    private static byte[] AesCbcDecrypt(byte[] key, byte[] iv, byte[] data)
    {
        using var aes = Aes.Create();
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.None;
        aes.Key = key;
        aes.IV = iv;
        using var dec = aes.CreateDecryptor();
        return dec.TransformFinalBlock(data, 0, data.Length);
    }

    private static byte[] AesCbcEncrypt(byte[] key, byte[] iv, byte[] data)
    {
        using var aes = Aes.Create();
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.None;
        aes.Key = key;
        aes.IV = iv;
        using var enc = aes.CreateEncryptor();
        return enc.TransformFinalBlock(data, 0, data.Length);
    }

    /// <summary>Truncate or zero-pad <paramref name="b"/> to exactly <paramref name="n"/> bytes.</summary>
    private static byte[] Fit(byte[] b, int n)
    {
        if (b.Length == n) return b;
        var r = new byte[n];
        Array.Copy(b, r, Math.Min(b.Length, n));
        return r;
    }

    /// <summary>Zero-pad up to a whole multiple of <paramref name="block"/> (AES needs block-aligned input).</summary>
    private static byte[] PadBlock(byte[] b, int block)
    {
        int pad = (block - b.Length % block) % block;
        if (pad == 0) return b;
        var r = new byte[b.Length + pad];
        Array.Copy(b, r, b.Length);
        return r;
    }

    private static byte[] IntLe(int v)
    {
        var b = new byte[4];
        BinaryPrimitives.WriteInt32LittleEndian(b, v);
        return b;
    }

    /// <summary>
    /// Constant DataSpaces streams Office writes alongside the encrypted package
    /// to describe the "StrongEncryptionTransform". They carry no per-file
    /// secrets, so they are emitted verbatim. Captured from an Office/Agile file.
    /// </summary>
    private static class DataSpaces
    {
        public static readonly byte[] Version = Convert.FromBase64String(
            "PAAAAE0AaQBjAHIAbwBzAG8AZgB0AC4AQwBvAG4AdABhAGkAbgBlAHIALgBEAGEAdABhAFMAcABhAGMAZQBzAAEAAAABAAAAAQAAAA==");
        public static readonly byte[] DataSpaceMap = Convert.FromBase64String(
            "CAAAAAEAAABoAAAAAQAAAAAAAAAgAAAARQBuAGMAcgB5AHAAdABlAGQAUABhAGMAawBhAGcAZQAyAAAAUwB0AHIAbwBuAGcARQBuAGMAcgB5AHAAdABpAG8AbgBEAGEAdABhAFMAcABhAGMAZQAAAA==");
        public static readonly byte[] StrongEncryptionDataSpace = Convert.FromBase64String(
            "CAAAAAEAAAAyAAAAUwB0AHIAbwBuAGcARQBuAGMAcgB5AHAAdABpAG8AbgBUAHIAYQBuAHMAZgBvAHIAbQAAAA==");
        public static readonly byte[] Primary = Convert.FromBase64String(
            "WAAAAAEAAABMAAAAewBGAEYAOQBBADMARgAwADMALQA1ADYARQBGAC0ANAA2ADEAMwAtAEIARABEADUALQA1AEEANAAxAEMAMQBEADAANwAyADQANgB9AE4AAABNAGkAYwByAG8AcwBvAGYAdAAuAEMAbwBuAHQAYQBpAG4AZQByAC4ARQBuAGMAcgB5AHAAdABpAG8AbgBUAHIAYQBuAHMAZgBvAHIAbQAAAAEAAAABAAAAAQAAAAAAAAAAAAAAAAAAAAQAAAA=");
    }
}
