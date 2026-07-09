module XLSXDecrypt

using Nettle   # used only for AES-CBC decryption
using SHA      # used for all hashing
using XML
using Base64

export decrypt_xlsx

include("cfb.jl")
include("encryption_info.jl")
include("key_derivation.jl")

# ─── Main public function ─────────────────────────────────────────────────────

"""
    decrypt_xlsx(path::String, password::String) -> IOBuffer

Decrypts a password-protected .xlsx file and returns an `IOBuffer` containing
the plaintext .xlsx data. Pass the result directly to `XLSX.readxlsx` or
`XLSX.openxlsx`.

```julia
buf = decrypt_xlsx("protected.xlsx", "secret")
xf  = XLSX.readxlsx(buf)
```

Only the modern ECMA-376 Agile Encryption scheme (Excel 2010+) is supported.
"""
function decrypt_xlsx(path::String, password::String)::IOBuffer
    raw     = read(path)
    file_io = IOBuffer(raw)

    # 1. Extract the two CFB streams
    enc_info_bytes = read_cfb_stream(file_io, "EncryptionInfo")
    enc_pkg_bytes  = read_cfb_stream(file_io, "EncryptedPackage")

    # 2. Parse encryption parameters from the XML inside EncryptionInfo
    p = parse_encryption_info(enc_info_bytes)

    # ECMA-376 §2.3.4.11 Table 1 — fixed block key constants
    BLOCK_VERIFIER_INPUT = UInt8[0xfe,0xa7,0xd2,0x76,0x3b,0x4b,0x9e,0x79]
    BLOCK_VERIFIER_HASH  = UInt8[0xd7,0xaa,0x0f,0x6d,0x30,0x61,0x34,0x4e]
    BLOCK_KEY_VALUE      = UInt8[0x14,0x6e,0x0b,0xe7,0xab,0xac,0xd0,0xd6]

    # 3. Derive three intermediate keys from password + encryptedKey salt.
    #    Each is independent (they only diverge at the final hash step), so
    #    when multiple threads are available we run them concurrently —
    #    spin_count is typically ~100k SHA calls per key, so this is the
    #    single biggest cost in the whole function.
    block_keys = (BLOCK_VERIFIER_INPUT, BLOCK_VERIFIER_HASH, BLOCK_KEY_VALUE)

    key_vi, key_vh, key_kv = if Threads.nthreads() > 1
        tasks = [Threads.@spawn derive_key(password, p.enc_key_salt, p.spin_count,
                                            p.enc_hash_alg, p.enc_key_bits, bk)
                 for bk in block_keys]
        fetch.(tasks)
    else
        [derive_key(password, p.enc_key_salt, p.spin_count,
                    p.enc_hash_alg, p.enc_key_bits, bk)
         for bk in block_keys]
    end

    # 4. Verify the password.
    dec_vi = aes_cbc_decrypt(key_vi, p.enc_key_salt, p.enc_verifier_hash_input)
    dec_vh = aes_cbc_decrypt(key_vh, p.enc_key_salt, p.enc_verifier_hash_value)

    computed = get_hash_fn(p.enc_hash_alg)(dec_vi[1:p.enc_salt_size])

    computed == dec_vh[1:length(computed)] ||
        error("Wrong password (verifier mismatch)")

    # 5. Decrypt the actual encryption key
    actual_key = aes_cbc_decrypt(key_kv, p.enc_key_salt, p.enc_key_value)
    actual_key = actual_key[1:(p.key_bits ÷ 8)]

    # 6. Decrypt EncryptedPackage
    #    First 8 bytes = uint64LE giving the true plaintext size
    plaintext_size = only(reinterpret(UInt64, enc_pkg_bytes[1:8]))
    ciphertext     = @view enc_pkg_bytes[9:end]

    seg_size       = 4096
    n_segments     = cld(length(ciphertext), seg_size)
    plaintext      = UInt8[]
    sizehint!(plaintext, length(ciphertext))

    for i in 0:(n_segments - 1)
        # Per-segment IV = H(keyDataSalt || LE32(i))[1:block_size]
        seg_iv = get_hash_fn(p.hash_alg)([p.key_data_salt; uint32le(UInt32(i))])[1:p.block_size]

        seg_start  = i * seg_size + 1
        seg_end    = min((i + 1) * seg_size, length(ciphertext))
        seg_cipher = ciphertext[seg_start:seg_end]

        # Pad to AES block boundary if needed
        rem = mod(length(seg_cipher), p.block_size)
        rem != 0 && append!(seg_cipher, zeros(UInt8, p.block_size - rem))

        append!(plaintext, aes_cbc_decrypt(actual_key, seg_iv, seg_cipher))
    end

    return IOBuffer(plaintext[1:plaintext_size])
end

end # module