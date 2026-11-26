module XLSXDecrypt

using Nettle   # AES-CBC and HMAC (see crypto.jl)
using SHA      # hashing for key derivation and IVs
using XML
using Base64

export decrypt_xlsx

include("cfb.jl")
include("encryption_info.jl")
include("crypto.jl")
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

    # 3. Derive three intermediate keys from password + encryptedKey salt
    key_vi, key_vh, key_kv = derive_keys(password, p)

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

    plaintext = crypt_package(actual_key, p, ciphertext; encrypt=false)
    return IOBuffer(plaintext[1:plaintext_size])
end

end # module