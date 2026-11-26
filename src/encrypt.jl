# ─── ECMA-376 Agile Encryption: encrypt an .xlsx package ─────────────────────

const ZIP_MAGIC        = UInt8[0x50, 0x4b, 0x03, 0x04]   # "PK\x03\x04"
const MAX_SPIN_COUNT   = 10_000_000                       # ECMA-376 §2.3.4.10 upper bound

read_input(x::AbstractString)        = read(x)
read_input(x::AbstractVector{UInt8}) = Vector{UInt8}(x)
# Seekable IO is read from the start (so a buffer just written by
# XLSX.writexlsx needs no seekstart); non-seekable streams from where they are.
read_input(x::IO)                    = read(x)
read_input(x::IOStream)              = read(seekstart(x))
read_input(x::IOBuffer)              = x.seekable ? read(seekstart(x)) : read(x)

write_output(x::AbstractString, bytes) = (write(x, bytes); nothing)
write_output(x::IO, bytes)             = (write(x, bytes); nothing)

# Encrypt plaintext package bytes and return the complete CFB file. Parameters
# match Excel's defaults: AES-256-CBC, SHA-512, 16-byte salts.
function encrypt_package(plain::Vector{UInt8}, password::String, spin_count::Int;
                         rng::AbstractRNG = RandomDevice())
    rnd(n) = rand(rng, UInt8, n)

    p = (cipher_alg = "AES", cipher_chaining = "ChainingModeCBC", hash_alg = "SHA512",
         key_bits = 256, block_size = 16, salt_size = 16, key_data_salt = rnd(16),
         spin_count = spin_count, enc_key_salt = rnd(16), enc_salt_size = 16,
         enc_block_size = 16, enc_hash_alg = "SHA512", enc_key_bits = 256)

    package_key = rnd(p.key_bits ÷ 8)
    verifier    = rnd(p.enc_salt_size)
    hmac_key    = rnd(hash_size(p.hash_alg))

    # Password keys wrap the verifier and the package key (IV = encryptedKey salt)
    key_vi, key_vh, key_kv = derive_keys(password, p)
    iv   = p.enc_key_salt
    bs   = p.enc_block_size
    evhi = aes_cbc_encrypt(key_vi, iv, pad_to_block(verifier, bs))
    evhv = aes_cbc_encrypt(key_vh, iv, pad_to_block(get_hash_fn(p.enc_hash_alg)(verifier), bs))
    ekv  = aes_cbc_encrypt(key_kv, iv, pad_to_block(package_key, bs))

    # EncryptedPackage = LE64(plaintext size) || segment-encrypted package
    pkg = [reinterpret(UInt8, [htol(UInt64(length(plain)))]);
           crypt_package(package_key, p, plain; encrypt=true)]

    ehk, ehv = encrypt_integrity(p, package_key, hmac_key, pkg)

    p = merge(p, (enc_verifier_hash_input = evhi, enc_verifier_hash_value = evhv,
                  enc_key_value = ekv, enc_hmac_key = ehk, enc_hmac_value = ehv))

    io = IOBuffer()
    write_cfb(io, ["EncryptionInfo"   => build_encryption_info(p),
                   "EncryptedPackage" => pkg,
                   DATASPACES_STREAMS...])
    return take!(io)
end

"""
    encrypt_xlsx(input, output, password::AbstractString; spin_count = 100_000) -> Nothing

Encrypts an `.xlsx` package with a password using ECMA-376 Agile Encryption
(AES-256, SHA-512 — the same scheme and defaults Excel uses), producing a file
Excel will open after prompting for the password.

- `input`: a file path, a byte vector, or an `IO`. A seekable `IO` (`IOBuffer`,
  open file) is read from the start, whatever its current position; a
  non-seekable stream (pipe, socket) is read from its current position.
- `output`: a file path or an `IO`. The input is read completely first, so
  `output` may be the same path as `input`.
- `spin_count`: key-derivation iterations (Excel uses 100,000; maximum 10,000,000).

```julia
encrypt_xlsx("plain.xlsx", "protected.xlsx", "secret")

# round trip
xf = XLSX.readxlsx(decrypt_xlsx("protected.xlsx", "secret"))

# encrypt a workbook straight from XLSX.jl, with no plaintext file on disk
buf = IOBuffer()
XLSX.writexlsx(buf, xf_writable)     # opened with mode="w"/"rw", or XLSX.newxlsx()
encrypt_xlsx(buf, "protected.xlsx", "secret")
```
"""
function encrypt_xlsx(input, output, password::AbstractString; spin_count::Integer = 100_000)
    0 <= spin_count <= MAX_SPIN_COUNT ||
        error("spin_count must be between 0 and $MAX_SPIN_COUNT, got $spin_count")
    plain = read_input(input)
    (length(plain) >= 4 && plain[1:4] == ZIP_MAGIC) ||
        error("Input is not an .xlsx package (missing ZIP signature)")
    write_output(output, encrypt_package(plain, String(password), Int(spin_count)))
end
