module XLSXDecrypt

using Nettle   # AES-CBC and HMAC (see crypto.jl)
using SHA      # hashing for key derivation and IVs
using XML
using Base64
using Random: AbstractRNG, RandomDevice   # OS CSPRNG for keys and salts

export decrypt_xlsx, encrypt_xlsx

include("cfb.jl")
include("encryption_info.jl")
include("crypto.jl")
include("cfb_writer.jl")
include("dataspaces.jl")
include("encrypt.jl")
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

    # 3–5. Derive password keys, verify the password, unwrap the package key
    actual_key = unwrap_package_key(password, p)

    # 6. Decrypt EncryptedPackage
    #    First 8 bytes = uint64LE giving the true plaintext size
    plaintext_size = only(reinterpret(UInt64, enc_pkg_bytes[1:8]))
    ciphertext     = @view enc_pkg_bytes[9:end]

    plaintext = crypt_package(actual_key, p, ciphertext; encrypt=false)
    return IOBuffer(plaintext[1:plaintext_size])
end

end # module