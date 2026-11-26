# ─── Low-level crypto primitives (Nettle) ────────────────────────────────────
#
# Shared by decryption and (future) encryption. ECMA-376 Agile Encryption uses
# raw AES-CBC with zero padding supplied by the caller — never PKCS#7 — so all
# AES inputs must already be a multiple of the 16-byte block size.

aes_cipher_name(key::AbstractVector{UInt8}) = "AES$(length(key) * 8)"

# Raw AES-CBC decrypt (no padding removal)
function aes_cbc_decrypt(key::Vector{UInt8}, iv::Vector{UInt8},
                         ciphertext::Vector{UInt8})
    return decrypt(Decryptor(aes_cipher_name(key), key), :CBC, iv, ciphertext)
end

# Raw AES-CBC encrypt (no padding added)
function aes_cbc_encrypt(key::Vector{UInt8}, iv::Vector{UInt8},
                         plaintext::Vector{UInt8})
    return encrypt(Encryptor(aes_cipher_name(key), key), :CBC, iv, plaintext)
end

# Zero-pad `v` up to a multiple of `bs` bytes (ECMA-376 padding, not PKCS#7).
# Always returns a fresh Vector.
function pad_to_block(v::AbstractVector{UInt8}, bs::Int)::Vector{UInt8}
    r = mod(length(v), bs)
    return r == 0 ? Vector{UInt8}(v) : [v; zeros(UInt8, bs - r)]
end

# ─── HMAC (for dataIntegrity, ECMA-376 §2.3.4.14) ────────────────────────────
#
# Not used by decryption yet; provided for encrypt_xlsx and optional integrity
# verification. Qualified Nettle names avoid clashing with SHA's update!/digest!.

const HMAC_NAME_MAP = Dict(
    "SHA512" => "sha512",
    "SHA384" => "sha384",
    "SHA256" => "sha256",
    "SHA1"   => "sha1",
)

function hmac(alg::String, key::AbstractVector{UInt8}, data::AbstractVector{UInt8})::Vector{UInt8}
    name = get(HMAC_NAME_MAP, alg, nothing)
    name === nothing && error("Unsupported HMAC algorithm: $alg")
    h = Nettle.HMACState(name, Vector{UInt8}(key))
    Nettle.update!(h, Vector{UInt8}(data))
    return Nettle.digest!(h)
end
