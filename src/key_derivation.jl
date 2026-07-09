# ─── ECMA-376 §2.3.4.11  key derivation ──────────────────────────────────────

const HASH_NAME_MAP = Dict(
    "SHA512" => SHA.sha512,
    "SHA256" => SHA.sha256,
    "SHA1"   => SHA.sha1,
    "SHA384" => SHA.sha384,
)

# Convert a UInt32 to 4 little-endian bytes
function uint32le(x::UInt32)::Vector{UInt8}
    return UInt8[(x >> 0) & 0xff, (x >> 8) & 0xff, (x >> 16) & 0xff, (x >> 24) & 0xff]
end

function get_hash_fn(alg::String)
    fn = get(HASH_NAME_MAP, alg, nothing)
    fn === nothing && error("Unsupported hash algorithm: $alg")
    fn
end

# NOTE: block_size removed — it was unused by this derivation (ECMA-376
# §2.3.4.11 doesn't need it; only key_bits determines output length).
function derive_key(password::String, salt::Vector{UInt8}, spin_count::Int,
                    hash_alg::String, key_bits::Int,
                    block_key::Vector{UInt8})
    hash_fn = get_hash_fn(hash_alg)

    # UTF-16LE encode the password. reinterpret gives byte-identical output to
    # a manual little-endian split on all realistic (little-endian) hosts.
    pwd_utf16 = reinterpret(UInt8, transcode(UInt16, password))

    # Step 1: H(salt || UTF-16LE(password))
    h_bytes = hash_fn([salt; pwd_utf16])

    # Step 2: iterate spin_count times: H(LE32(i) || h_bytes)
    for i in 0:(spin_count - 1)
        h_bytes = hash_fn([uint32le(UInt32(i)); h_bytes])
    end

    # Step 3: H(h_bytes || block_key), truncate/pad to key_bits÷8 bytes
    dk = hash_fn([h_bytes; block_key])

    key_len = key_bits ÷ 8
    if length(dk) < key_len
        append!(dk, fill(0x36, key_len - length(dk)))
    end
    return dk[1:key_len]
end

# ─── AES-CBC decrypt via Nettle ───────────────────────────────────────────────

function aes_cbc_decrypt(key::Vector{UInt8}, iv::Vector{UInt8},
                          ciphertext::Vector{UInt8})
    cipher_name = "AES$(length(key) * 8)"
    dec = Decryptor(cipher_name, key)
    return decrypt(dec, :CBC, iv, ciphertext)
end