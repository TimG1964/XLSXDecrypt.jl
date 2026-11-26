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

# ─── ECMA-376 §2.3.4.11 / §2.3.4.14 fixed block keys ─────────────────────────

const BLOCK_VERIFIER_INPUT  = UInt8[0xfe,0xa7,0xd2,0x76,0x3b,0x4b,0x9e,0x79]
const BLOCK_VERIFIER_HASH   = UInt8[0xd7,0xaa,0x0f,0x6d,0x30,0x61,0x34,0x4e]
const BLOCK_KEY_VALUE       = UInt8[0x14,0x6e,0x0b,0xe7,0xab,0xac,0xd0,0xd6]
const BLOCK_INTEGRITY_KEY   = UInt8[0x5f,0xb2,0xad,0x01,0x0c,0xb9,0xe1,0xf6]  # dataIntegrity (encrypt)
const BLOCK_INTEGRITY_VALUE = UInt8[0xa0,0x67,0x7f,0x02,0xb2,0x2c,0x84,0x33]  # dataIntegrity (encrypt)

# Derive the three password-based keys (verifier input, verifier hash, key value)
# from the <encryptedKey> parameters. Each is independent (they only diverge at
# the final hash step), so when multiple threads are available they run
# concurrently — spin_count is typically ~100k hash calls per key, the single
# biggest cost of both decryption and encryption.
function derive_keys(password::String, p; threaded::Bool = Threads.nthreads() > 1)
    block_keys = (BLOCK_VERIFIER_INPUT, BLOCK_VERIFIER_HASH, BLOCK_KEY_VALUE)
    derive(bk) = derive_key(password, p.enc_key_salt, p.spin_count,
                            p.enc_hash_alg, p.enc_key_bits, bk)

    keys = if threaded
        fetch.([Threads.@spawn derive(bk) for bk in block_keys])
    else
        [derive(bk) for bk in block_keys]
    end
    return (keys[1], keys[2], keys[3])
end

# ─── IVs ─────────────────────────────────────────────────────────────────────

# Per-segment IV for EncryptedPackage: H(keyDataSalt || LE32(i))[1:block_size]
segment_iv(hash_alg::String, salt::Vector{UInt8}, i::Integer, bs::Int) =
    get_hash_fn(hash_alg)([salt; uint32le(UInt32(i))])[1:bs]

# IV for a block-key-tagged value (e.g. dataIntegrity): H(keyDataSalt || blockKey)[1:block_size]
block_iv(hash_alg::String, salt::Vector{UInt8}, block_key::Vector{UInt8}, bs::Int) =
    get_hash_fn(hash_alg)([salt; block_key])[1:bs]

# ─── Password verification and package key unwrapping (§2.3.4.13) ─────────────

# Check the password against the encrypted verifier and return the decrypted
# package (intermediate) key. Errors with "Wrong password" on mismatch.
function unwrap_package_key(password::String, p)::Vector{UInt8}
    key_vi, key_vh, key_kv = derive_keys(password, p)

    dec_vi = aes_cbc_decrypt(key_vi, p.enc_key_salt, p.enc_verifier_hash_input)
    dec_vh = aes_cbc_decrypt(key_vh, p.enc_key_salt, p.enc_verifier_hash_value)

    computed = get_hash_fn(p.enc_hash_alg)(dec_vi[1:p.enc_salt_size])
    computed == dec_vh[1:length(computed)] ||
        error("Wrong password (verifier mismatch)")

    actual_key = aes_cbc_decrypt(key_kv, p.enc_key_salt, p.enc_key_value)
    return actual_key[1:(p.key_bits ÷ 8)]
end

# ─── Data integrity (§2.3.4.14) ──────────────────────────────────────────────

hash_size(alg::String) = length(get_hash_fn(alg)(UInt8[]))

integrity_key_iv(p)   = block_iv(p.hash_alg, p.key_data_salt, BLOCK_INTEGRITY_KEY,   p.block_size)
integrity_value_iv(p) = block_iv(p.hash_alg, p.key_data_salt, BLOCK_INTEGRITY_VALUE, p.block_size)

# Build the (encryptedHmacKey, encryptedHmacValue) pair for an EncryptedPackage
# stream `pkg` (including its 8-byte size prefix), given a random HMAC key.
function encrypt_integrity(p, package_key::Vector{UInt8}, hmac_key::Vector{UInt8},
                           pkg::AbstractVector{UInt8})
    mac = hmac(p.hash_alg, hmac_key, pkg)
    enc_key   = aes_cbc_encrypt(package_key, integrity_key_iv(p),   pad_to_block(hmac_key, p.block_size))
    enc_value = aes_cbc_encrypt(package_key, integrity_value_iv(p), pad_to_block(mac, p.block_size))
    return enc_key, enc_value
end

# Verify the HMAC in <dataIntegrity> against the whole EncryptedPackage stream
# (size prefix + ciphertext, as stored). Errors if the file has no <dataIntegrity>.
function verify_integrity(p, package_key::Vector{UInt8}, pkg::AbstractVector{UInt8})::Bool
    (p.enc_hmac_key === nothing || p.enc_hmac_value === nothing) &&
        error("No <dataIntegrity> element in EncryptionInfo")
    hs       = hash_size(p.hash_alg)
    hmac_key = aes_cbc_decrypt(package_key, integrity_key_iv(p),   p.enc_hmac_key)[1:hs]
    expected = aes_cbc_decrypt(package_key, integrity_value_iv(p), p.enc_hmac_value)[1:hs]
    return hmac(p.hash_alg, hmac_key, pkg) == expected
end

# ─── EncryptedPackage segment cipher (§2.3.4.15) ─────────────────────────────

const PACKAGE_SEGMENT_SIZE = 4096

# Encrypt or decrypt package bytes (excluding the 8-byte size prefix) in
# 4096-byte segments, each with its own IV. The final segment is zero-padded to
# the block size, so the output may be longer than `data`; callers trim (decrypt)
# or record the true size in the prefix (encrypt).
function crypt_package(key::Vector{UInt8}, p, data::AbstractVector{UInt8};
                       encrypt::Bool)::Vector{UInt8}
    aes = encrypt ? aes_cbc_encrypt : aes_cbc_decrypt
    n_segments = cld(length(data), PACKAGE_SEGMENT_SIZE)
    out = UInt8[]
    sizehint!(out, n_segments * PACKAGE_SEGMENT_SIZE)

    for i in 0:(n_segments - 1)
        seg_start = i * PACKAGE_SEGMENT_SIZE + 1
        seg_end   = min((i + 1) * PACKAGE_SEGMENT_SIZE, length(data))
        seg       = pad_to_block(@view(data[seg_start:seg_end]), p.block_size)
        iv        = segment_iv(p.hash_alg, p.key_data_salt, i, p.block_size)
        append!(out, aes(key, iv, seg))
    end
    return out
end