# ─── XML.jl helper: depth-first search for first element with a given tag ────

# Match on local name only (strips namespace prefix like "p:" or "enc:")
local_tag(node) = let t = XML.tag(node); something(findfirst(':', t), 0) == 0 ? t : t[findfirst(':', t)+1:end] end

function find_node(node, target_tag::String)
    XML.nodetype(node) == XML.Element || return nothing
    local_tag(node) == target_tag && return node
    for child in XML.children(node)
        XML.nodetype(child) == XML.Element || continue
        result = find_node(child, target_tag)
        result !== nothing && return result
    end
    return nothing
end

# ─── ECMA-376 Agile Encryption: build EncryptionInfo stream ──────────────────

# Version 4.4, reserved flags 0x40 (Agile), as written by Excel 2010+
const AGILE_INFO_HEADER = UInt8[0x04, 0x00, 0x04, 0x00, 0x40, 0x00, 0x00, 0x00]

# Inverse of `parse_encryption_info`: serialise the parameter set `p` (same field
# names) into an EncryptionInfo stream, mirroring Excel's XML layout exactly.
function build_encryption_info(p)::Vector{UInt8}
    b64 = base64encode
    xml = string(
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n""",
        """<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" """,
        """xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" """,
        """xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">""",
        """<keyData saltSize="$(p.salt_size)" blockSize="$(p.block_size)" keyBits="$(p.key_bits)" """,
        """hashSize="$(hash_size(p.hash_alg))" cipherAlgorithm="$(p.cipher_alg)" """,
        """cipherChaining="$(p.cipher_chaining)" hashAlgorithm="$(p.hash_alg)" """,
        """saltValue="$(b64(p.key_data_salt))"/>""",
        """<dataIntegrity encryptedHmacKey="$(b64(p.enc_hmac_key))" """,
        """encryptedHmacValue="$(b64(p.enc_hmac_value))"/>""",
        """<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">""",
        """<p:encryptedKey spinCount="$(p.spin_count)" saltSize="$(p.enc_salt_size)" """,
        """blockSize="$(p.enc_block_size)" keyBits="$(p.enc_key_bits)" """,
        """hashSize="$(hash_size(p.enc_hash_alg))" cipherAlgorithm="$(p.cipher_alg)" """,
        """cipherChaining="$(p.cipher_chaining)" hashAlgorithm="$(p.enc_hash_alg)" """,
        """saltValue="$(b64(p.enc_key_salt))" """,
        """encryptedVerifierHashInput="$(b64(p.enc_verifier_hash_input))" """,
        """encryptedVerifierHashValue="$(b64(p.enc_verifier_hash_value))" """,
        """encryptedKeyValue="$(b64(p.enc_key_value))"/>""",
        """</keyEncryptor></keyEncryptors></encryption>""",
    )
    return [AGILE_INFO_HEADER; Vector{UInt8}(xml)]
end

# ─── ECMA-376 Agile Encryption: parse EncryptionInfo XML ─────────────────────

function parse_encryption_info(info_bytes::Vector{UInt8})
    # Bytes 1–8 are a version/reserved header; XML starts at byte 9
    xml_str = String(info_bytes[9:end])
    doc     = XML.parse(XML.Node, xml_str)

    # XML.parse may return a Document node; unwrap to the first Element child
    root = if XML.nodetype(doc) == XML.Element
        doc
    else
        first(c for c in XML.children(doc) if XML.nodetype(c) == XML.Element)
    end

    kd = find_node(root, "keyData")
    ek = find_node(root, "encryptedKey")

    kd === nothing && error("Could not find <keyData> in EncryptionInfo")
    ek === nothing && error("Could not find <encryptedKey> in EncryptionInfo")

    ka = XML.attributes(kd)
    ea = XML.attributes(ek)

    # <dataIntegrity> is optional for reading (decrypt never requires it)
    di = find_node(root, "dataIntegrity")
    da = di === nothing ? nothing : XML.attributes(di)

    return (
        # keyData attributes (used for final package decryption)
        cipher_alg      = ka["cipherAlgorithm"],
        cipher_chaining = ka["cipherChaining"],
        hash_alg        = ka["hashAlgorithm"],
        key_bits        = parse(Int, ka["keyBits"]),
        block_size      = parse(Int, ka["blockSize"]),
        salt_size       = parse(Int, ka["saltSize"]),
        key_data_salt   = base64decode(ka["saltValue"]),

        # encryptedKey attributes (used for key unwrapping + password verification)
        spin_count      = parse(Int, ea["spinCount"]),
        enc_key_salt    = base64decode(ea["saltValue"]),
        enc_salt_size   = parse(Int, ea["saltSize"]),
        enc_block_size  = parse(Int, ea["blockSize"]),
        enc_hash_alg    = ea["hashAlgorithm"],
        enc_key_bits    = parse(Int, ea["keyBits"]),
        enc_verifier_hash_input = base64decode(ea["encryptedVerifierHashInput"]),
        enc_verifier_hash_value = base64decode(ea["encryptedVerifierHashValue"]),
        enc_key_value           = base64decode(ea["encryptedKeyValue"]),

        # dataIntegrity attributes (HMAC over the EncryptedPackage stream), or nothing
        enc_hmac_key   = da === nothing ? nothing : base64decode(da["encryptedHmacKey"]),
        enc_hmac_value = da === nothing ? nothing : base64decode(da["encryptedHmacValue"]),
    )
end