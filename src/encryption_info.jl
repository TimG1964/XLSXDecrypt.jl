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
    )
end