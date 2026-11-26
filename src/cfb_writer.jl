# ─── CFB (OLE2 Compound File Binary) minimal writer ──────────────────────────
#
# Writes a version 3 compound file (512-byte sectors) per MS-CFB. Streams are
# given as "storage/substorage/stream" => bytes pairs; intermediate storages are
# created automatically. Streams < 4096 bytes live in the mini stream, larger
# ones in regular sectors; DIFAT sectors are emitted when > 109 FAT sectors are
# needed. Directory siblings form a balanced, correctly coloured red-black tree.

const NOSTREAM          = UInt32(0xFFFFFFFF)
const CFB_SECTOR_SIZE   = 512
const CFB_MINI_SIZE     = 64
const CFB_MINI_CUTOFF   = 4096
const CFB_IDS_PER_SECT  = CFB_SECTOR_SIZE ÷ 4      # 128 FAT / mini-FAT entries per sector
const CFB_HEADER_DIFAT  = 109

const CFB_TYPE_STORAGE = 0x01
const CFB_TYPE_STREAM  = 0x02
const CFB_TYPE_ROOT    = 0x05
const CFB_RED   = 0x00
const CFB_BLACK = 0x01

mutable struct CFBEntry
    name::String
    type::UInt8
    data::Vector{UInt8}
    children::Vector{CFBEntry}
    sid::UInt32
    left::UInt32
    right::UInt32
    child::UInt32
    color::UInt8
    start::UInt32
    size::UInt64
end
# Storages must have start sector 0 and size 0; empty streams use ENDOFCHAIN
CFBEntry(name, type, data=UInt8[]) =
    CFBEntry(name, type, data, CFBEntry[], 0, NOSTREAM, NOSTREAM, NOSTREAM, CFB_BLACK,
             type == CFB_TYPE_STREAM ? ENDOFCHAIN : UInt32(0), 0)

# MS-CFB §2.6.4 ordering: shorter names first, then by upper-cased UTF-16 code units
cfb_sort_key(e::CFBEntry) = (length(transcode(UInt16, e.name)), transcode(UInt16, uppercase(e.name)))

function cfb_build_tree(streams)
    root = CFBEntry("Root Entry", CFB_TYPE_ROOT)
    for (path, data) in streams
        parts = split(path, '/')
        node  = root
        for (k, part) in enumerate(parts)
            name = String(part)
            (isempty(name) || length(transcode(UInt16, name)) > 31) &&
                error("Invalid CFB entry name '$name' in '$path'")
            idx = findfirst(c -> uppercase(c.name) == uppercase(name), node.children)
            if k == length(parts)
                idx === nothing || error("Duplicate CFB entry '$path'")
                push!(node.children, CFBEntry(name, CFB_TYPE_STREAM, Vector{UInt8}(data)))
            elseif idx === nothing
                push!(node.children, CFBEntry(name, CFB_TYPE_STORAGE))
                node = node.children[end]
            else
                node = node.children[idx]
                node.type == CFB_TYPE_STORAGE || error("'$name' in '$path' is a stream, not a storage")
            end
        end
    end
    return root
end

# Flatten depth-first (root = SID 0) and link every sibling set as a balanced BST.
# Middle-split keeps all null links within one level of each other, so colouring
# the deepest level red (and everything else black) satisfies the red-black rules.
function cfb_assign!(root::CFBEntry)
    entries = CFBEntry[]
    visit(e) = (e.sid = length(entries); push!(entries, e); foreach(visit, e.children))
    visit(root)

    function link(sorted, lo, hi, depth, maxdepth)
        lo > hi && return NOSTREAM
        mid = (lo + hi) ÷ 2
        e = sorted[mid]
        e.left  = link(sorted, lo, mid - 1, depth + 1, maxdepth)
        e.right = link(sorted, mid + 1, hi, depth + 1, maxdepth)
        e.color = (depth == maxdepth && maxdepth > 0) ? CFB_RED : CFB_BLACK
        return e.sid
    end
    for e in entries
        isempty(e.children) && continue
        sorted   = sort(e.children; by=cfb_sort_key)
        maxdepth = floor(Int, log2(length(sorted)))      # depth of deepest node
        e.child  = link(sorted, 1, length(sorted), 0, maxdepth)
    end
    return entries
end

function write_cfb(io::IO, streams)
    root    = cfb_build_tree(streams)
    entries = cfb_assign!(root)
    all_streams = filter(e -> e.type == CFB_TYPE_STREAM, entries)
    any(e -> length(e.data) > typemax(UInt32), all_streams) &&
        error("Streams larger than 4 GB are not supported (CFB version 3)")

    # ── Mini stream: small streams, each padded to 64-byte mini sectors
    mini_fat    = UInt32[]
    mini_stream = UInt8[]
    for e in all_streams
        e.size = length(e.data)
        (e.size == 0 || e.size >= CFB_MINI_CUTOFF) && continue
        n = cld(length(e.data), CFB_MINI_SIZE)
        e.start = length(mini_fat)
        append!(mini_fat, UInt32.(e.start .+ (1:n-1)), [ENDOFCHAIN])
        append!(mini_stream, pad_to_block(e.data, CFB_MINI_SIZE))
    end
    big = filter(e -> e.size >= CFB_MINI_CUTOFF, all_streams)

    # ── Sector counts
    n_dir      = cld(length(entries), CFB_SECTOR_SIZE ÷ 128)
    n_minifat  = cld(length(mini_fat), CFB_IDS_PER_SECT)
    n_ministrm = cld(length(mini_stream), CFB_SECTOR_SIZE)
    n_big      = [cld(length(e.data), CFB_SECTOR_SIZE) for e in big]
    n_data     = n_dir + n_minifat + n_ministrm + sum(n_big; init=0)

    # FAT must also cover its own sectors and any DIFAT sectors: iterate to a fixpoint
    n_fat, n_difat = 1, 0
    while true
        n_difat = n_fat > CFB_HEADER_DIFAT ? cld(n_fat - CFB_HEADER_DIFAT, CFB_IDS_PER_SECT - 1) : 0
        need    = cld(n_data + n_fat + n_difat, CFB_IDS_PER_SECT)
        need == n_fat && break
        n_fat = need
    end

    # ── Layout: [FAT][DIFAT][directory][mini-FAT][mini stream][big streams]
    fat  = fill(FREESECT, n_fat * CFB_IDS_PER_SECT)
    next = 0
    function alloc!(n, marker=nothing)
        n == 0 && return ENDOFCHAIN
        first = next
        for s in first:(first + n - 1)
            fat[s + 1] = marker !== nothing ? marker : (s == first + n - 1 ? ENDOFCHAIN : UInt32(s + 1))
        end
        next += n
        return UInt32(first)
    end
    fat_start   = alloc!(n_fat, FATSECT)
    difat_start = alloc!(n_difat, DIFSECT)
    dir_start   = alloc!(n_dir)
    mfat_start  = alloc!(n_minifat)
    root.start  = alloc!(n_ministrm)
    root.size   = length(mini_stream)
    for (e, n) in zip(big, n_big)
        e.start = alloc!(n)
    end
    fat_sectors = collect(UInt32, fat_start .+ (0:n_fat-1))

    # ── Header (512 bytes)
    w32(x) = write(io, htol(UInt32(x)))
    write(io, CFB_MAGIC, zeros(UInt8, 16))                         # magic, CLSID
    write(io, htol(UInt16(0x003E)), htol(UInt16(3)), htol(UInt16(0xFFFE)),
              htol(UInt16(9)), htol(UInt16(6)), zeros(UInt8, 6))   # versions, BOM, shifts
    w32(0); w32(n_fat); w32(dir_start); w32(0); w32(CFB_MINI_CUTOFF)
    w32(n_minifat == 0 ? ENDOFCHAIN : mfat_start); w32(n_minifat)
    w32(n_difat == 0 ? ENDOFCHAIN : difat_start); w32(n_difat)
    for i in 1:CFB_HEADER_DIFAT
        w32(i <= n_fat ? fat_sectors[i] : FREESECT)
    end

    # ── FAT sectors
    foreach(w32, fat)

    # ── DIFAT sectors: 127 FAT locations + pointer to the next DIFAT sector
    rest = fat_sectors[CFB_HEADER_DIFAT+1:end]
    for d in 0:(n_difat - 1)
        chunk = rest[d*127+1 : min((d+1)*127, length(rest))]
        foreach(w32, chunk)
        foreach(_ -> w32(FREESECT), 1:(127 - length(chunk)))
        w32(d == n_difat - 1 ? ENDOFCHAIN : difat_start + d + 1)
    end

    # ── Directory sectors (128-byte entries, unused slots empty)
    for e in entries
        units = transcode(UInt16, e.name)
        name  = zeros(UInt8, 64)
        copyto!(name, reinterpret(UInt8, units))
        write(io, name, htol(UInt16(2 * (length(units) + 1))), e.type, e.color)
        w32(e.left); w32(e.right); w32(e.child)
        write(io, zeros(UInt8, 16 + 4 + 8 + 8))                    # CLSID, state, times
        w32(e.start); write(io, htol(UInt64(e.size)))
    end
    for _ in (length(entries) + 1):(n_dir * 4)
        write(io, zeros(UInt8, 66), 0x00, 0x00)
        w32(NOSTREAM); w32(NOSTREAM); w32(NOSTREAM)
        write(io, zeros(UInt8, 16 + 4 + 8 + 8 + 4 + 8))
    end

    # ── Mini-FAT, mini stream container, big streams (each padded to a sector)
    foreach(w32, mini_fat)
    foreach(_ -> w32(FREESECT), 1:(n_minifat * CFB_IDS_PER_SECT - length(mini_fat)))
    write(io, pad_to_block(mini_stream, CFB_SECTOR_SIZE))
    for e in big
        write(io, pad_to_block(e.data, CFB_SECTOR_SIZE))
    end
    return nothing
end
