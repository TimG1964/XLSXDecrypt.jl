# ─── CFB (OLE2 Compound File Binary) minimal parser ──────────────────────────

const CFB_MAGIC  = UInt8[0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1]
const ENDOFCHAIN = UInt32(0xFFFFFFFE)   # end of a sector chain
const FREESECT   = UInt32(0xFFFFFFFF)   # unallocated sector
const FATSECT    = UInt32(0xFFFFFFFD)   # this sector IS a FAT sector
const DIFSECT    = UInt32(0xFFFFFFFC)   # this sector IS a DIFAT sector

# All special marker values are >= DIFSECT.
# IMPORTANT: arithmetic like FREESECT-2 wraps around on UInt32 and gives the
# wrong answer — always use this predicate instead of range comparisons.
is_special(s::UInt32) = s >= DIFSECT

function read_cfb_stream(io::IO, target_name::String)
    seekstart(io)
    read(io, 8) == CFB_MAGIC || error("Not a CFB/OLE2 file")

    # CFB header layout (byte offsets) per MS-CFB spec:
    #  0– 7  signature / magic (8 bytes)         ← already read
    #  8–23  CLSID (16 bytes, unused)
    # 24–25  minor version
    # 26–27  major version
    # 28–29  byte order mark (0xFFFE)
    # 30–31  sector size (power of 2)
    # 32–33  mini sector size (power of 2)
    # 34–39  reserved (6 bytes)
    # 40–43  num directory sectors (v4 only; 0 for v3)
    # 44–47  num FAT sectors
    # 48–51  first directory sector
    # 52–55  transaction signature
    # 56–59  mini stream cutoff size
    # 60–63  first mini-FAT sector
    # 64–67  num mini-FAT sectors
    # 68–71  first DIFAT sector
    # 72–75  num DIFAT sectors
    # 76–    inline DIFAT table (109 × 4 bytes = 436 bytes)
    seek(io, 30)
    sector_size      = 1 << read(io, UInt16)    # 30–31: sector size power (512 → 9)
    mini_sector_size = 1 << read(io, UInt16)    # 32–33: mini sector size power (64 → 6)
    seek(io, 44)
    num_fat          = read(io, UInt32)          # 44–47: number of FAT sectors
    first_dir_sector = read(io, UInt32)          # 48–51: first directory sector
    read(io, UInt32)                             # 52–55: transaction signature (skip)
    mini_cutoff      = read(io, UInt32)          # 56–59: mini stream cutoff (usually 4096)
    first_minifat    = read(io, UInt32)          # 60–63: first mini-FAT sector
    num_minifat      = read(io, UInt32)          # 64–67: number of mini-FAT sectors
    first_difat      = read(io, UInt32)          # 68–71: first DIFAT sector
    num_difat        = read(io, UInt32)          # 72–75: number of DIFAT sectors

    # ── Collect FAT sector locations from the inline DIFAT table (109 entries)
    difat = UInt32[]
    for _ in 1:109
        e = read(io, UInt32)
        !is_special(e) && push!(difat, e)
    end

    # ── Walk any extra DIFAT sectors (only for files with >109 FAT sectors — rare)
    sec = first_difat
    for _ in 1:num_difat
        is_special(sec) && break
        seek(io, (sec + 1) * sector_size)
        for _ in 1:(sector_size ÷ 4 - 1)
            e = read(io, UInt32)
            !is_special(e) && push!(difat, e)
        end
        sec = read(io, UInt32)
    end

    # ── Build the FAT — read exactly num_fat sectors
    fat = UInt32[]
    for s in difat[1:min(num_fat, length(difat))]
        seek(io, (s + 1) * sector_size)
        buf = read(io, sector_size)
        append!(fat, reinterpret(UInt32, buf))
    end

    # Follow a FAT chain one step. Returns ENDOFCHAIN sentinel on any problem.
    # This is intentionally lenient: 0, out-of-bounds, and special values all
    # terminate the chain, because some CFB writers use 0 instead of ENDOFCHAIN.
    function next_fat(s::UInt32)::UInt32
        is_special(s) && return ENDOFCHAIN
        idx = Int(s) + 1
        idx > length(fat) && return ENDOFCHAIN
        nxt = fat[idx]
        # Treat 0 as end-of-chain: some writers zero-fill unused FAT entries
        nxt == 0x00000000 && return ENDOFCHAIN
        nxt
    end

    # ── Build the mini-FAT (walk exactly num_minifat sectors)
    mini_fat = UInt32[]
    if num_minifat > 0 && !is_special(first_minifat)
        sec = first_minifat
        for _ in 1:num_minifat
            is_special(sec) && break
            seek(io, (sec + 1) * sector_size)
            buf = read(io, sector_size)
            append!(mini_fat, reinterpret(UInt32, buf))
            sec = next_fat(sec)
        end
    end

    # ── Read directory entries (128 bytes each).
    # Upper bound: a typical CFB has very few directory sectors (often just 1).
    # We cap at 8192 entries (1 MB of directory data) to prevent runaway reads.
    dir_entries = NamedTuple[]
    seen_dir    = Set{UInt32}()
    sec         = first_dir_sector
    max_dir_sectors = max(1, cld(8192 * 128, sector_size))
    for _ in 1:max_dir_sectors
        is_special(sec) && break
        sec in seen_dir && break        # genuine cycle — stop, don't error
        push!(seen_dir, sec)
        seek(io, (sec + 1) * sector_size)
        for _ in 1:(sector_size ÷ 128)
            raw_name  = read(io, 64)
            name_len  = read(io, UInt16)
            obj_type  = read(io, UInt8)
            read(io, 1)          # color flag
            read(io, 4)          # left-sibling SID
            read(io, 4)          # right-sibling SID
            read(io, 4)          # child SID
            read(io, 36)         # CLSID + state + timestamps
            start_sec = read(io, UInt32)
            stream_sz = read(io, UInt64)

            entry_name = ""
            if name_len >= 2
                nchars = (name_len - 2) ÷ 2
                chars = reinterpret(UInt16, raw_name[1:2*nchars])
                entry_name = transcode(String, Vector{UInt16}(chars))
            end
            push!(dir_entries, (name=entry_name, type=obj_type,
                                 start=start_sec, size=stream_sz))
        end
        sec = next_fat(sec)
    end

    isempty(dir_entries) && error("No directory entries found in CFB file")

    # ── Entry 0 is always root storage; its sector chain holds the mini-stream
    root_entry = dir_entries[1]

    mini_container = UInt8[]
    seen_mc = Set{UInt32}()
    sec = root_entry.start
    while !is_special(sec) && !(sec in seen_mc)
        push!(seen_mc, sec)
        seek(io, (sec + 1) * sector_size)
        append!(mini_container, read(io, sector_size))
        sec = next_fat(sec)
    end

    # ── Find the requested stream (directory entry type 0x02 = stream object)
    entry = nothing
    for e in dir_entries
        if e.name == target_name && e.type == 0x02
            entry = e; break
        end
    end
    entry === nothing && error("Stream '$target_name' not found in CFB file")

    # ── Read the stream data.
    # `remaining` is the authoritative termination condition — we stop as soon
    # as we have the declared number of bytes, regardless of FAT chain values.
    data      = UInt8[]
    sizehint!(data, entry.size)
    remaining = Int(entry.size)

    if entry.size < mini_cutoff
        # Small stream: data lives in the mini-stream
        sec = entry.start
        while remaining > 0
            is_special(sec) && error("mini-stream chain ended early for '$target_name'")
            off   = Int(sec) * mini_sector_size
            chunk = min(mini_sector_size, remaining)
            off + chunk > length(mini_container) && error("mini-stream out of bounds for '$target_name'")
            append!(data, mini_container[off+1 : off+chunk])
            remaining -= chunk
            remaining == 0 && break
            idx = Int(sec) + 1
            idx > length(mini_fat) && error("mini-FAT chain ended early for '$target_name'")
            sec = mini_fat[idx]
        end
    else
        # Large stream: data lives in normal FAT sectors
        sec = entry.start
        while remaining > 0
            is_special(sec) && error("FAT chain ended early for '$target_name'")
            seek(io, (sec + 1) * sector_size)
            chunk = min(sector_size, remaining)
            append!(data, read(io, chunk))
            remaining -= chunk
            remaining == 0 && break
            sec = next_fat(sec)
        end
    end

    return data
end