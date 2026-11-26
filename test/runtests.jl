import XLSX
import XLSXDecrypt as XD
using Test
using Dates
using SHA
using Random

data_directory = joinpath(dirname(pathof(XD)), "..", "test", "data")

@assert isdir(data_directory)

@testset "Basic functionality              " begin
    test_file = joinpath(data_directory, raw"password-is-w23$er3.xlsx")
    io=XD.decrypt_xlsx(test_file, raw"w23$er3")
    test_file2 = joinpath(data_directory, raw"password-is-very$long^password#3245301!.xlsx")
    io2=XD.decrypt_xlsx(test_file2, raw"very$long^password#3245301!")

    @testset "number formats               " begin
        XLSX.openxlsx(io) do f
            show(IOBuffer(), f)
            sheet = f["Sheet1"]
            @test sheet["A1"] == 1
            @test isapprox(sheet["B1"], 0.546832775750823)
            @test sheet["C1"] == "kjghfjvila"
            @test sheet["A2"] == 2
            @test isapprox(sheet["B2"], 0.381845788463574)
            @test sheet["C2"] == "ghfjkqwefg"
            @test isapprox(sheet["B3"], 0.541686223027816)

            @test sheet["A5"] == 10
            @test isapprox(sheet["B5"], 1.78799032829419)
            @test isapprox(f["Sheet2!B3"], 0.541686223027816)
            @test f["Sheet2!C3"] == "fhlAWETYUUI"
            @test isapprox(f["Sheet2!B4"], 0.317625541051977)
            @test f["Sheet2!C4"] == "HFJuwe"
        end
        XLSX.openxlsx(io2) do f
            show(IOBuffer(), f)
            sheet = f["Sheet1"]
            @test sheet["A1"] == 1
            @test isapprox(sheet["B1"], 0.546832775750823)
            @test sheet["C1"] == "kjghfjvila"
            @test sheet["A2"] == 2
            @test isapprox(sheet["B2"], 0.381845788463574)
            @test sheet["C2"] == "ghfjkqwefg"
            @test isapprox(sheet["B3"], 0.541686223027816)

            @test sheet["A5"] == 10
            @test isapprox(sheet["B5"], 1.78799032829419)
            @test isapprox(f["Sheet2!B3"], 0.541686223027816)
            @test f["Sheet2!C3"] == "fhlAWETYUUI"
            @test isapprox(f["Sheet2!B4"], 0.317625541051977)
            @test f["Sheet2!C4"] == "HFJuwe"
        end
    end

    @testset "Defined Names                " begin

        seekstart(io) 
        f = XLSX.openxlsx(io, mode="rw")
        s = f["Sheet2"]
        @test all(isapprox.(s["Floats"], [0.546832775750823; 0.381845788463574; 0.541686223027816; 0.317625541051977;;]))
        @test s["SortedStrings"] == Any["fhlAWETYUUI"; "ghfjkqwefg"; "HFJuwe"; "kjghfjvila";;]

        seekstart(io2) 
        f = XLSX.openxlsx(io2, mode="rw")
        s = f["Sheet2"]
        @test all(isapprox.(s["Floats"], [0.546832775750823; 0.381845788463574; 0.541686223027816; 0.317625541051977;;]))
        @test s["SortedStrings"] == Any["fhlAWETYUUI"; "ghfjkqwefg"; "HFJuwe"; "kjghfjvila";;]

    end
end

v= pkgversion(XLSX)
if (v.major, v.minor) >= (0, 11)

    @testset "Newer functionality              " begin

        test_file = joinpath(data_directory, raw"password-is-w23$er3.xlsx")
        io=XD.decrypt_xlsx(test_file, raw"w23$er3")
        f = XLSX.openxlsx(io, mode="rw")
        test_file = joinpath(data_directory, raw"password-is-very$long^password#3245301!.xlsx")
        io=XD.decrypt_xlsx(test_file, raw"very$long^password#3245301!")
        f2 = XLSX.openxlsx(io, mode="rw")

        @testset "formulas" begin
            s = f[1]
            wb = XLSX.get_workbook(s)
            @test XLSX.getcell(s, "A5") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A5"), "", "13", "10", "", true)
            @test XLSX.get_formula_from_cache(s, XLSX.CellRef("A5")) == XLSX.Formula("SUM(A1:A4)", nothing, nothing, nothing)
            @test XLSX.getcell(s, "D1").style == 6
            @test s["D1"] == "fhlAWETYUUI"
            @test XLSX.get_formula_from_cache(s, XLSX.CellRef("D1")) == XLSX.Formula("_xlfn._xlws.SORT(C1:C4)", "array", "D1:D4", nothing)

            s = f2[1]
            wb = XLSX.get_workbook(s)
            wb = XLSX.get_workbook(s)
            @test XLSX.getcell(s, "A5") == XLSX.Cell(XLSX.get_workbook(f), XLSX.CellRef("A5"), "", "13", "10", "", true)
            @test XLSX.get_formula_from_cache(s, XLSX.CellRef("A5")) == XLSX.Formula("SUM(A1:A4)", nothing, nothing, nothing)
            @test XLSX.getcell(s, "D1").style == 6
            @test s["D1"] == "fhlAWETYUUI"
            @test XLSX.get_formula_from_cache(s, XLSX.CellRef("D1")) == XLSX.Formula("_xlfn._xlws.SORT(C1:C4)", "array", "D1:D4", nothing)
        end

        @testset "formatting" begin

            s = f[1]
            @test XLSX.getFont(s, "A1").font == Dict("name" => Dict("val" => "Aptos Narrow"), "family" => Dict("val" => "2"), "sz" => Dict("val" => "11"), "color" => Dict("rgb" => "FF006100"), "scheme" => Dict("val" => "minor"))
            @test XLSX.getFont(s, "B2").font == Dict("name" => Dict("val" => "Aptos Narrow"), "family" => Dict("val" => "2"), "sz" => Dict("val" => "11"), "color" => Dict("rgb" => "FF9C5700"), "scheme" => Dict("val" => "minor"))
            @test XLSX.getFont(s, "C3").font == Dict("name" => Dict("val" => "Aptos Narrow"), "family" => Dict("val" => "2"), "sz" => Dict("val" => "11"), "color" => Dict("rgb" => "FF9C0006"), "scheme" => Dict("val" => "minor"))
            @test XLSX.getFont(s, "D4").font == Dict("name" => Dict("val" => "Aptos Narrow"), "family" => Dict("val" => "2"), "sz" => Dict("val" => "11"), "color" => Dict("theme" => "1"), "scheme" => Dict("val" => "minor"))
            @test XLSX.getFill(s, "D2").fill == Dict("patternFill" => Dict("patternType" => "solid", "fgrgb" => "FFFFFFCC"))
            @test XLSX.getFill(s, "C3").fill == Dict("patternFill" => Dict("patternType" => "solid", "fgrgb" => "FFFFC7CE"))
            @test XLSX.getFill(s, "B4").fill == Dict("patternFill" => Dict("patternType" => "solid", "fgrgb" => "FFFFEB9C"))
            @test XLSX.getFill(s, "A5").fill == Dict("patternFill" => Dict("patternType" => "solid", "fgrgb" => "FFF2F2F2"))
            @test XLSX.getBorder(s, "A5").border == Dict("left" => Dict("rgb" => "FF7F7F7F", "style" => "thick"), "bottom" => Dict("rgb" => "FF7F7F7F", "style" => "thick"), "right" => Dict("rgb" => "FF7F7F7F", "style" => "thin"), "top" => Dict("rgb" => "FF7F7F7F", "style" => "double"), "diagonal" => nothing)
            @test XLSX.getBorder(s, "B5").border == Dict("left" => Dict("rgb" => "FF7F7F7F", "style" => "thin"), "bottom" => Dict("rgb" => "FF7F7F7F", "style" => "thick"), "right" => Dict("rgb" => "FF7F7F7F", "style" => "thick"), "top" => Dict("rgb" => "FF7F7F7F", "style" => "double"), "diagonal" => nothing)
            @test XLSX.getBorder(s, "C2").border == Dict("left" => Dict("indexed" => "64", "style" => "thin"), "bottom" => Dict("indexed" => "64", "style" => "thin"), "right" => Dict("indexed" => "64", "style" => "thin"), "top" => Dict("indexed" => "64", "style" => "thin"), "diagonal" => nothing)
            @test XLSX.getBorder(s, "D1").border == Dict("left" => Dict("indexed" => "64", "style" => "thin"), "bottom" => Dict("indexed" => "64", "style" => "thin"), "right" => Dict("indexed" => "64", "style" => "medium"), "top" => Dict("indexed" => "64", "style" => "medium"), "diagonal" => nothing)

            s = f2[1]
            @test XLSX.getFont(s, "A1").font == Dict("name" => Dict("val" => "Aptos Narrow"), "family" => Dict("val" => "2"), "sz" => Dict("val" => "11"), "color" => Dict("rgb" => "FF006100"), "scheme" => Dict("val" => "minor"))
            @test XLSX.getFont(s, "B2").font == Dict("name" => Dict("val" => "Aptos Narrow"), "family" => Dict("val" => "2"), "sz" => Dict("val" => "11"), "color" => Dict("rgb" => "FF9C5700"), "scheme" => Dict("val" => "minor"))
            @test XLSX.getFont(s, "C3").font == Dict("name" => Dict("val" => "Aptos Narrow"), "family" => Dict("val" => "2"), "sz" => Dict("val" => "11"), "color" => Dict("rgb" => "FF9C0006"), "scheme" => Dict("val" => "minor"))
            @test XLSX.getFont(s, "D4").font == Dict("name" => Dict("val" => "Aptos Narrow"), "family" => Dict("val" => "2"), "sz" => Dict("val" => "11"), "color" => Dict("theme" => "1"), "scheme" => Dict("val" => "minor"))
            @test XLSX.getFill(s, "D2").fill == Dict("patternFill" => Dict("patternType" => "solid", "fgrgb" => "FFFFFFCC"))
            @test XLSX.getFill(s, "C3").fill == Dict("patternFill" => Dict("patternType" => "solid", "fgrgb" => "FFFFC7CE"))
            @test XLSX.getFill(s, "B4").fill == Dict("patternFill" => Dict("patternType" => "solid", "fgrgb" => "FFFFEB9C"))
            @test XLSX.getFill(s, "A5").fill == Dict("patternFill" => Dict("patternType" => "solid", "fgrgb" => "FFF2F2F2"))
            @test XLSX.getBorder(s, "A5").border == Dict("left" => Dict("rgb" => "FF7F7F7F", "style" => "thick"), "bottom" => Dict("rgb" => "FF7F7F7F", "style" => "thick"), "right" => Dict("rgb" => "FF7F7F7F", "style" => "thin"), "top" => Dict("rgb" => "FF7F7F7F", "style" => "double"), "diagonal" => nothing)
            @test XLSX.getBorder(s, "B5").border == Dict("left" => Dict("rgb" => "FF7F7F7F", "style" => "thin"), "bottom" => Dict("rgb" => "FF7F7F7F", "style" => "thick"), "right" => Dict("rgb" => "FF7F7F7F", "style" => "thick"), "top" => Dict("rgb" => "FF7F7F7F", "style" => "double"), "diagonal" => nothing)
            @test XLSX.getBorder(s, "C2").border == Dict("left" => Dict("indexed" => "64", "style" => "thin"), "bottom" => Dict("indexed" => "64", "style" => "thin"), "right" => Dict("indexed" => "64", "style" => "thin"), "top" => Dict("indexed" => "64", "style" => "thin"), "diagonal" => nothing)
            @test XLSX.getBorder(s, "D1").border == Dict("left" => Dict("indexed" => "64", "style" => "thin"), "bottom" => Dict("indexed" => "64", "style" => "thin"), "right" => Dict("indexed" => "64", "style" => "medium"), "top" => Dict("indexed" => "64", "style" => "medium"), "diagonal" => nothing)

        end
    end
end

@testset "Error handling                   " begin
    test_file = joinpath(data_directory, raw"password-is-w23$er3.xlsx")

    @testset "wrong password" begin
        @test_throws ErrorException XD.decrypt_xlsx(test_file, "definitely-wrong")
    end

    @testset "not a CFB file" begin
        # any plain file, e.g. a non-encrypted xlsx or this test file itself
        bad_file = joinpath(@__DIR__, "runtests.jl")
        @test_throws ErrorException XD.decrypt_xlsx(bad_file, "irrelevant")
    end

    @testset "file does not exist" begin
        @test_throws SystemError XD.decrypt_xlsx(joinpath(data_directory, "nope.xlsx"), "x")
    end
end

@testset "Internal helpers                 " begin
    @testset "uint32le" begin
        @test XD.uint32le(UInt32(1)) == UInt8[0x01, 0x00, 0x00, 0x00]
        @test XD.uint32le(UInt32(256)) == UInt8[0x00, 0x01, 0x00, 0x00]
        @test XD.uint32le(UInt32(0)) == UInt8[0x00, 0x00, 0x00, 0x00]
    end

    @testset "is_special" begin
        @test XD.is_special(XD.ENDOFCHAIN)
        @test XD.is_special(XD.FREESECT)
        @test XD.is_special(XD.FATSECT)
        @test XD.is_special(XD.DIFSECT)
        @test !XD.is_special(UInt32(0))
        @test !XD.is_special(UInt32(1_000_000))
    end

    @testset "derive_key is deterministic" begin
        salt = rand(UInt8, 16)
        bk   = UInt8[0x14,0x6e,0x0b,0xe7,0xab,0xac,0xd0,0xd6]
        k1 = XD.derive_key("password", salt, 100, "SHA512", 256, bk)
        k2 = XD.derive_key("password", salt, 100, "SHA512", 256, bk)
        @test k1 == k2
        @test length(k1) == 32  # 256 bits

        # different password -> different key
        k3 = XD.derive_key("different", salt, 100, "SHA512", 256, bk)
        @test k1 != k3
    end

    @testset "derive_key respects key_bits" begin
        salt = rand(UInt8, 16)
        bk   = UInt8[0x14,0x6e,0x0b,0xe7,0xab,0xac,0xd0,0xd6]
        @test length(XD.derive_key("pw", salt, 10, "SHA1", 128, bk)) == 16
        @test length(XD.derive_key("pw", salt, 10, "SHA256", 256, bk)) == 32
    end
end

@testset "get_hash_fn                      " begin
    @test XD.get_hash_fn("SHA512") === SHA.sha512
    @test XD.get_hash_fn("SHA256") === SHA.sha256
    @test XD.get_hash_fn("SHA1")   === SHA.sha1
    @test XD.get_hash_fn("SHA384") === SHA.sha384
    @test_throws ErrorException XD.get_hash_fn("MD5")
end

@testset "returned buffer is fresh/seekable" begin
    test_file = joinpath(data_directory, raw"password-is-w23$er3.xlsx")
    io = XD.decrypt_xlsx(test_file, raw"w23$er3")
    @test position(io) == 0
    bytes1 = read(io)
    seekstart(io)
    bytes2 = read(io)
    @test bytes1 == bytes2
end

@testset "non-BMP UTF-16 password (emoji)  " begin
    test_file = joinpath(data_directory, "password-is-😀🙂🚲.xlsx")
    password  = "😀🙂🚲"

    @test isfile(test_file)

    io = XD.decrypt_xlsx(test_file, password)
    @test io isa IOBuffer

    XLSX.openxlsx(io) do f
        sheet = f["Sheet1"]
        @test sheet["A1"] !== missing
    end

    @testset "wrong emoji password rejected" begin
        @test_throws ErrorException XD.decrypt_xlsx(test_file, "😀🙂🚗")
    end
end

@testset "Crypto primitives                " begin
    iv = hex2bytes("000102030405060708090a0b0c0d0e0f")
    pt = hex2bytes("6bc1bee22e409f96e93d7e117393172a")

    @testset "AES-CBC known answers (NIST SP 800-38A)" begin
        k128 = hex2bytes("2b7e151628aed2a6abf7158809cf4f3c")
        c128 = hex2bytes("7649abac8119b246cee98e9b12e9197d")
        @test XD.aes_cbc_encrypt(k128, iv, pt) == c128
        @test XD.aes_cbc_decrypt(k128, iv, c128) == pt

        k256 = hex2bytes("603deb1015ca71be2b73aef0857d77811f352c073b6108d72d9810a30914dff4")
        c256 = hex2bytes("f58c4c04d6e5f1ba779eabfb5f7bfbd6")
        @test XD.aes_cbc_encrypt(k256, iv, pt) == c256
        @test XD.aes_cbc_decrypt(k256, iv, c256) == pt
    end

    @testset "AES-CBC round trip, no padding" begin
        data = rand(UInt8, 16 * 5)
        for nbytes in (16, 24, 32)
            key = rand(UInt8, nbytes)
            ct  = XD.aes_cbc_encrypt(key, iv, data)
            @test length(ct) == length(data)
            @test XD.aes_cbc_decrypt(key, iv, ct) == data
        end
    end

    @testset "pad_to_block" begin
        aligned = rand(UInt8, 32)
        @test XD.pad_to_block(aligned, 16) == aligned
        @test XD.pad_to_block(UInt8[1, 2, 3], 16) == [UInt8[1, 2, 3]; zeros(UInt8, 13)]
        @test XD.pad_to_block(UInt8[], 16) == UInt8[]
    end

    @testset "crypt_package round trip" begin
        p    = (block_size = 16, hash_alg = "SHA512", key_data_salt = rand(UInt8, 16))
        key  = rand(UInt8, 32)
        data = rand(UInt8, 2 * XD.PACKAGE_SEGMENT_SIZE + 100)
        ct   = XD.crypt_package(key, p, data; encrypt=true)
        @test length(ct) == 2 * XD.PACKAGE_SEGMENT_SIZE + 112  # last segment padded to block
        @test XD.crypt_package(key, p, ct; encrypt=false)[1:length(data)] == data
    end

    @testset "derive_keys threaded == serial" begin
        p = (enc_key_salt = rand(UInt8, 16), spin_count = 50, enc_hash_alg = "SHA512", enc_key_bits = 256)
        @test XD.derive_keys("pw", p; threaded=true) == XD.derive_keys("pw", p; threaded=false)
    end

    @testset "derive_key pads short hashes with 0x36" begin
        bk = XD.BLOCK_KEY_VALUE
        k  = XD.derive_key("pw", rand(UInt8, 16), 10, "SHA1", 256, bk)  # SHA1 = 20 bytes < 32
        @test length(k) == 32
        @test all(==(0x36), k[21:32])
    end

    @testset "HMAC (RFC 4231 test case 2)" begin
        mac = XD.hmac("SHA256", Vector{UInt8}("Jefe"), Vector{UInt8}("what do ya want for nothing?"))
        @test mac == hex2bytes("5bdcc146bf60754e6a042426089575c75a003f089d2739839dec58b964ec3843")
        @test length(XD.hmac("SHA512", rand(UInt8, 64), rand(UInt8, 10))) == 64
        @test_throws ErrorException XD.hmac("MD5", UInt8[1], UInt8[2])
    end
end

# Excel-generated fixtures with their passwords, shared by the tests below
const FIXTURES = [
    (joinpath(data_directory, raw"password-is-w23$er3.xlsx"), raw"w23$er3"),
    (joinpath(data_directory, raw"password-is-very$long^password#3245301!.xlsx"), raw"very$long^password#3245301!"),
    (joinpath(data_directory, "password-is-😀🙂🚲.xlsx"), "😀🙂🚲"),
]

function read_fixture_streams(path)
    io = IOBuffer(read(path))
    return XD.read_cfb_stream(io, "EncryptionInfo"), XD.read_cfb_stream(io, "EncryptedPackage")
end

@testset "Data integrity (Excel fixtures)  " begin
    for (path, pw) in FIXTURES
        info, pkg = read_fixture_streams(path)
        p   = XD.parse_encryption_info(info)
        key = XD.unwrap_package_key(pw, p)

        @test p.enc_hmac_key !== nothing
        @test XD.verify_integrity(p, key, pkg)

        tampered = copy(pkg); tampered[end] ⊻= 0x01
        @test !XD.verify_integrity(p, key, tampered)

        # Re-encrypting Excel's own HMAC key must reproduce Excel's dataIntegrity
        # fields exactly — proves encrypt_integrity matches Excel bit-for-bit.
        hs       = XD.hash_size(p.hash_alg)
        hmac_key = XD.aes_cbc_decrypt(key, XD.integrity_key_iv(p), p.enc_hmac_key)[1:hs]
        @test XD.encrypt_integrity(p, key, hmac_key, pkg) == (p.enc_hmac_key, p.enc_hmac_value)

        @test_throws ErrorException XD.unwrap_package_key("wrong", p)
    end

    @testset "missing <dataIntegrity>" begin
        info, pkg = read_fixture_streams(FIXTURES[1][1])
        xml   = replace(String(info[9:end]), r"<dataIntegrity[^>]*/>" => "")
        p     = XD.parse_encryption_info([info[1:8]; Vector{UInt8}(xml)])
        @test p.enc_hmac_key === nothing && p.enc_hmac_value === nothing
        @test_throws ErrorException XD.verify_integrity(p, zeros(UInt8, 32), pkg)
    end

    @test XD.hash_size("SHA512") == 64
    @test XD.hash_size("SHA1") == 20
end

@testset "encrypt_xlsx                     " begin
    tmp = mktempdir()

    @testset "open -> encrypted -> open (new workbook written by XLSX.jl)" begin
        plain = joinpath(tmp, "plain.xlsx")
        XLSX.openxlsx(plain, mode="w") do xf
            s = xf[1]
            s["A1"] = "hello"
            s["B2"] = 42
            s["C3"] = 3.25
            s["D4"] = Date(2024, 1, 31)
        end
        enc = joinpath(tmp, "enc.xlsx")
        @test XD.encrypt_xlsx(plain, enc, "s3cret!") === nothing
        @test read(XD.decrypt_xlsx(enc, "s3cret!")) == read(plain)    # byte-identical package
        XLSX.openxlsx(XD.decrypt_xlsx(enc, "s3cret!")) do xf
            s = xf[1]
            @test s["A1"] == "hello"
            @test s["B2"] == 42
            @test s["C3"] == 3.25
            @test s["D4"] == Date(2024, 1, 31)
        end
        @test_throws ErrorException XD.decrypt_xlsx(enc, "wrong")
    end

    @testset "encrypted -> open -> encrypted -> open, on disk (Excel fixtures)" begin
        for (i, (path, pw)) in enumerate(FIXTURES)
            # open the Excel-encrypted file, edit it with XLSX.jl, save plaintext to disk
            xf = XLSX.openxlsx(XD.decrypt_xlsx(path, pw), mode="rw")
            s  = xf[1]
            a1 = s["A1"]
            s["E1"] = "edited after decrypt"
            plain = joinpath(tmp, "modified_$i.xlsx")
            XLSX.writexlsx(plain, xf, overwrite=true)

            # re-encrypt on disk with a new password and open again
            enc2 = joinpath(tmp, "reencrypted_$i.xlsx")
            XD.encrypt_xlsx(plain, enc2, pw * "-new")
            XLSX.openxlsx(XD.decrypt_xlsx(enc2, pw * "-new")) do f
                @test f[1]["E1"] == "edited after decrypt"
                @test f[1]["A1"] == a1
                if i <= 2
                    @test f["Sheet1"]["C1"] == "kjghfjvila"
                    @test f["Sheet2!C3"] == "fhlAWETYUUI"
                end
            end
            @test_throws ErrorException XD.decrypt_xlsx(enc2, pw)          # old password no longer works

            # encrypted -> encrypted directly (decrypt result fed straight back in)
            enc3 = joinpath(tmp, "reencrypted_again_$i.xlsx")
            XD.encrypt_xlsx(XD.decrypt_xlsx(enc2, pw * "-new"), enc3, pw)
            @test read(XD.decrypt_xlsx(enc3, pw)) == read(plain)
        end
    end

    @testset "output is a well-formed Agile file" begin
        out = IOBuffer()
        XD.encrypt_xlsx(read(joinpath(tmp, "plain.xlsx")), out, "pw"; spin_count=1000)
        io   = IOBuffer(take!(out))
        info = XD.read_cfb_stream(io, "EncryptionInfo")
        pkg  = XD.read_cfb_stream(io, "EncryptedPackage")
        @test info[1:8] == XD.AGILE_INFO_HEADER
        p = XD.parse_encryption_info(info)
        @test (p.cipher_alg, p.cipher_chaining, p.hash_alg, p.key_bits) == ("AES", "ChainingModeCBC", "SHA512", 256)
        @test (p.spin_count, p.enc_hash_alg, p.enc_key_bits) == (1000, "SHA512", 256)
        @test length(p.key_data_salt) == 16 && length(p.enc_key_salt) == 16
        @test XD.verify_integrity(p, XD.unwrap_package_key("pw", p), pkg)
        @test XD.build_encryption_info(p) == info
        for (cfb_path, blob) in XD.DATASPACES_STREAMS
            @test XD.read_cfb_stream(io, String(last(split(cfb_path, '/')))) == blob
        end
    end

    @testset "EncryptionInfo XML matches Excel byte-for-byte" begin
        for (path, _) in FIXTURES
            info, _ = read_fixture_streams(path)
            @test XD.build_encryption_info(XD.parse_encryption_info(info)) == info
        end
    end

    @testset "input and output forms" begin
        plain = joinpath(tmp, "plain.xlsx")
        bytes = read(plain)
        for input in (plain, bytes, IOBuffer(bytes))
            out = IOBuffer()
            XD.encrypt_xlsx(input, out, "pw"; spin_count=10)
            f = joinpath(tmp, "forms.xlsx")
            write(f, take!(out))
            @test read(XD.decrypt_xlsx(f, "pw")) == bytes
        end

        # output may overwrite the input path
        inplace = joinpath(tmp, "inplace.xlsx")
        cp(plain, inplace; force=true)
        XD.encrypt_xlsx(inplace, inplace, "pw"; spin_count=10)
        @test read(XD.decrypt_xlsx(inplace, "pw")) == bytes
    end

    @testset "fresh salts and keys every time" begin
        bytes = read(joinpath(tmp, "plain.xlsx"))
        o1, o2 = IOBuffer(), IOBuffer()
        XD.encrypt_xlsx(bytes, o1, "pw"; spin_count=10)
        XD.encrypt_xlsx(bytes, o2, "pw"; spin_count=10)
        @test take!(o1) != take!(o2)
    end

    @testset "segment boundaries" begin
        for n in (4, 16, 4096, 4097, 3 * 4096)
            data = [XD.ZIP_MAGIC; rand(UInt8, n - 4)]
            f = joinpath(tmp, "seg_$n.xlsx")
            XD.encrypt_xlsx(data, f, "pw"; spin_count=10)
            @test read(XD.decrypt_xlsx(f, "pw")) == data
        end
    end

    @testset "invalid input" begin
        @test_throws ErrorException XD.encrypt_xlsx(Vector{UInt8}("not a zip"), IOBuffer(), "pw")
        @test_throws ErrorException XD.encrypt_xlsx(UInt8[0x50, 0x4b], IOBuffer(), "pw")
        bytes = read(joinpath(tmp, "plain.xlsx"))
        @test_throws ErrorException XD.encrypt_xlsx(bytes, IOBuffer(), "pw"; spin_count=-1)
        @test_throws ErrorException XD.encrypt_xlsx(bytes, IOBuffer(), "pw"; spin_count=10_000_001)
    end
end

@testset "DataSpaces constants match Excel " begin
    for (path, _) in FIXTURES
        io = IOBuffer(read(path))
        for (cfb_path, blob) in XD.DATASPACES_STREAMS
            @test XD.read_cfb_stream(io, String(last(split(cfb_path, '/')))) == blob
        end
    end
end

cfb_bytes(streams) = (io = IOBuffer(); XD.write_cfb(io, streams); take!(io))
hdr_u32(b, off) = only(reinterpret(UInt32, b[off+1:off+4]))

@testset "CFB writer                       " begin
    @testset "round trip: sizes across the mini-stream cutoff" begin
        sizes   = [0, 1, 63, 64, 65, 511, 512, 513, 4095, 4096, 4097, 100_000]
        streams = ["S$(i)_$(n)" => rand(UInt8, n) for (i, n) in enumerate(sizes)]
        bytes   = cfb_bytes(streams)
        @test length(bytes) % 512 == 0
        @test bytes[1:8] == XD.CFB_MAGIC
        io = IOBuffer(bytes)
        for (name, data) in streams
            @test XD.read_cfb_stream(io, name) == data
        end
    end

    @testset "nested storages" begin
        streams = ["Top" => rand(UInt8, 10),
                   "A/B/Deep" => rand(UInt8, 5000),
                   "A/Mid" => rand(UInt8, 100),
                   "A/B/C/Deeper" => UInt8[1, 2, 3]]
        io = IOBuffer(cfb_bytes(streams))
        for (path, data) in streams
            @test XD.read_cfb_stream(io, String(last(split(path, '/')))) == data
        end
    end

    @testset "large stream needs DIFAT sectors (> 109 FAT sectors)" begin
        data  = rand(UInt8, 8 * 1024 * 1024)
        bytes = cfb_bytes(["Big" => data, "small" => rand(UInt8, 100)])
        @test hdr_u32(bytes, 44) > 109          # num FAT sectors
        @test hdr_u32(bytes, 72) >= 1           # num DIFAT sectors
        @test XD.read_cfb_stream(IOBuffer(bytes), "Big") == data
    end

    @testset "invalid input" begin
        @test_throws ErrorException cfb_bytes(["a" => UInt8[1], "A" => UInt8[2]])   # case-insensitive duplicate
        @test_throws ErrorException cfb_bytes(["x" * "y"^31 => UInt8[1]])           # > 31 UTF-16 units
        @test_throws ErrorException cfb_bytes(["a//b" => UInt8[1]])                  # empty name
        @test_throws ErrorException cfb_bytes(["s" => UInt8[1], "s/t" => UInt8[2]])  # stream used as storage
    end

    @testset "sibling trees are sorted, valid red-black trees" begin
        # returns black-height, checking BST order and the red-black rules on the way
        function check(entries, sid, lo, hi)
            sid == XD.NOSTREAM && return 1
            e = entries[sid + 1]
            key = XD.cfb_sort_key(e)
            @test (lo === nothing || lo < key) && (hi === nothing || key < hi)
            if e.color == XD.CFB_RED
                for c in (e.left, e.right)
                    c == XD.NOSTREAM || @test entries[c + 1].color == XD.CFB_BLACK
                end
            end
            bl = check(entries, e.left, lo, key)
            br = check(entries, e.right, key, hi)
            @test bl == br
            return bl + (e.color == XD.CFB_BLACK)
        end
        for n in 1:40
            names   = shuffle(["n$(i)" * "x"^(i % 5) for i in 1:n])
            entries = XD.cfb_assign!(XD.cfb_build_tree(["Dir/$nm" => UInt8[] for nm in names]))
            dir     = entries[2]
            @test entries[dir.child + 1].color == XD.CFB_BLACK      # tree root is black
            check(entries, dir.child, nothing, nothing)
        end
    end
end