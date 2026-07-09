import XLSX
import XLSXDecrypt as XD
using Test
using Dates
using SHA

data_directory = joinpath(dirname(pathof(XD)), "..", "data")

@assert isdir(data_directory)

@testset "Basic functionality" begin
    test_file = joinpath(data_directory, raw"password-is-w23$er3.xlsx")
    io=XD.decrypt_xlsx(test_file, raw"w23$er3")
    test_file2 = joinpath(data_directory, raw"password-is-very$long^password#3245301!.xlsx")
    io2=XD.decrypt_xlsx(test_file2, raw"very$long^password#3245301!")

    @testset "number formats" begin
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

    @testset "Defined Names" begin

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

    @testset "Newer functionality" begin

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

@testset "Error handling" begin
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

@testset "Internal helpers" begin
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

@testset "get_hash_fn" begin
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

@testset "non-BMP UTF-16 password (emoji)" begin
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