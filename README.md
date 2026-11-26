
# XLSXDecrypt.jl

An experimental package to allow XLSX.jl to open password encrypted Excel workbooks.

This package was written by Claude with a small amount of input from me. I do not understand encryption. My simple test case works but it has not been extensively tested.

## Requirements

* Julia v1.10 or higher

* XLSX.jl v0.10 or higher

* Linux, macOS or Windows.

## Installation

From a Julia session, run:

```julia
julia> using Pkg

julia> Pkg.add("XLSXDecrypt")
```

## Usage

This package offers two public functions:

```julia
    decrypt_xlsx(filename::String, password::String)
    encrypt_xlsx(input, output, password::AbstractString; spin_count = 100_000)
```

### Decrypting

`decrypt_xlsx`

returns an `IOBuffer` that either `XLSX.readxlsx` or `XLSX.openxlsx` can ingest.

Thus:

```julia
julia> using XLSXDecrypt, XLSX

julia> buf = decrypt_xlsx("password.xlsx", "password")

julia> f=openxlsx(buf, mode="rw")
XLSXFile("IOBuffer(data=UInt8[...], readable=true, writable=false, seekable=true, append=false, size=8554, maxsize=Inf, ptr=8555, mark=-1)") containing 1 Worksheet
            sheetname size          range
-------------------------------------------------
               Sheet1 3x1           A1:A3
```

### Encrypting

`encrypt_xlsx` password-protects an `.xlsx` file so that Excel will prompt for the
password when opening it. `input` may be a file path, a byte vector or an `IO`;
`output` may be a file path or an `IO` (and may be the same path as `input`).
Encryption uses the same scheme and settings as Excel (AES-256, SHA-512,
100,000 key-derivation iterations).

```julia
julia> using XLSXDecrypt, XLSX

julia> XLSX.writexlsx("plain.xlsx", f, overwrite=true)

julia> encrypt_xlsx("plain.xlsx", "protected.xlsx", "secret")

julia> XLSX.readxlsx(decrypt_xlsx("protected.xlsx", "secret"))
```

A workbook built with XLSX.jl can be encrypted without ever writing a plaintext
copy to disk, by saving it to an `IOBuffer` first. There is no need to rewind
the buffer: `encrypt_xlsx` always reads a seekable `IO` from the start.

```julia
julia> xf = XLSX.newxlsx()       # XLSX.jl v0.11+; or any workbook opened with mode="w" or "rw"

julia> xf[1]["A1"] = "confidential"

julia> buf = IOBuffer()

julia> XLSX.writexlsx(buf, xf)

julia> encrypt_xlsx(buf, "protected.xlsx", "secret")
```

The same works for a workbook that was itself decrypted, edited and is being
saved again, all in memory:

```julia
julia> xf = XLSX.openxlsx(decrypt_xlsx("protected.xlsx", "secret"), mode="rw")

julia> xf[1]["A2"] = "more"

julia> buf = IOBuffer(); XLSX.writexlsx(buf, xf)

julia> encrypt_xlsx(buf, "protected.xlsx", "secret")
```

A decrypted buffer can be re-encrypted directly, e.g. to change a password:

```julia
julia> encrypt_xlsx(decrypt_xlsx("protected.xlsx", "secret"), "protected.xlsx", "new secret")
```

Only the modern ECMA-376 Agile Encryption scheme (Excel 2010+) is supported.

## Source Code

The source code for this package is hosted at
[https://github.com/TimG1964/XLSXDecrypt.jl](https://github.com/TimG1964/XLSXDecrypt.jl).

## License

The source code for the package **XLSXDecrypt.jl** is licensed under
the [MIT License](https://raw.githubusercontent.com/TimG1964/XLSXDecrypt.jl/master/LICENSE).

## Getting Help

If you're having any trouble, have any questions about this package
or want to ask for a new feature,
just open a new [issue](https://github.com/TimG1964/XLSXDecrypt.jl/issues).

## Contributing

Contributions are always welcome!

To contribute, fork the project from [GitHub](https://github.com/TimG1964/XLSXDecrypt.jl)
and send a Pull Request.
