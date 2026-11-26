# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [unreleased]

- Drop support for Julia v1.9 (its binaries no longer run on current Ubuntu CI runners); minimum is now v1.10 (LTS)
- Move test fixtures into `test/data`

## [v0.1.1](https://github.com/TimG1964/XLSXDecrypt.jl/tree/v0.1.0) - 2026-07-09

- Add support for XML v0.4 and XLSX v0.12
- Extend tests
- Separate functional elements into separate files
- Add dependabot support

## [v0.1.0](https://github.com/TimG1964/XLSXDecrypt.jl/tree/v0.1.0) - 2026-03-14

This is the first experimental version of the package and was largely written by Claude.

- Supports decryption of password-encrypted .xlsx Excel files to a plain text IO Buffer.
