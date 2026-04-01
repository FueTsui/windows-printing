# windows-printing

An OpenClaw skill for Windows local printing, including printer listing and configurable printing options for local files.

## What It Does

`windows-printing` helps an assistant work with local Windows printing workflows.

It supports tasks such as:
- listing available printers
- choosing a target printer
- printing local files
- configuring practical print options

## Repository Structure

```text
SKILL.md
references/
  options.md
scripts/
  list_printers.ps1
  print_file.ps1
  resolve_file.ps1
```

## Main Files

- `SKILL.md` — skill instructions and behavior
- `references/options.md` — supported printing options and notes
- `scripts/list_printers.ps1` — list local printers
- `scripts/print_file.ps1` — send a local file to print
- `scripts/resolve_file.ps1` — resolve file selection or helper logic for printing flows

## Usage

Use this skill on a Windows machine where local printer access is available.

Typical workflow:
1. list printers
2. select a target printer
3. choose print options
4. print the local file

## Notes

This skill is intended for local Windows environments and depends on the available printers, Windows printing stack, and local file access.
