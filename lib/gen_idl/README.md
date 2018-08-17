# GenIdl
VbCairo Project Typelib Generator

### Description

This folder contains a C/C++ headers PEG parser that extracts function declarations along with enums, structs and typedes from Cairo Graphics C API sources.

### Compilation

You'll need [VbPeg.exe](https://github.com/wqweto/VbPeg) parser generator to compile this parser's grammar -- just use `compile_parser.bat` to compile `gen_idl.peg` to `mdParser.bas`.

Once the parser `.bas` file is generated you can compile `gen_idl.vbp` project in VB6 IDE as usual.

### Usage

This tool is used by `build.bat` in `src/typelib` folder to compile `VbCairo.idl` to `bin/typelib`.

    GenIdl 0.1 (c) 2018 by the community at vbforums.com (17.8.2018 13:58:39)

    Usage: gen_idl.exe [options] <include_files> <include_dirs> ...

    Options:
      -o OUTFILE      output .idl file [default: stdout]
      -def DEFFILE    input .def file
      -json           dump includes parser result
      -types          dump types in use
