" Syntax highlighting for excel.nvim
" Excel spreadsheet view

if exists("b:current_syntax")
  finish
endif

" Headers
syntax match excelHeader /^   |.*/ contains=excelColumnHeader,excelSeparator
syntax match excelColumnHeader /\s[A-Z]\+\s/ contained
syntax match excelSeparator /|/ contained

" Row numbers
syntax match excelRowNumber /^\s*\d\+|/

" Cell separators
syntax match excelCellSeparator /|/

" Formulas (cells starting with =)
syntax match excelFormula /=\w\+([^)]*)/

" Numbers
syntax match excelNumber /\s-\?\d\+\(\.\d\+\)\?\s/

" Strings
syntax region excelString start=/"/ end=/"/ oneline

" Highlighting
highlight default link excelHeader Comment
highlight default link excelColumnHeader Identifier
highlight default link excelRowNumber LineNr
highlight default link excelSeparator Comment
highlight default link excelCellSeparator Comment
highlight default link excelFormula Function
highlight default link excelNumber Number
highlight default link excelString String

let b:current_syntax = "excel"
