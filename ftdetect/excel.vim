" Filetype detection for excel.nvim

augroup filetypedetect
  " Excel files
  autocmd BufNewFile,BufRead *.xlsx setfiletype excel
  autocmd BufNewFile,BufRead *.xls setfiletype excel
  autocmd BufNewFile,BufRead *.xlsm setfiletype excel
  autocmd BufNewFile,BufRead *.xlsb setfiletype excel
  autocmd BufNewFile,BufRead *.csv setfiletype excel
augroup END
