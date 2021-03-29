# Introduction

Programa que lê um ficheiro numa pasta chamada 'input_file', geralmente output de um outro, e cria um ficheiro word com os dados obtidos. 

Podendo ser transformado em pdf.

Os ficheiros que retornam são guardados na pasta 'output_file' com o número de cliente e uma descrição breve do documento.

* Exemplo:
    
    input_file:
  
        <?xml version="1.0" encoding="ISO-8859-1"?>
        <response code="200">
            <errors>
                <error></error>
            </errors>
            <totalFaturas>999</totalFaturas>
            <totalCreditos>9999.99</totalCreditos>
            <totalDebitos>0.00</totalDebitos>
            <warning>
                
            <warn>Envio de testes. Ficheiro não será considerado para processamento.</warn>
            </warning>
            <nomeFicheiro>SAFT_9999_99999_99999.resumido.xml</nomeFicheiro>
        </response>

Output file example is in output_file directory as '10101 - SAFT 02-2020.pdf'


# Modules to install


# How to get the date of file creation

link: https://stackoverflow.com/questions/237079/how-to-get-file-creation-modification-date-times-in-python

By: https://stackoverflow.com/users/3585557/steven-c-howell

In Python 3.4 and above, you can use the object oriented pathlib module interface which includes wrappers for much of the os module. Here is an example of getting the file stats.

>>> import pathlib
>>> fname = pathlib.Path('test.py')
>>> assert fname.exists(), f'No such file: {fname}'  # check that the file exists
>>> print(fname.stat())
os.stat_result(st_mode=33206, st_ino=5066549581564298, st_dev=573948050, st_nlink=1, st_uid=0, st_gid=0, st_size=413, st_atime=1523480272, st_mtime=1539787740, st_ctime=1523480272)

For more information about what os.stat_result contains, refer to the documentation. For the modification time you want fname.stat().st_mtime:

>>> import datetime
>>> mtime = datetime.datetime.fromtimestamp(fname.stat().st_mtime)
>>> print(mtime)
datetime.datetime(2018, 10, 17, 10, 49, 0, 249980)

If you want the creation time on Windows, or the most recent metadata change on Unix, you would use fname.stat().st_ctime:

>>> ctime = datetime.datetime.fromtimestamp(fname.stat().st_ctime)
>>> print(ctime)
datetime.datetime(2018, 4, 11, 16, 57, 52, 151953)

This article has more helpful info and examples for the pathlib module.

# How to transform a word file into pdf

link: https://stackoverflow.com/questions/6011115/doc-to-pdf-using-python

by: https://stackoverflow.com/users/601581/steven

A simple example using comtypes, converting a single file, input and output filenames given as commandline arguments:

import sys
import os
import comtypes.client

wdFormatPDF = 17

in_file = os.path.abspath(sys.argv[1])
out_file = os.path.abspath(sys.argv[2])

word = comtypes.client.CreateObject('Word.Application')
doc = word.Documents.Open(in_file)
doc.SaveAs(out_file, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()
You could also use pywin32, which would be the same except for:

import win32com.client
and then:

word = win32com.client.Dispatch('Word.Application')


# How to watermark a pdf

https://stackabuse.com/working-with-pdfs-in-python-adding-images-and-watermarks/
    
    import PyPDF2
    
    input_file = "example.pdf"
    output_file = "example-drafted.pdf"
    watermark_file = "draft.pdf"
  
    with open(input_file, "rb") as filehandle_input:
        # read content of the original file
        pdf = PyPDF2.PdfFileReader(filehandle_input)
        
        with open(watermark_file, "rb") as filehandle_watermark:
            # read content of the watermark
            watermark = PyPDF2.PdfFileReader(filehandle_watermark)
            
            # get first page of the original PDF
            first_page = pdf.getPage(0)
            
            # get first page of the watermark PDF
            first_page_watermark = watermark.getPage(0)
            
            # merge the two pages
            first_page.mergePage(first_page_watermark)
            
            # create a pdf writer object for the output file
            pdf_writer = PyPDF2.PdfFileWriter()
            
            # add page
            pdf_writer.addPage(first_page)
            
            with open(output_file, "wb") as filehandle_output:
                # write the watermarked file to the new file
                pdf_writer.write(filehandle_output)