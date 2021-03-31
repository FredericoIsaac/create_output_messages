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

    pip install psycopg2
    pip install Flask
    pip install flask-sqlalchemy
    pip install docx-mailmerge
    pip install pywin32

# How to get the date of file creation

link: https://stackoverflow.com/questions/237079/how-to-get-file-creation-modification-date-times-in-python

By: 
* https://stackoverflow.com/users/779118/igracia
* https://stackoverflow.com/users/1709587/mark-amery

> Putting this all together, cross-platform code should look something like this...

    import os
    import platform
    
    def creation_date(path_to_file):
        """
        Try to get the date that a file was created, falling back to when it was
        last modified if that isn't possible.
        See http://stackoverflow.com/a/39501288/1709587 for explanation.
        """
        if platform.system() == 'Windows':
            return os.path.getctime(path_to_file)
        else:
            stat = os.stat(path_to_file)
            try:
                return stat.st_birthtime
            except AttributeError:
                # We're probably on Linux. No easy way to get creation dates here,
                # so we'll settle for when its content was last modified.
                return stat.st_mtime

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