import os
import psycopg2
import datetime
from mailmerge import MailMerge
from win32com import client

# ---------------------------------------- CONSTANT VARIABLES ----------------------------------- #
INPUT_PATH = 'input_file'
OUTPUT_PATH = 'output_file'
TEMPLATE_PATH = 'templates/template_safts.docx'
DB_PASS = os.environ['DB_PASS']


# ---------------------------------------- EXTRACTOR -------------------------------------------- #
def query_db(company_id, extractor_type):
    """
    Query data base for specific information
    :param company_id: int
    :param extractor_type: string
    :return: Tuple with the info required for the specific extractor type
    """
    result = None
    conn = psycopg2.connect(host='localhost', database='confere', user='postgres', password=DB_PASS)
    cur = conn.cursor()

    # Extract nif and company name:
    if extractor_type == 'saft':
        cur.execute(
            'SELECT company, nif FROM companies WHERE client_id = (%s);', (company_id,)
        )
        result = cur.fetchone()
    elif extractor_type == 'mail':
        # Not needed for now.
        pass

    conn.close()
    return result


# ---------------------------------------- POPULATE WORD -------------------------------------------- #
def saft_to_word(template_path, output_path, *args):
    """
    Populate the word Document and save to a new file
    :param output_path: path to the output directory
    :param template_path: path to the word template
    :param args: In order of the fields in the word template
    :return: The path of the saved file
    """
    document = MailMerge(template_path)

    # Get the name of the fields in Word:
    # for field in document.get_merge_fields():
    #     print(field)

    # Get the args order:
    # for arg in args:
    #     print(arg)

    document.merge(
        cliente=args[0] + ' - ' + args[1],
        nif=args[2],
        ano=args[3],
        mes=args[4],
        data=args[5],
        n_fat=args[6],
        t_creditos=args[7],
        t_debitos=args[8],
    )

    file_name = f'{args[0]} - SAFT {args[4]}-{args[3]}.docx'

    # Handle multiple outputs of the same company
    list_output_dir = os.listdir(output_path)
    string_output = ' '.join(list_output_dir)
    already_there = string_output.count(file_name)

    if already_there != 0:
        file_name = f'{args[0]} - SAFT Loja {str(already_there + 1)} {args[4]}-{args[3]}.docx'

    # Save in output path word document
    full_output_path = os.path.join(output_path, file_name)
    document.write(full_output_path)

    return full_output_path


# ---------------------------------------- TRANSFORM TO PDF -------------------------------------------- #
def word_to_pdf(doc_to_transform):
    # Get current directory
    full_directory = os.path.abspath(os.getcwd())
    # Join current with the word path to transform
    full_path_doc = os.path.join(full_directory, doc_to_transform)
    # Create the path of the new file pdf with the same name of the word
    full_output_path = os.path.join(full_directory, doc_to_transform[:-5])

    word_format_pdf = 17

    word = client.Dispatch('Word.Application')
    doc = word.Documents.Open(full_path_doc)
    doc.SaveAs(full_output_path, FileFormat=word_format_pdf)
    doc.Close()
    word.Quit()

    # Delete Word that has already transform to pdf
    os.remove(full_path_doc)

    return full_output_path


# ------------------------------------------------------------------------------------------------------------------ #
if __name__ == '__main__':
    # What params do you want to save
    type_of_files = input('What type of files do you want to transform? SAFT\'s (S) or Mail (M) ').lower()

    # List all the files in input file
    input_files = os.listdir(INPUT_PATH)

    for file in input_files:
        full_path = os.path.join(INPUT_PATH, file)
        with open(full_path) as f:
            text = f.read()

        if type_of_files == 's':
            # The beginning of the file has the id and the end the date
            id_company = file[:5]
            month, year = file[-7:].split('-')

            name_company, nif = query_db(int(id_company), 'saft')

            # Get the date of creation and transform to string '2020-02-28 14:30'
            file_create_date = datetime.datetime.fromtimestamp(os.path.getctime(full_path)).strftime('%Y-%m-%d %H:%M')
            # Search in the file
            t_invoice = text[text.find('<totalFaturas>') + len('<totalFaturas>'): text.find('</totalFaturas>')]
            t_credit = text[text.find('<totalCreditos>') + len('<totalCreditos>'): text.find('</totalCreditos>')]
            t_debit = text[text.find('<totalDebitos>') + len('<totalDebitos>'): text.find('</totalDebitos>')]

            # Populate word
            path_word = saft_to_word(TEMPLATE_PATH, OUTPUT_PATH, id_company, name_company, str(nif), year, month,
                                     file_create_date, t_invoice, t_credit, t_debit)

            print(
                f'Document Successfully saved in "output_file" with the name {path_word}.docx\nTransforming to pdf...')

            # Transform to pdf
            path_pdf = word_to_pdf(path_word)
            print(f'Successfully create pdf at {path_pdf}.pdf')
            print('-' * 200)


        elif type_of_files == 'm':
            # Not needed for now.
            print('Program is not ready yet to transform to mail')
            print('Exiting program...')
            exit()
        else:
            print('Wrong command, exiting the program...')
            exit()

    print('End Program...')
