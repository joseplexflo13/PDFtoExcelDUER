def validate_pdf_file(file_path):
    return file_path.lower().endswith('.pdf')

def validate_excel_file(file_path):
    return file_path.lower().endswith('.xlsx')

def get_file_name_without_extension(file_path):
    import os
    return os.path.splitext(os.path.basename(file_path))[0]

def create_directory_if_not_exists(directory):
    import os
    if not os.path.exists(directory):
        os.makedirs(directory)