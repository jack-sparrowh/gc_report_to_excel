class gc_to_excel(exce_path=None, excel_name=None, groupped=True):
    
    if (excel_path != None) and (re.match(r'.xl', excel_path) == None) and (excel_name == None):
        raise NameError('The extension of the file was not supplied.')