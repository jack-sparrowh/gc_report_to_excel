class gc_to_excel(exce_path=None, groupped=True):
    
    if (excel_path != None) or (re.match(r'.xl', excel_path) == None):
        raise NameError('The extension of the file was not supplied.')
