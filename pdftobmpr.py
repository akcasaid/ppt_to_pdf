import os

def change_extension(file_name, new_extension):
    base = os.path.splitext(file_name)[0]
    os.rename(file_name, base + new_extension)


change_extension('dosyaismi.pdf', '.bmpr')
