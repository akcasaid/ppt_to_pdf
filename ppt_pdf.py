import comtypes.client
import os
import glob
from pptx import Presentation

def powerpoint_to_pdf(input_folder, output_folder):
    powerpoint_files = glob.glob(input_folder + '/*.pptx')
    powerpoint_files.extend(glob.glob(input_folder + '/*.ppt'))

    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    for ppt_file in powerpoint_files:
        presentation = powerpoint.Presentations.Open(ppt_file)
        presentation.SaveAs(os.path.join(output_folder, os.path.splitext(os.path.basename(ppt_file))[0] + ".pdf"), 32)
        presentation.Close()

    powerpoint.Quit()

if __name__ == '__main__':
    input_folder = r"C:\Users\saidakca\Desktop\ppt"
    output_folder = r"C:\Users\saidakca\Desktop\pdf"

    powerpoint_to_pdf(input_folder, output_folder)
    powerpoint_to_pdf(input_folder, output_folder)