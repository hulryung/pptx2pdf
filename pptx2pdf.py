''' This program was written by 'Professor' Daekeun Kang to make it easy for 
students to distribute materials.

This program works by opening all the pptx in a subdirectory of the running
location and saving it as a pdf.

It runs only on Windows, and PowerPoint must be installed.
'''
import win32com.client

''' convert pptx to PDF'''
def PPTtoPDF(inputFileName, outputFileName, formatType = 32):
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    powerpoint.Visible = 1
    print(inputFileName)
    print(outputFileName)

    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName + ".pdf"
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType) # formatType = 32 for ppt to pdf
    deck.Close()
    powerpoint.Quit()

''' find all pptx files in subfolders '''
def listup_files_with_subfolders():
    import os
    # full path files
    files = [os.path.join(root, file) for root, dirs, files in os.walk(os.getcwd()) for file in files]
    
    print(files)
    return files

'''select files from listed files'''
def select_files(files):
    ppt_files = []
    for file in files:
        if file[-4:] == 'pptx':
            ppt_files.append(file)
    return ppt_files

'''convert pptx to pdf'''
def pptx2pdf(ppt_files):
    for ppt_file in ppt_files:
        pdf_file = ppt_file[:-5]
        PPTtoPDF(ppt_file, pdf_file)

''' main function '''        
if __name__ == '__main__':
    files = listup_files_with_subfolders()
    ppt_files = select_files(files)
    pptx2pdf(ppt_files)
