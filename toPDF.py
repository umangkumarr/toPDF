import sys
import os
import comtypes.client
import multiprocessing

def WORDtoPDF(in_file, out_file, wdFormatPDF=17):
    if in_file.split('.')[-1]=='doc' or in_file.split('.')[-1]=='docx':
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(in_file)
        doc.SaveAs(out_file, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()
        os.remove(in_file)

def main():
    file_names=list(list(os.walk("docs"))[0])

    processes=[]

    for _ in range(len(file_names[2])):
        p = multiprocessing.Process(target=WORDtoPDF, args=[os.path.abspath(file_names[0])+"/"+file_names[2][_],os.path.abspath(file_names[0])+"/"+"".join(file_names[2][_].split(".")[:-1])+".pdf"])
        p.start()
        processes.append(p)

    for process in processes:
        process.join()

if __name__=="__main__":
    main()