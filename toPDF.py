import os
import multiprocessing
import numpy as np

path = ''

def WORDtoPDF(in_file):
    if in_file.split('.')[-1] in ['doc', 'docx']:
        os.system(
            f'unoconv --connection \'socket,host=127.0.0.1,port=2220,tcpNoDelay=1;urp;StarOffice.ComponentContext\' -f pdf {path}{in_file}')
        os.system(f'rm {path}{in_file}')


func_vect = np.vectorize(WORDtoPDF, otypes=[str])
process = []


def multi_process(file_name):
    l = len(file_name)
    div = l//4
    if div != 0:
        for i in range(4):
            p = multiprocessing.Process(target=func_vect, args=[
                                        file_name[i*div:(i+1)*div]])
            p.start()
            process.append(p)
        if (l-4*div) != 0:
            multi_process(file_name[4*div:l])
    else:
        p = multiprocessing.Process(target=func_vect, args=[file_name])
        p.start()
        process.append(p)


def main():
    file_names = np.array(os.listdir(path))
    multi_process(file_names)


if __name__ == "__main__":
    main()
