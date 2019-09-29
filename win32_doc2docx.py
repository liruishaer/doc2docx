from win32com import client as wc
import time

def doc2docx(doc_path,docx_path):
    start = time.time()
    # word = wc.DispatchEx("Word.Application")
    word = wc.Dispatch('word.Application')
    doc = word.Documents.Open(doc_path)
    doc.SaveAs(docx_path,16)  #16 doc2docx
    doc.Close()
    word.Quit()

    end = time.time()
    use_time = end-start
    print('used time:',use_time)


doc_path = 'C:\Users\LiRui\Desktop\doc2docx\doc_dir\error.doc'
docx_path = 'C:\Users\LiRui\Desktop\doc2docx\docx_dir'
doc2docx(doc_path,docx_path)
