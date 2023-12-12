# convert rtf to docx and embed all pictures in the final document
def ConvertRtfToDocx(rootDir, file):
    print("Conversion of rtf to docx started")
    word = win32com.client.Dispatch("Word.Application")
    wdFormatDocumentDefault = 16
    wdHeaderFooterPrimary = 1
    doc = word.Documents.Open(rootDir + "\\" + file)
    doc.SaveAs(str(rootDir + "\\test.docx"), FileFormat=wdFormatDocumentDefault)
    doc.Close()
    word.Quit()
    print("Conversion of rtf to docx completed. Please check this directory %s"%rootDir)

ConvertRtfToDocx("C:\\Users\\Documents", "test.rtf")

