#!/usr/bin/env python3
#import camelot
#import PyPDF2
import sys

from pdfminer.pdfparser import PDFParser
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBox,LTChar, LTFigure, LTRect
import sys

class PdfMinerWrapper(object):
    """
    Usage:
    with PdfMinerWrapper('2009t.pdf') as doc:
        for page in doc:
            #do something with the page
    """
    def __init__(self, pdf_doc, pdf_pwd=""):
        self.pdf_doc = pdf_doc
        self.pdf_pwd = pdf_pwd
    def __enter__(self):
        #open the pdf file
        self.fp = open(self.pdf_doc, 'rb')
        # create a parser object associated with the file object
        parser = PDFParser(self.fp)
        # create a PDFDocument object that stores the document structure
        doc = PDFDocument(parser, password=self.pdf_pwd)
        # connect the parser and document objects
        parser.set_document(doc)
        self.doc=doc
        return self
    
    def _parse_pages(self):
        rsrcmgr = PDFResourceManager()
        laparams = LAParams(char_margin=3.5, all_texts = True)
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)
    
        for page in PDFPage.create_pages(self.doc):
            interpreter.process_page(page)
            # receive the LTPage object for this page
            layout = device.get_result()
            # layout is an LTPage object which may contain child objects like LTTextBox, LTFigure, LTImage, etc.
            yield layout
    def __iter__(self): 
        return iter(self._parse_pages())
    
    def __exit__(self, _type, value, traceback):
        self.fp.close()



class ExcelPDFParser(object):
    
    def __init__(self):
        """
        """
        self.lineTolerance = 2
        self.stripCells = True
        self.returnNocells = False
        

    def overlapY(self,bbox, rowy):
        if bbox[1] <= rowy and bbox[3] >= rowy:
            return True
        return False 

    def prepareCellsForLine(self, vlines, rowy):
        cells = []
        
        for line in vlines:
            if self.overlapY(line,rowy):
                if len(cells) == 0:
                    cells.append([line[0],0])
                else:
                    cells[-1][1] = line[0]
                    cells[-1].reverse()
                    cells.append([line[0],0])
                #endif
            #endif
        #endfor
        if len(cells) != 0:
            cells = cells[:-1]
            cells.reverse()
        else:
            if self.returnNocells:
                cells.append([0,self.pageWidth])
            else:
                return None

        #print(cells)
        return {'rowy': rowy,
                'cells':cells,
                 'cellsData':[''] * len(cells)}

    def computeCellIndex(self, cells, x):
        idx = 0
        for cell in cells:
            if cell[1] >= x and cell[0] <= x:
                #have it
                return idx
            idx += 1
        return -1
    
    def processTextToCells(self,vlines,tboxes):
        #sort vlines and tboxex by y2
        vlines.sort(key = lambda x: x[0], reverse=True)
        tboxes.sort(key = lambda x: x.bbox[3], reverse=True)
        for tbox in tboxes:
            rowy = tbox.bbox[3]
            if len(self.data) == 0 or abs(self.data[-1]['rowy'] - rowy) >  self.lineTolerance:
                if len(self.data) != 0 and self.stripCells:
                    #strip last
                    #print("%s " % (self.data[-1]['cellsData'],))
                    self.data[-1]['cellsData'] = list(map(str.strip,self.data[-1]['cellsData']))
                    #print("%s " % (self.data[-1]['cellsData'],)) 
                newRow = self.prepareCellsForLine(vlines,rowy)
                if newRow:
                    self.data.append(newRow) 
                else:
                    continue

            #endif
            #now parse char by char and append it to right cell in this row
            for obj in tbox:
                #print (' '*2, obj.get_text()[:-1], '(%0.2f, %0.2f, %0.2f, %0.2f)'% tbox.bbox)
                for c in obj:
                    if not isinstance(c, LTChar):
                        continue
                    idx = self.computeCellIndex(self.data[-1]['cells'],c.bbox[0])
                    #print(idx)
                    if idx != -1:
                        self.data[-1]['cellsData'][idx] += c.get_text()
                    #print (c.get_text().encode('UTF-8'), '(%0.2f, %0.2f, %0.2f, %0.2f)'% c.bbox, c.fontname, c.size,)
                
        if len(self.data) != 0 and self.stripCells:
            self.data[-1]['cellsData'] = list(map(str.strip,self.data[-1]['cellsData']))   
        #endfor


    def parse(self, filename):
        self.data = []
        self.info = []
        with PdfMinerWrapper(filename) as doc:
            self.info = doc.doc.info
            for page in doc:     
                #print ('Page no.', page.pageid, 'Size',  (page.height, page.width) ) 
                self.pageWidth = page.width
                self.pageHeight = page.height
                vlines = [] 
                tboxes = []
                for tbox in page:
                    #print (tbox)
                    if isinstance(tbox,LTRect):
                        #is horizontal or vertical ?
                        dx = tbox.bbox[2] - tbox.bbox[0]
                        dy = tbox.bbox[3] - tbox.bbox[1]
                        if dx < dy:
                            #print("Vertical Box %s lenght %s" % (str(tbox.bbox),dy))
                            vlines.append(tbox.bbox)

                    if not isinstance(tbox, LTTextBox):
                        continue
                    #print (' '*1, 'Block', 'bbox=(%0.2f, %0.2f, %0.2f, %0.2f)'% tbox.bbox)
                    tboxes.append(tbox)
                #process page text boxes    
                self.processTextToCells(vlines,tboxes)
                #break   
        return self.info,self.data





def main():
    p = ExcelPDFParser()
    info,data = p.parse(sys.argv[1]) 
    print("Author %s" % str(info[0]["Author"]))    
    #print(info)
    for line in data:
        print(line['cellsData'])
if __name__=='__main__':
    main()


