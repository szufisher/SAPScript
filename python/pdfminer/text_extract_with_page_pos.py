# -*- coding: utf-8 -*-
from pdfminer.pdfdocument import PDFDocument, PDFNoOutlines
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTPage, LTChar, LTAnno, LAParams, LTTextBox, LTTextLine
from pprint import pprint
from pdfminer.pdfpage import PDFPage
import operator


class PDFPageDetailedAggregator(PDFPageAggregator):
    def __init__(self, rsrcmgr, pageno=1, laparams=None):
        PDFPageAggregator.__init__(self, rsrcmgr, pageno=pageno, laparams=laparams)
        self.rows = []
        self.page_number = 0
    def receive_layout(self, ltpage):        
        def render(item, page_number):
            if isinstance(item, LTPage) or isinstance(item, LTTextBox):
                for child in item:
                    render(child, page_number)
            elif isinstance(item, LTTextLine):
                child_str = ''
                for child in item:
                    if isinstance(child, (LTChar, LTAnno)):
                        child_str += child.get_text()
                child_str = ' '.join(child_str.split()).strip()
                if child_str:
                    row = (page_number, item.bbox[0], item.bbox[1], item.bbox[2], item.bbox[3], child_str) # bbox == (x1, y1, x2, y2)
                    self.rows.append(row)
                for child in item:
                    render(child, page_number)
            return
        render(ltpage, self.page_number)
        self.page_number += 1
        self.rows = sorted(self.rows, key = lambda x: (x[0], -x[2], x[1]))   #sort by page, y, x         
        self.result = ltpage

def main(file):
    fp = open(file, 'rb')
    parser = PDFParser(fp)
    doc = PDFDocument(parser)
    #doc.initialize() # leave empty for no password

    rsrcmgr = PDFResourceManager()
    laparams = LAParams()
    device = PDFPageDetailedAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)

    for page in PDFPage.create_pages(doc):
        interpreter.process_page(page)
        # receive the LTPage object for this page
        device.get_result()

    pprint(device.rows)
    
    extract_data(device.rows)
    
def extract_data(rows):
    header_fields = ['Order number 订单号', 'Order date 订单日期']
    item_fields=['描述','币别']
    table_end_row_indicators = ['Company Address:','Total value of the order without VAT']
    header = {}
    item=[]
    cur_page = -1
    new_page = False
    row_count= -1
    for idx, val in enumerate(rows):
        if val[0] != cur_page:
            new_page = True
            item_row_started = False
            item_fields=['描述','币别']
            cur_page = val[0]
        value = val[-1]
        if header_fields and value in header_fields:
            header.update({value: rows[idx+1][-1]})
            header_fields.remove(value)
        if item_fields and value == item_fields[0] and rows[idx+1][-1] == item_fields[1]:
            item_fields = []
            item_row_started = True
            item_start_row = idx + 3
            cur_row_idx = -1
            
        if item_row_started:                        
            if  value in table_end_row_indicators:  #with only table header column, no content rows on 1st page
                item_row_started = False                
            elif idx - item_start_row >=0:
                if len(value) == 5 and value.isnumeric():  #if len(value) == 5 and value.isnumeric():                                
                    item.append([value])                    
                    row_count += 1
                    col_num=0
                else:
                    col_num += 1
                    if col_num < 6 or not all(x in value for x in ['Page','of']): # remove the trailing page counter
                        item[row_count].append(value)            
                    
    print(header)
    pprint(item)
    
        
if __name__ == "__main__":
    for f in ['sample.pdf','po1.pdf','po2.pdf','po3.pdf','po4.pdf','po5.pdf']: #
        main(f)    
