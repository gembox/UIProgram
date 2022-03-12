from importlib.resources import path
from sre_parse import State
from sysconfig import get_path
from tkinter import BOTTOM, NORMAL, DISABLED, LEFT, RIGHT, Button, Frame, StringVar, Tk, BOTH, Text, Menu, END
from tkinter import filedialog
import tkinter
from tkinter.constants import DISABLED, NORMAL
from tkinter.ttk import Label
from turtle import begin_fill, width
from xml.dom.pulldom import START_DOCUMENT
from lxml import etree
from spellchecker import SpellChecker
import re

def spellcheck_workbook(wbname,path='',new_text=''):
    '''
    This will take in a workbook and check titles and text box objects for spelling errors. If errors are found
    it will tell you the text string, which word was flagged and where it was located. It is currently avoiding
    processing tooltips because of their messy representation in the XML, but may be implemented later.
    '''
    #Read workbook xml and instatiate spellchecker
    tree = etree.parse(wbname)
    troot = tree.getroot()
    spell = SpellChecker()
    etree.register_namespace('user', "http://www.tableausoftware.com/xml/user")

    def update_element(path,new_text):
        print(path)
        for child in tree.xpath(path):
         if child is not None:
            child.text = new_text
            et = etree.ElementTree(troot)
            et.write(wbname, xml_declaration=True, encoding ='utf-8')
    
    def find_errors(objecttype):
        '''
        This will take in a object type to find the element, parent path to get attributes from, and checks their 
        spelling for errors. Right now it supports dashboards or worksheets. In the future this could be expanded 
        to include additional types of objects/paths.
        '''
        slabel = ''
        swords = ''
        twords = ()
        rwords = ''
        trwords = ()
        stext = ''
        stuple = ()
        textpath = ''
        matches = []
        elempath = './/dashboard/zones//formatted-text/run' if objecttype == 'textbox' else './/worksheet//title/formatted-text/run'
        elemparent = 'ancestor::dashboard' if objecttype == 'textbox' else 'ancestor::worksheet'
        for w in troot.findall(elempath):
            itempath = tree.getpath(w)
            elemtype = 'dashboard' if objecttype == 'textbox' else 'worksheet'
            words = re.sub(r'[^\w\s]','',w.text)
            misspelled_ws = spell.unknown(words.split())
            if len(misspelled_ws) > 0:
                stext = w.text
                for word in misspelled_ws:
                    swords = 'Flagged word: ' + word
                    twords = twords + (swords,)
                    rwords = 'Suggested replacement: ' + spell.correction(word)
                    trwords = trwords + (rwords,)
                for parent in w.xpath(elemparent):
                    slabel = 'Found in {}: '.format(elemtype) + parent.attrib['name']
                stuple = (itempath,slabel,stext,twords,trwords)
                matches.append(stuple)
                twords = ()
                trwords = ()
        return matches

    #make updates for values passed in
    if len(path) > 0:
        update_element(path,new_text)
    
    #let's call find_errors for worksheets and dashboards
    spelling = find_errors('title')
    spelling += find_errors('textbox')
    return spelling

class SpellCheckUI(Frame):

    def __init__(self, parent):
        Frame.__init__(self, parent)   
        self.parent = parent        
        self.initUI()
    
    def initUI(self):
        self.parent.title("File dialog")
        self.pack(fill=BOTH, expand=1)

        menubar = Menu(self.parent)
        self.parent.config(menu=menubar)
        
        fileMenu = Menu(menubar)
        fileMenu.add_command(label="Open", command=self.onOpen)
        fileMenu.add_command(label="Exit", command=self.quit)
        menubar.add_cascade(label="File", menu=fileMenu)        

        Label(self, text='Open a .twb file to check the spelling of titles and textboxes.', font=('Helvetica bold',10)).pack(pady=5)
        Label(self, text='To update a title, change the text in the box and press the Change button. To leave it as is press Ignore.', font=('Helvetica bold',10)).pack(pady=5)
        self.lbltext = StringVar()
        self.sourcelabel = Label(self, textvariable=self.lbltext, font=('Helvetica bold',10)).pack()
        self.txt = Text(self,height=4,width=50)
        self.txt.pack(expand=True)
        self.errorstext = StringVar()
        self.errorslabel = Label(self, textvariable=self.errorstext, font=('Helvetica bold',10)).pack()
        self.suggestiontext = StringVar()
        self.suggestionlabel = Label(self, textvariable=self.suggestiontext, font=('Helvetica bold',10)).pack()
        self.changeButton= tkinter.Button(self, text="Change", command=self.cycle_next, font=('Helvetica bold',10), state="disabled")
        self.changeButton.pack(side=RIGHT,padx=40)
        self.ignoreButton = tkinter.Button(self, text="Ignore", command=self.skip_text, font=('Helvetica bold',10), state="disabled")
        self.ignoreButton.pack(side=RIGHT,padx=5)

    def onOpen(self):
        ftypes = [('Tableau workbook files', '*.twb')]
        dlg = filedialog.Open(self, filetypes = ftypes)
        self.fl = dlg.show()

        if self.fl != '':
            output = spellcheck_workbook(self.fl)
            if len(output) > 0:
                self.ignoreButton.config(state=NORMAL)
                self.changeButton.config(state=NORMAL)
                self.txt.insert(END, output[0][2])
                self.lbltext.set('{}'.format(output[0][1]))
                self.errorstext.set('{}'.format(output[0][3]))
                self.suggestiontext.set('{}'.format(output[0][4]))
                self.outputreturn = output
                self.activeelement = self.outputreturn.pop(0)
            else:
                self.lbltext.set('No spelling errors found.')
    
    def cycle_next(self):
        result=self.txt.get("1.0","end").strip()
        spellcheck_workbook(self.fl,self.activeelement[0],result)
        i = 0
        if i < len(self.outputreturn):
            self.txt.delete("1.0","end")
            self.txt.insert(END, self.outputreturn[0][2])
            self.lbltext.set('{}'.format(self.outputreturn[0][1]))
            self.errorstext.set('{}'.format(self.outputreturn[0][3]))
            self.suggestiontext.set('{}'.format(self.outputreturn[0][4]))
            self.activeelement = self.outputreturn.pop(0)
        else:
            self.lbltext.set("End of list.")
            self.txt.delete("1.0","end")
            self.errorstext.set('')
            self.suggestiontext.set('')
            self.ignoreButton.config(state=DISABLED)
            self.changeButton.config(state=DISABLED)
    
    def skip_text(self):
        i = 0
        if i < len(self.outputreturn):
            self.txt.delete("1.0","end")
            self.txt.insert(END, self.outputreturn[0][2])
            self.lbltext.set('{}'.format(self.outputreturn[0][1]))
            self.errorstext.set('{}'.format(self.outputreturn[0][3]))
            self.suggestiontext.set('{}'.format(self.outputreturn[0][4]))
            self.outputreturn.pop(0)
        else:
            self.lbltext.set("End of list.")
            self.txt.delete("1.0","end")
            self.errorstext.set('')
            self.suggestiontext.set('')
            self.ignoreButton.config(state=DISABLED)
            self.changeButton.config(state=DISABLED)

def main():
    root = Tk()
    ex = SpellCheckUI(root)
    root.title('.twb Spellchecker')
    root.geometry('600x300+50+50')
    #root.iconbitmap('./Bars.ico')
    root.mainloop()

if __name__ == '__main__':
    main()

