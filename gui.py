from tkinter import *
from operator import itemgetter

def Create_New():
    import excelPython, pdb
    pdb.set_trace()
    excelPython.main()
    return 0

class ResultManager():
    
    def __init__(self, parent):
        self.myEntry = Entry(parent.LevWindow, text = 'Enter the element to be searched: ')
        self.parent = parent

    def searchInResult(self, event):
        string2Search = self.searchInResultEntry.get()
        print(string2Search)
        # remove previous uses of tag `found', if any
        self.myText.tag_remove('found', '1.0', END)
        first_index = None
        if string2Search:
            # start from the beginning (and when we come to the end, stop)
            idx = '1.0'
            while 1:
                # find next occurrence, exit loop if no more
                idx = self.myText.search(string2Search, idx, nocase=1, stopindex=END)
                
                if not idx: break
                if not first_index: first_index = idx
                # index right after the end of the occurrence
                lastidx = '%s+%dc' % (idx, len(string2Search))
                # tag the whole occurrence (start included, stop excluded)
                self.myText.tag_add('found', idx, lastidx)
                # prepare to search for next occurrence
                idx = lastidx
            # use a red foreground for all the tagged occurrences
            self.myText.tag_config('found', background='red')
        # give focus back to the Entry field

        if not(not(first_index)): self.myText.see(first_index)
        self.searchInResultEntry.focus_set(  )
        return 0
        
    
    def Search_Element(self):
        import pickle, re, pdb
        the_filename = 'matrixFile'
        with open(the_filename, 'rb') as f:
            my_list = pickle.load(f)
        mykeys = list(my_list[0].keys())
        switcher = {0:'LEVEL 1', 1:'LEVEL 2' , 2:'LEVEL 3', 3:'LEVEL 6', 4:'LEVEL 7', 5:'LEVEL 4', 6:'LEVEL 5' }
        #pdb.set_trace()
        matrix = sorted(my_list, key=itemgetter(switcher[self.parent.myvar.get()],
                    switcher[(self.parent.myvar.get()+1)%len(switcher)], 
                    switcher[(self.parent.myvar.get()+2)%len(switcher)],
                    switcher[(self.parent.myvar.get()+3)%len(switcher)],
                    switcher[(self.parent.myvar.get()+4)%len(switcher)],
                    switcher[(self.parent.myvar.get()+5)%len(switcher)],
                    switcher[(self.parent.myvar.get()+6)%len(switcher)]))
        print(self.myEntry.get())
        print(self.parent.myvar.get())
        result = list()
        resultWindow = Toplevel(self.parent.LevWindow)
        resultFrame = Frame(resultWindow)
        self.myText = Text(resultFrame, width = 160, height = 120)
        self.searchInResultEntry = Entry(resultFrame, text = 'Enter strin: ')
        for item in matrix:
            if self.myEntry.get() in item[switcher[self.parent.myvar.get()]]:
                result.append(item)
                self.myText.insert(END, item[switcher[0]] + '\n')
                self.myText.insert(END, item[switcher[1]] + '\n')
                self.myText.insert(END, item[switcher[2]] + '\n')
                self.myText.insert(END, item[switcher[3]] + '\n')
                self.myText.insert(END, item[switcher[4]] + '\n')
                self.myText.insert(END, item[switcher[5]] + '\n')
                self.myText.insert(END, item[switcher[6]])
                self.myText.insert(END, '\n\n')
                self.myText.insert(END, '-------------------------------------------------------------------------------------------------\n')
                self.myText.insert(END, '\n\n')
        self.myEntry.pack()
        myscroll = Scrollbar(resultFrame, command = self.myText.yview)
        self.myText.configure(yscrollcommand = myscroll.set)

        #PACKING
        myscroll.pack(side= RIGHT, fill=Y)
        self.myText.pack(side=LEFT)
        self.searchInResultEntry.bind('<Return>', self.searchInResult)
        self.searchInResultEntry.pack()
        self.myEntry.pack()
        resultFrame.pack()
        #resultWindow.pack()
        
class ChildWindow(Frame):

    Element = 'Insert an element'
    
    def CommandSelection(self):
        return 0

    def BindCallbackInvokeSearchButton(event, button):
        button.invoke()
    
    def __init__(self, parent):       
        self.ParentWindow = parent
        self.LevWindow = Toplevel(self.ParentWindow)
        self.myvar = IntVar(self.LevWindow)
        self.ResWin = Text(self.LevWindow)
        L1Element = Radiobutton(self.LevWindow, text = 'LEVEL 1', variable = self.myvar, value = 0, command = self.CommandSelection, anchor = NW)
        L2Element = Radiobutton(self.LevWindow, text = 'LEVEL 2', variable = self.myvar, value = 1, command = self.CommandSelection, anchor = NW)
        L3Element = Radiobutton(self.LevWindow, text = 'LEVEL 3', variable = self.myvar, value = 2, command = self.CommandSelection, anchor = NW)
        L6Element = Radiobutton(self.LevWindow, text = 'LEVEL 6', variable = self.myvar, value = 3, command = self.CommandSelection, anchor = NW)
        L7Element = Radiobutton(self.LevWindow, text = 'LEVEL 7', variable = self.myvar, value = 4, command = self.CommandSelection, anchor = NW)
        L4Element = Radiobutton(self.LevWindow, text = 'LEVEL 4', variable = self.myvar, value = 5, command = self.CommandSelection, anchor = NW)
        L5Element = Radiobutton(self.LevWindow, text = 'LEVEL 5', variable = self.myvar, value = 6, command = self.CommandSelection, anchor = NW)
        #self.myEntry = Entry(self.LevWindow, text = 'Enter the element to be searched: ')
        self.myResultManager = ResultManager(self)
        self.Element = self.myResultManager.myEntry.get()
        backButton = Button(self.LevWindow, text = 'Back', command = self.goToParentWindow)
        searchButton = Button(self.LevWindow, text = 'Start Search', command = self.myResultManager.Search_Element)
        self.myResultManager.myEntry.bind('<Return>', self.BindCallbackInvokeSearchButton(searchButton))
        L1Element.pack(anchor = NW)
        L2Element.pack(anchor = NW)
        L3Element.pack(anchor = NW)
        L6Element.pack(anchor = NW)
        L7Element.pack(anchor = NW)
        L4Element.pack(anchor = NW)
        L5Element.pack(anchor = NW)
        self.myResultManager.myEntry.pack()
        L1Element.pack()
        L2Element.pack()
        L3Element.pack()
        L4Element.pack()
        L5Element.pack()
        L1Element.select()
        searchButton.pack()
        backButton.pack()       
       
    def goToParentWindow(self):
        global top
        self.LevWindow.wm_withdraw()
        self.ParentWindow.wm_deiconify()        
        
###END OF CLASS

def Search_Element():    
    SearchWindow.LevWindow.wm_deiconify()
    top.wm_withdraw()
    return 0


top = Tk(baseName = 'Levels matching')
SearchWindow = ChildWindow(top)

newMatch_button = Button(top, text = 'Create New matches', command = Create_New)
search_el = Button(top, text = 'Search', command = Search_Element)
SearchWindow.LevWindow.wm_withdraw()
newMatch_button.pack()
search_el.pack()

top.mainloop()
