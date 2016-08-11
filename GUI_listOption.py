# This serve as a module to get user input - the easy way!
# Some GUI selection
import Tkinter

default_kwargs = { 
                  'selectmode'  : "single"          ,
                  'width'       : "150"             ,
                  'height'      : "30"              ,
                  'title'       : "Choose from list",
                  'buttonText'  : "Submit"  
}



class easyListBox:

    def __init__(self, options_list, **kwargs) :

        #options
        opt = default_kwargs #default options
        opt.update(kwargs) #overrides default if existant

        #Return value
        self.selected = 0;

        # GUI master object (life-time component)
        self.master = Tkinter.Tk()

        # Checklist with options
        listbox_options = { key: opt[key] for key in opt if key in['selectmode','width','height'] } #options slice for GUI
        self.listbox = Tkinter.Listbox(self.master, listbox_options)
        self.listbox.master.title(opt['title'])

        #Options to be checked
        for option in options_list:
            self.listbox.insert(0,option)
        self.listbox.pack()

        # Submit callback
        self.OKbutton = Tkinter.Button(self.master, command=self.OKaction, text=opt['buttonText'] )
        self.OKbutton.pack()

        #Main loop
        self.master.mainloop()

    # Action to be done when the user press submit
    def OKaction(self):
        self.selected =  self.listbox.selection_get()
        self.master.destroy() 

    # Return the selection
    def getInput(self):
        return self.selected
