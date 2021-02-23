import tkinter as tk

def app(propagate = False, expand = False ):
    w = tk.Tk() # New window

    tk.Label( w, text = 'Propagate: {} \nExpand: {}'.format(propagate, expand) ).grid()

    f = tk.Frame(w, width=300, height=500, bg='red') # New frame with specific size
    f.grid_propagate(propagate)
    f.grid( row=1, column=0 )

    lb = tk.Listbox(f, bg='blue') # New listbox
    if expand:
        f.columnconfigure(0, weight = 1 )
        f.rowconfigure(0, weight = 1 )

    # lb.pack(fill=tk.BOTH, expand=True)
    lb.grid( row=0, column=0, sticky = 'nsew' )
    # My guess is that grid_propagate has changed the behaviour of grid, not of pack.

    lb.insert(tk.END, 'Test 1', 'Test 2', 'Test 3')

    w.mainloop()

if __name__ == '__main__':
    #app(True, True)  # propagate and Expand
    #app(False, True)  # no propagate but expand
    #app(True, False)  # propagate without expand
    app()  # no propagate or expand