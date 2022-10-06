


import win32com.client

import pandas as pd


xlleft = -4159
xlup = -4162
xldown = -4121
xlright = -4161


xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = True

wb = xl.Workbooks("test.xlsx")
wb.Sheets(1).Activate()


class Array:
    def __init__(self, xlRange):
        self.xlRange = xlRange

        self.set_rows()

        self.set_size()
        self.set_position()

    def __repr__(self):
        return "<Array object:{0},{1},{2}>".format(
            self.xlRange.Parent.Parent.Name,
            self.xlRange.Parent.Name,
            self.pos,
            )
    
    def __str__(self):
        return self.Rows.__str__()

    def clear(self):
        # clears cell contents to not leave history
        self.xlRange.Cells.Clear()

    def reset(self):
        # use reset if cells are overwritten by mistake for example
        self.set(self.Rows)

    def set_position(self):
        self.pos = (self.xlRange.Row, self.xlRange.Column)

    def set(self, array, axis=0):
        # set the range values (changes size if required)
        self.clear()
        if axis==0:
            rows = len(array)
            cols = len(array[0])
        elif axis==1:
            cols = len(array)
            rows = len(array[0])

        self.xlRange = Range(self.pos[0], self.pos[1], rows, cols,
            scope=self.xlRange.Parent)

        self.set_values(array, axis)

    def set_size(self):
        self.size = (self.xlRange.Rows.Count, self.xlRange.Columns.Count)

    def set_rows(self):
        vals = self.xlRange.Cells.Value
        if type(vals) is not tuple:  # convert single cells into tuples
            vals = ((vals,),)
        self.Rows = [[vals[j][i] for i in range(len(vals[j]))] for j in range(len(vals))]
        self.set_cols()


    def set_cols(self):
        self.Cols = list(map(list, zip(*self.Rows)))


    def to_df(self):
        return pd.DataFrame(self.Rows)

    def set_values(self, array, axis=0):
        # axis=0 [[row],[row],...]
        # axis=1 [[col],[col],...]
        if axis == 0:
            for i in range(len(array)): # rows
                for j in range(len(array[i])): # cols
                    self.xlRange.Cells(i+1,j+1).Value = array[i][j]
        elif axis == 1:
            for i in range(len(array)): # cols
                for j in range(len(array[i])): # rows
                    self.xlRange.Cells(j+1,i+1).Value = array[i][j]
        self.set_rows()
    
    def Select(self):
        self.xlRange.Select()


def Range(a=1, b=1, row_off=None, col_off=None, scope=xl):
    if (row_off==None) or (col_off==None):
        return absRange(a,b)
    else:
        return relRange(a, b, row_off, col_off, scope)

def relRange(a, b, row_off, col_off, scope=xl):
    return absRange(a, b, a+row_off-1, b+col_off-1, scope)

def absRange(a, b, c=None, d=None, scope=xl):
    if (c is None) or (d is None):
        return scope.Range(Cell(a, b, scope), Cell(a, b, scope))
    return scope.Range(Cell(a, b, scope), Cell(c, d, scope))


def Cell(a, b, scope=xl):
    return scope.Cells(a, b)

def transpose(array):
    return list(map(list, zip(*array)))


a = Array(Range(3,2,10,1))
a.Select()

d = Array(Range(3,4,2,3))

e = Array(Range())

print("done")

