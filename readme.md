# Python Excel COM Test

## Introduction

Potentially useful attempt at Excel-Python interoperability out of curiosity and as a challenge.

The main focus of this project is to tie `Excel.Range` and custom 2-dimensional `Array` (`list`s of `list`s) objects together so the data can be updated in realtime on an open Excel workbook, whilst maintaining data accessibility.

I'm working on another project to automate Minitab analysis, so this could potentially get `COM`s (Component Object Model) of Excel and Minitab Statistical Software working together.

Relatively simple to implement benefits of this workflow:
- 2D array data visibility comparable to MATLAB
- Statistics utilities of Minitab
- Easily writeable logic
- Custom plot generation using either Python or Excel's more familiar formatting tools for my line of work.

## Considerations

It may be worth changing from a custom `list` of `list`s object basis to support different data types (e.g. `numpy` `dtype`s). This would limit input type flexibility in the case of `numpy`, however the structure was designed to work with `pandas.Dataframe` to perform filtering logic and be easily fed back in to the `Array` object.