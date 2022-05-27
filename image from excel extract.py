# -*- coding: utf-8 -*-
"""
Created on Wed May 11 11:48:15 2022

@author: FA10041
"""
import pandas as pd
from PIL import ImageGrab, Image
import win32com.client
import os
import matplotlib.pyplot as plt
import seaborn as sns

############################################################################################################
#extract imgaes from excel file
############################################################################################################

#path to folder in which the files you're extracting are kept
folder = '/Users/FA10041/'

#creates a list of strings of filenames for all .xls in a folder
for root, dirs, files in os.walk(folder):
    xlsfiles=[ _ for _ in files if _.endswith('.xls') ]

# This function extracts an Image from the input excel file and saves it into the specified PNG image path
def saveExcelImageAsPNG(xlsfile):
    # Open the excel application using win32com
    win32 = win32com.client.gencache.EnsureDispatch("Excel.Application")
    # Disable alerts and visibility to the user
    win32.Visible = 0
    win32.DisplayAlerts = 0
    # Open workbook
    wb = win32.Workbooks.Open(folder + xlsfile)

    # Extract second sheet
    sheet = win32.Sheets(2)
    # Extract first image
    shape = sheet.Shapes(1)
    # Copy shape to clipboard
    shape.Copy()
    image = ImageGrab.grabclipboard()
    
    outputPNGImagePath = xlsfile[:len(xlsfile) - 4]
    outputPNGImage = outputPNGImagePath + '.png'
    # Save image
    image.save(folder + outputPNGImage, 'PNG')
    pass

    wb.Close(True)
    win32.Quit()

#loop afor all .xls files in the folder
for xlsfile in xlsfiles:
    saveExcelImageAsPNG(xlsfile)

############################################################################################################
#get colour bar
############################################################################################################

folder = '/Users/FA10041/'

xlsfile = 'report1.xls'

# Open the excel application using win32com
win32 = win32com.client.gencache.EnsureDispatch("Excel.Application")
# Disable alerts and visibility to the user
win32.Visible = 0
win32.DisplayAlerts = 0
# Open workbook
wb = win32.Workbooks.Open(folder + xlsfile)

# Extract second sheet
sheet = win32.Sheets(2)
# Extract second image (the colourbar)
shape = sheet.Shapes(2)
# Copy shape to clipboard
shape.Copy()
colour_bar_img = ImageGrab.grabclipboard()

#get pixel data
width_cb, height_cb = colour_bar_img.size
#convert into list
width_list_cb = list(range(1,width_cb))
height_list_cb = list(range(1,height_cb))

#create a data frame with the XxY pixel coords       
pixel_coords_cb = []
for x in width_list_cb:
    for y in height_list_cb:
        pixel_coords_cb.append(f'{x}x{y}')

df_cb = pd.DataFrame(pixel_coords_cb,columns=['pixel coords'])
#add width and height columns
df_cb[['width', 'height']] = df_cb['pixel coords'].str.split('x', expand=True).astype(int)

#get colour data for each pixel
for index in list(df_cb.index.values):
    df_cb.loc[index,'r'] = colour_bar_img.getpixel((df_cb.loc[index,'width'],df_cb.loc[index,'height']))[0]
    df_cb.loc[index,'g'] = colour_bar_img.getpixel((df_cb.loc[index,'width'],df_cb.loc[index,'height']))[1]
    df_cb.loc[index,'b'] = colour_bar_img.getpixel((df_cb.loc[index,'width'],df_cb.loc[index,'height']))[2]

#select just  a column from the centre of the colour bar
df_cb_only = df_cb.loc[df_cb['width'] == 17]
#select just the 'coloured' pixels from this one column of pixels
df_cb_only = df_cb_only[(df_cb_only['height'] < 380) & (df_cb_only['height'] > 12)]

df_cb_only.reset_index(inplace=True)

#add a values column
for index in list(df_cb_only.index.values):
    df_cb_only.loc[index,'values'] = 200 - index/366*400 #values ranged from 200 to -200

#save as .csv
df_cb_only.to_csv('/Users/FA10041/cb.csv')

#pixel alues were too specific and pixels in the images weren't being matched to the pixels in the colour bar
#therefore I rounded the r,b,g values to the nearest 2

#create values and dicts
values = list(range(0,256,2))
values_dict1 = {value+1:value for value in values}

#change dataframe values
df_cb_rounded = df_cb_only.replace({'r': values_dict1})
df_cb_rounded = df_cb_rounded.replace({'g': values_dict1})
df_cb_rounded = df_cb_rounded.replace({'b': values_dict1})

#test whether any duplicate r,g,b rows have been created - each 'value' most only have 1 set of r,g,b, values
df_cb_rounded['r,g,b'] = df_cb_rounded['r'].astype(str) + df_cb_rounded['g'].astype(str) + df_cb_rounded['b'].astype(str)
df_cb_rounded['r,g,b'].nunique()
df_cb_rounded[df_cb_rounded['r,g,b'].duplicated(keep=False)]

#drop duplicates
df_cb_rounded.drop_duplicates(subset = 'r,g,b', keep='first', inplace=True)

#save as .csv
df_cb_rounded.to_csv('/Users/FA10041/Documents/Camera Windows/Powerview/Lightweight/cb_rounded.csv')

#rounding to the nearest 2 was still creating a lot of nulls therefore I am now rounding ot the nearest 3
#create values and dicts
values = list(range(0,256,3))
values_dict1 = {value+1:value for value in values}
values_dict2 = {value+2:value for value in values}

#change dataframe values - commented out so currently only to nearest 3
df_cb_rounded3 = df_cb_only.replace({'r': values_dict1})
df_cb_rounded3 = df_cb_rounded3.replace({'r': values_dict2})
df_cb_rounded3 = df_cb_rounded3.replace({'g': values_dict1})
df_cb_rounded3 = df_cb_rounded3.replace({'g': values_dict2})
df_cb_rounded3 = df_cb_rounded3.replace({'b': values_dict1})
df_cb_rounded3 = df_cb_rounded3.replace({'b': values_dict2})

#drop duplicates
df_cb_rounded3['r,g,b'] = df_cb_rounded3['r'].astype(str) + df_cb_rounded3['g'].astype(str) + df_cb_rounded3['b'].astype(str)
df_cb_rounded3['r,g,b'].nunique()
df_cb_rounded3[df_cb_rounded3['r,g,b'].duplicated(keep=False)]

df_cb_rounded3.drop_duplicates(subset = 'r,g,b', keep='first', inplace=True)

#save as .csv
df_cb_rounded3.to_csv('/Users/FA10041/Documents/Camera Windows/Powerview/Lightweight/cb_rounded3.csv')

############################################################################################################
#get pixel data for image in .xls file
############################################################################################################
#path to folder in which the files you're extracting are kept 
folder = '/Users/FA10041/D'

#creates a list of strings of filenames for all .png in a folder - that we extracted from the .xls files earlier
for root, dirs, files in os.walk(folder):
    images=[ _ for _ in files if _.endswith('.png') ]

def get_pixel_data(image_name):
    #open image
    image = Image.open(image_name)
    
    pathname, extension = os.path.splitext(image_name)
    
    #get width and height in pixels of image and create lists
    width, height = image.size
    
    width_list = list(range(1,width))
    height_list = list(range(1,height))
    
    #create a pixel coordinate in form XxY for all pixels
    pixel_coords = []
    for x in width_list:
        for y in height_list:
            pixel_coords.append(f'{x}x{y}')
    #create a Dataframe of the pixel coords        
    df = pd.DataFrame(pixel_coords,columns=['pixel coords'])
    
    #add width and height columns
    df[['width', 'height']] = df['pixel coords'].str.split('x', expand=True).astype(int)
        
    #get colour data for each pixel
    for index in list(df.index.values):
        df.loc[index,'r'] = image.getpixel((df.loc[index,'width'],df.loc[index,'height']))[0]
        df.loc[index,'g'] = image.getpixel((df.loc[index,'width'],df.loc[index,'height']))[1]
        df.loc[index,'b'] = image.getpixel((df.loc[index,'width'],df.loc[index,'height']))[2]
    
    #create a df without the black pixels
    df_colour = df.loc[(df['r'] != 0) & (df['g'] != 0) & (df['b'] != 0)]
       
    return df_colour.to_csv(pathname + '.csv')

#loop for images
[get_pixel_data(folder + image_name) for image_name in images ]

############################################################################################################
#get 'value' data and extract values for just the centre of the image
############################################################################################################

#create values and dicts for rounding
values = list(range(0,256,3))
values_dict1 = {value+1:value for value in values}
values_dict2 = {value+2:value for value in values}

#open colour bar data we extraced earlier
df_cb_rounded3 = pd.read_csv('/Users/FA10041/cb_rounded3.csv')

def get_data(pixels):
    
    pathname, extension = os.path.splitext(pixels)
    
    #open data and create df
    df_colour = pd.read_csv(pixels)
    
    #rounding to r,g,b values to nearest 3 to try and remove nans
    #change dataframe values 
    df_colour_rounded3 = df_colour.replace({'r': values_dict1})
    df_colour_rounded3 = df_colour_rounded3.replace({'r': values_dict2})
    df_colour_rounded3 = df_colour_rounded3.replace({'g': values_dict1})
    df_colour_rounded3 = df_colour_rounded3.replace({'g': values_dict2})
    df_colour_rounded3 = df_colour_rounded3.replace({'b': values_dict1})
    df_colour_rounded3 = df_colour_rounded3.replace({'b': values_dict2})
    
    #match the 'values' from the colour bar with the corresponding pixel in the image
    for index in list(df_colour_rounded3.index.values):
        for index_cb in list(df_cb_rounded3.index.values):
            if df_colour_rounded3.loc[index,'r'] == df_cb_rounded3.loc[index_cb,'r'] and df_colour_rounded3.loc[index,'b'] == df_cb_rounded3.loc[index_cb,'b'] and df_colour_rounded3.loc[index,'g'] == df_cb_rounded3.loc[index_cb,'g']:
                df_colour_rounded3.loc[index,'value'] = df_cb_rounded3.loc[index_cb,'value']
    
    #Extracting the 'values' from the centreline of the image
    
    #due to the images being captued by physical measurement the centreline of the image is not neccessarily a the middle pixel of the image
    
    #find which rows from the original image were coloured and create a df
    coloured_rows = df_colour_rounded3['height'].unique()
    coloured_rows_df = pd.DataFrame(coloured_rows,columns=['height'])
    
    #loop to find the centre column of the image
    for row in coloured_rows:
        #find the middle column for each row in the image
        coloured_rows_df.loc[(coloured_rows_df['height'] == row), 'centre'] = df_colour_rounded3['width'].loc[df_colour_rounded3['height'] == row].median(skipna=True)
        #find the right most column for each row in the image so I can make sure pixel positions make sense
        coloured_rows_df.loc[(coloured_rows_df['height'] == row), 'right'] = df_colour_rounded3['width'].loc[df_colour_rounded3['height'] == row].min(skipna=True)
        #find the left most column for each row in the image so I can make sure pixel positions make sense
        coloured_rows_df.loc[(coloured_rows_df['height'] == row), 'left'] = df_colour_rounded3['width'].loc[df_colour_rounded3['height'] == row].max(skipna=True)
    #centre is the meadian of the medians for each row
    centre = coloured_rows_df['centre'].median()
    
    #create a df of the data along the centreline
    linescan_df = df_colour_rounded3.loc[(df_colour['width'] == centre)]
    
    #save image and linescan data
    df_colour_rounded3.to_csv(pathname + ' rounded3.csv')
    linescan_df.to_csv(pathname + ' linescan.csv')
    
    #plot the centreline data
    sns.set_theme(style='whitegrid')
    return sns.scatterplot(x='height', y='distortion', data=linescan_df)

#path to folder in which the files you're extracting are kept 
folder = '/Users/FA10041/'

#get list of the .csv files we extracted earlier for the r,g,b values for the images 
for root, dirs, files in os.walk(folder):
    images_data=[ _ for _ in files if _.endswith('.csv') ]

#loop for all images in the folder
[get_data(folder + image_name) for image_name in images_data ]
