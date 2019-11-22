import os
import time
from fpdf import FPDF
import pyautogui
from PIL import ImageTk , Image as pic
from pynput.keyboard import Key, Controller
from tkinter import *
import subprocess
import PySimpleGUI as sg
import shutil
keyboard = Controller()
sg.change_look_and_feel('DarkAmber')  # Add a touch of color
layout = [[sg.Text('Make your life easy!')],
          [sg.Text('Do you have Windows'), sg.Checkbox("Yes")],
          [sg.Text('Where is your ppsx/pptx?'), sg.FilesBrowse('Findfile')],
          [sg.Text('Where do you wanna safe the pdf?'), sg.FileSaveAs('Pdfname')],
          [sg.Text('Which aspect raton has you presentation?'), sg.Radio('4:3', "ratio", default=True), sg.Radio('16:9', "ratio")],
          [sg.Text('How much time would you like to have between slides?'), sg.Slider(range=(0, 10), orientation='h', size=(5, 20), default_value=1, key='wait')],
          [sg.Text('Which format has your Presentation'), sg.Radio('ppsx', "ratio2", default=True, key='format'), sg.Radio('pptx', "ratio2")],
          [sg.Submit("Start")]]
window = sg.Window('Screenshotking', layout)
button, values = window.Read()
ratio = values.get(1)
var1 = values.get(0)
wait = int(values.get('wait'))
file_path = values.get('Findfile')
pdfname = values.get('Pdfname')
file_format = values.get('format')
window.close()
path = "tmp/"
try:
    os.mkdir(path)
except OSError:
    print("Files will be written in existing Folder")

def choose():
    pictures = 0
    global checked
    global choosenimg
    choosenimg = []
    checked = [False] * (k)
    if ratio:
        x = 800
        y = 600
    else:
        x = 800
        y = 450
    layout = [
        [sg.Input(size=(6, 1), justification='left', key='input')],
        [sg.Canvas(size=(x, y), key='canvas', background_color= 'black'), sg.Checkbox("selectet", key='select')],
        [sg.Button('back'), sg.Button('next'), sg.Button('finish')]
    ]
    window = sg.Window('Select pictures', layout)
    window.Finalize()
    canvas = window['canvas']
    window['input'].update(str(pictures+1) + "/" + str(k))
    org = pic.open(path + str(pictures) + ".png")
    resiced = org.resize((x, y),pic.ANTIALIAS)
    im = ImageTk.PhotoImage(resiced)
    canvas.TKCanvas.create_image(0, 0, image=im, anchor=NW)
    while True:
        event, values = window.read()
        checked[pictures] = values.get('select')
        if event is None:
            break
        if event == 'next':
            if pictures < k-1:
                pictures += 1
        elif event == 'back':
            if pictures > 0:
                pictures -= 1
        elif event == 'finish':
            break
        org = pic.open(path + str(pictures) + ".png")
        resiced = org.resize((x, y), pic.ANTIALIAS)
        im = ImageTk.PhotoImage(resiced)
        canvas.TKCanvas.create_image(0, 0, image=im, anchor=NW)
        window['input'].update(str(pictures + 1) + "/" + str(k))
        window['select'].update(value=checked[pictures])
    counter = 0
    for i in checked:
        if i:
            choosenimg.append(counter)
        counter += 1

def calculate_aspect(width: int, height: int):
    temp = 0

    def gcd(a, b):
        """The GCD (greatest common divisor) is the highest number that evenly divides both width and height."""
        return a if b == 0 else gcd(b, a % b)

    if width == height:
        return "1:1"

    if width < height:
        temp = width
        width = height
        height = temp

    divisor = gcd(width, height)

    x = int(width / divisor) if not temp else int(height / divisor)
    y = int(height / divisor) if not temp else int(width / divisor)

    return [x, y]


def autoscreen():
    if var1:
        os.startfile(file_path)
    else:
        subprocess.run(['open', file_path], check=True)
    if file_format:
        time.sleep(3)
    else:
        time.sleep(3)
        keyboard.press(Key.f5)
        time.sleep(2)
    i = True
    global k
    k = 0
    while i:
        myScreenshot = pyautogui.screenshot()
        myScreenshot.save(path + str(k) + '.png')
        im = pic.open(path + str(k) + '.png')
        pixels = im.getdata()  # get the pixels as a flattened sequence
        black_thresh = 260
        nblack = 0
        for pixel in pixels:
            if sum(pixel) < black_thresh:
                nblack += 1
        n = len(pixels)

        if (nblack / float(n)) > 0.5:
            i = False
        else:
            i = True
            l = im.size
            aspect = calculate_aspect(l[0], l[1])
            if ratio:
                multi = (aspect[0] - (4 * (aspect[1] / 3))) / 2
                w = l[0] / aspect[0]
                h = l[1]
                imn = im.crop((multi * w, 0, l[0] - multi * w, h))
            else:
                multi = (aspect[1]-(9 * (aspect[0] / 16))) / 2
                w = l[0]
                h = l[1] / aspect[0]
                imn = im.crop((0, multi * h, w , l[1] - multi * h))
            imn.save(path + str(k) + '.png', "PNG")
            k += 1

        keyboard.press(Key.space)
        time.sleep(wait)
    choose()
    fpdf = FPDF('L', 'mm', 'A4')
    while len(choosenimg) > 0:
        fpdf.add_page()
        print(len(choosenimg))
        if ratio:
            if len(choosenimg) > 3:
                fpdf.image(path + str(choosenimg.pop(0)) + '.png', 2, 2, 136, 102)
                fpdf.image(path + str(choosenimg.pop(0)) + '.png', 159, 2, 136, 102)
                fpdf.image(path + str(choosenimg.pop(0)) + '.png', 2, 106, 136, 102)
                fpdf.image(path + str(choosenimg.pop(0)) + '.png', 159, 106, 136, 102)
            elif len(choosenimg) > 2:
                fpdf.image(path + str(choosenimg.pop(0)) + '.png', 2, 2, 136, 102)
                fpdf.image(path + str(choosenimg.pop(0)) + '.png', 159, 2, 136, 102)
                fpdf.image(path + str(choosenimg.pop(0)) + '.png', 2, 106, 136, 102)
            elif len(choosenimg) > 1:
                fpdf.image(path + str(choosenimg.pop(0)) + '.png', 2, 2, 136, 102)
                fpdf.image(path + str(choosenimg.pop(0)) + '.png', 159, 2, 136, 102)
            else:
                fpdf.image(path + str(choosenimg.pop(0)) + '.png', 2, 2, 136, 102)
        else:
            if len(choosenimg) > 3:
                fpdf.image(path + str(choosenimg.pop(0)) + '.png', 2, 2, 144, 81)
                fpdf.image(path + str(choosenimg.pop(0)) + '.png', 151, 2, 144, 81)
                fpdf.image(path + str(choosenimg.pop(0)) + '.png', 2, 125, 144, 81)
                fpdf.image(path + str(choosenimg.pop(0)) + '.png', 151, 125, 144, 81)
            elif len(choosenimg) > 2:
                fpdf.image(path + str(choosenimg.pop(0)) + '.png', 2, 2, 144, 81)
                fpdf.image(path + str(choosenimg.pop(0)) + '.png', 151, 2, 144, 81)
                fpdf.image(path + str(choosenimg.pop(0)) + '.png', 2, 125, 144, 81)
            elif len(choosenimg) > 1:
                fpdf.image(path + str(choosenimg.pop(0)) + '.png', 2, 2, 144, 81)
                fpdf.image(path + str(choosenimg.pop(0)) + '.png', 151, 2, 144, 81)
            else:
                fpdf.image(path + str(choosenimg.pop(0)) + '.png', 2, 2, 144, 81)
    fpdf.output(pdfname + ".pdf", "F")


autoscreen()
shutil.rmtree(path)
