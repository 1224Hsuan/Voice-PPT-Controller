import PPT_Controller as pptc
import SpeechRecognition as sr
import Start
import tkinter as tk
from PIL import Image, ImageTk
from tkinter import filedialog
import threading

class UI():
    def __init__(self, control):
        global canvas, root
        self.root = control
        self.root.config()
        self.root.title('Voice PPT Controller')
        self.root.geometry('1000x600')
        canvas = tk.Canvas(self.root,
                           width = 1000,
                           height = 600,
                           highlightthickness = 0,
                           borderwidth = 0)
        canvas.place(x = 0, y = 0)
        window(self.root)

class window():
    def __init__(self, control):
        self.control = control
        self.interface = tk.Frame(self.control)
        self.interface.place(x = 0, y = 0)
        self.start()

    def start(self):
        global image, imageFile, filePath
        global bg, title, msg, msg2, labelPath, btnOpen, btnOK
        global img1, img2, img3, img4

        # set the initial parameter
        filePath = ''
        
        # set the background image
        image = Image.open('background.jpg')
        image = image.resize((1000, 600))
        imageFile = ImageTk.PhotoImage(image)
        bg = canvas.create_image(0, 0, image = imageFile, anchor = 'nw')

        # set the label
        img1 = self.imageCrop(200, 25, 800, 100)
        title = self.createLabel('Voice PPT Controller',
                                 '#FFBB00', 'Cooper Black', 28, img1,
                                 200, 25)
        img2 = self.imageCrop(50, 140, 380, 210)
        msg = self.createLabel('Please choose the data: ',
                               '#5500FF', 'Times New Roman', 23, img2,
                               50, 140)
        img3 = self.imageCrop(50, 250, 210, 300)
        msg2 = self.createLabel('PPT Path: ',
                                '#5500FF', 'Times New Roman', 23, img3,
                                50, 250)
        img4 = self.imageCrop(0, 300, 1000, 400)
        labelPath = self.createLabel('',
                                     '#0000FF', 'Times New Roman', 20, img4,
                                     0, 300)

        # set the button
        btnOpen = tk.Button(text = 'Choose',
                            font = ('Arial', 15, 'bold'),
                            command = self.pptChoose)
        btnOpen.place(x = 400, y = 210)
        btnOK   = tk.Button(text = 'OK',
                            font = ('Arial', 15, 'bold'),
                            command = self.pptOpen,
                            state = tk.DISABLED)
        btnOK.place(x = 420, y = 400)

    # crop the image
    def imageCrop(self, x1, y1, x2, y2):
        global image
        cropped = image.crop((x1, y1, x2, y2))
        img = ImageTk.PhotoImage(cropped)
        return img

    # create the label
    def createLabel(self, text, color, textType, textSize, img, x1, y1):
        label = tk.Label(image = img,
                         fg    = color,
                         text  = text,
                         highlightthickness = 0,
                         borderwidth = 0,
                         padx  = 0,
                         pady  = 0,
                         compound = 'center',
                         font = (textType, textSize, 'bold'))
        label.place(x = x1, y = y1)
        return label

    # to let user choose the data
    def pptChoose(self):
        global filePath, labelPath
        temp = filedialog.askopenfilename(
                   title = u'PowerPoint Open',
                   filetypes = (('PowerPoint', ['*.pptx', '*.ppt']), ))
        if temp != '':
            filePath = temp
        if filePath == '':
            return
        labelPath.config(text = filePath)
        btnOK.config(state = tk.NORMAL)     

    # open the data
    def pptOpen(self):
        global filePath, PPT
        if filePath != '':
            thread1 = threading.Thread(target = self.pptControll())
            thread1.start()
            thread2 = threading.Thread(target = sr.Recognition())
            thread2.start()

    # make the slide fullscreen and close the user interface
    def pptControll(self):
        global PPT
        btnOK.config(state = tk.DISABLED)
        PPT = pptc.Controller()
        PPT.slideShow()
        Start.closeWindow()

