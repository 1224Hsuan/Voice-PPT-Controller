import UserInterface as ui
import win32com
import win32com.client
import pythoncom
from pptx import Presentation
import os
import threading

class Controller():
    def __init__(self):
        global filePath, slideNum, index
        pythoncom.CoInitialize()
        self.ppt = win32com.client.Dispatch('PowerPoint.Application')
        self.pptfile = self.ppt.Presentations.Open(ui.filePath)
        slideNum = self.pptfile.Slides.Count
        index = 1

    # fullscreen mode
    def slideShow(self):
        if self.isActive():
            try:
                self.ppt.ActivePresentation.SlideShowSettings.Run()
            except:
                pass

    # get the current slide index
    def getSlideIndex(self):
        if self.isActive():
            global index
            try:
                index = self.ppt.ActiveWindow.View.Slide.SlideIndex
            except:
                index = self.ppt.SlideShowWindows(1).View.CurrentShowPosition
            
    # go to previous slide
    def previousSlide(self):
        global index
        self.getSlideIndex()
        if index > 1:
            self.gotoSlide(index - 1)

    # go to next slide
    def nextSlide(self):
        global slideNum, index
        self.getSlideIndex()
        if index < slideNum:
            self.gotoSlide(index + 1)
                
    # go to specific slide
    def specificSlide(self, temp):
        global slideNum
        if temp >= 1 and temp <= slideNum:
            self.gotoSlide(temp)

    # get the number of powerpoint-windows
    def isActive(self):
        active = self.ppt.Presentations.Count
        if active > 0:
            return True
        return False

    # show the specific slide
    def gotoSlide(self, idx):
        if self.isActive():
            try:
                self.ppt.ActiveWindow.View.GotoSlide(idx)
            except:
                self.ppt.SlideShowWindows(1).View.GotoSlide(idx)
    

