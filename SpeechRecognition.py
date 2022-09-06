import UserInterface as ui
import Start
import pyaudio
import speech_recognition as sr
import logging
import sys

class Recognition:
    def __init__(self):
        recognition = sr.Recognizer()
        logging.basicConfig(level = logging.DEBUG)
        exitState = 0

        while True:
            with sr.Microphone() as source:
                recognition.adjust_for_ambient_noise(source)
                logging.info('Please say something...')
                audio = recognition.record(source, duration = 5)
            logging.info('Recognizing...')

            transcript = recognition.recognize_google(audio,
                                                    show_all = True)

            if transcript != []:
                for sentence in transcript['alternative']:
                    # turn the upper-case letter to lower-case letter
                    tmp = sentence['transcript'].lower()
                    temp = tmp.split(' ')
                    
                    if temp[0] == 'slide':
                        try:
                            slide = int(temp[1])
                            print(f'get slide {slide}')
                            ui.PPT.specificSlide(slide)
                            break
                        except:
                            pass
                        
                    # delete '-' and ' ' characters
                    characters = ' -'
                    temp = ''.join(x for x in tmp if x not in characters)
                    if   temp == 'slideshow':
                        print('get slideshow')
                        ui.PPT.slideShow()
                        break
                    elif temp == 'nextslide':
                        print('get next slide')
                        ui.PPT.nextSlide()
                        break
                    elif temp == 'previousslide':
                        print('get previous slide')
                        ui.PPT.previousSlide()
                        break
                    elif temp == 'exit':
                        print('get exit')
                        sys.exit()
                        break
                
