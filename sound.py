import pygame.mixer
import time

def Sound():
    pygame.mixer.init()
    pygame.mixer.music.load("po.mp3")
    pygame.mixer.music.play(1)

    time.sleep(3)

    pygame.mixer.music.stop()

#if __name__ == '__main__':
    #Sound()
    
