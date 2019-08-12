from tkinter import *
from time import sleep
import threading

class GUI(Frame):
    def __init__(self, master):
        super().__init__()
        self.master = master
        self.button = Button(master, text='inicializar', command=change_loop)
        self.button.pack()
        self.button2 = Button(master, text='parar', command=lambda:x)
        self.button2.pack()

    def init_loop(self):
        loop_instance = Loop()
        loop_instance.loop_status = True

    def stop_loop(self):
        pass
      

class Loop:
    def __init__(self):
        self.loop_status = False
        self.thread = threading.Thread(target=self.run_loop)

    @property
    def loop_status(self):
        print(self.loop_status)
        return self.loop_status

    @loop_status.setter
    def loop_status(self, status):
        self.loop_status = status

    def run_loop():
        if loop_status: 
            self.thread.start()      
            while True:       
                print('Tou rodando o loop')        
                print(loop_status)
                sleep(1)
loop = False
def change_loop():
    loop = True

root = Tk()
gui = GUI(root)
root.mainloop()