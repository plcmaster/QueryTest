from tkinter import *
from tkcalendar import *



root = Tk()
root.title("Regalatelo Shopify Data")
root.geometry("700x600+10+20")


cal = Calendar(root,selectmode="day",year= 2021, month=8, day=4)
cal.pack()

def get_date():
        label.config(text=cal.get_date())

quit = Button(root, text="Salir", command=root.destroy).place(x=650, y=550)

button= Button(root, text= "Select the Date", command= get_date)
button.pack(pady=20)

label= Label(root, text="")
label.pack(pady=20)



root.mainloop()