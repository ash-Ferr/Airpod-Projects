#Automated Air Pod Serial Retrieval and Label Printing
#Ashley Ferrell  Ashley.Ferrell@assurant.com
#Detects new usb devices, if the new device is an Air Pod Print a SerialNumber Label


from io import BytesIO
import libusb
import libusb_package
import time
import usb.core
import usb.util
import usb.backend.libusb1
import logging
import PIL
from PIL import Image, ImageTk, ImageDraw, ImageFont
import os
import pandas as pd
import queue
import zebra as Zebra
import threading
import tkinter as tk
from tkinter import Tk, Checkbutton, IntVar, Label, Entry, Button,Frame,Canvas,StringVar
from barcode import Code128
from barcode.writer import ImageWriter
import atexit
import sys
from hanging_threads import start_monitoring
from openpyxl import load_workbook
monitoring_thread = start_monitoring(seconds_frozen=10,test_interval=100)
#MESSAGE CONSTANTS
REQUEST_FIN = "attempt ended"

#Create the Error Log File
def create_error_log():
    logging.basicConfig(filename ='error.log',level = logging.INFO)
    #print the program start to the error log
    logging.info("Program Has Started")
    
    #Check for an overwrite condition for the error.log
    error_log_path = 'error.log'
    max_size = 5 * 1024 * 1024  # 5 MB
    #Get error log size
    error_log_size = os.path.getsize(error_log_path)
    #print the log size
    print(f'The Error Log File size is: {error_log_size/1024/1024:.2f}/{max_size/1024/1024} MB')
    # Check if the size of the ERROR.log file is over 5MB
    if error_log_size > max_size:
        print('Now Deleting error log')
        # Overwrite the file
        with open(error_log_path, 'w') as f:
            f.write('')
atexit.register(logging.shutdown)
class USBPrinterManager:
    def __init__(self,update_serial_callback, options_dialog):
        self.serial_number = None #Serial Number of the Air POD connected
        self.model_id = None   #Model Number of the Air POD connected
        self.model_name = None #Name of the AirPod Case Model ex.  A2700,A2190...
        #Libraries for model name comparison
                #Library mapping of the bcdDevice Values and the model names
                #Library mapping of A2190 variation between Lighting and MagSafe Chargers
        self.model_id_to_name = {0x200: "A1602",0x205: 'A1938',0x139:'A2190',0x135:'A2190',0x1a6:'A2566',0x194:'A2566',0x33e:'A2700',0x361a:'A2879'}
        self.serial_check = {'0C6L':'A2190 Lightning','LKKT':'A2190 Lightning','1059':'A2190 MagSafe','1NRC':'A2190 MagSafe'}
        self.serial_lock = threading.Lock()  # Lock for safely updating serial number
        self.z=self.create_printer() #Create the printer object instance for the class
        self.stop_event = threading.Event()#Stop the threads to exit the program
        self.update_serial_callback = update_serial_callback #Call back to pass serial to the UI interface
        self.loop_count = 0 # give a count of the process loop.
        self.usb_thread=None # Initialize 
        self.print_thread=None# Initialize
        self.root = options_dialog.root
     
    #Polls the USB host for New Serial Number, Sends the enw serial to the print thread, and updates
        #the last Serial Number used.  
    def usb_detection_thread(self,root):
        previous_serial = None
        
        while not self.stop_event.is_set():
            #check to see if the main application is still running
            if not self.root.winfo_exists():  sys.exit()
            #Pause the resource
            time.sleep(.5)
            new_serial,model_id = detect_new_device(previous_serial,self.stop_event)
            if new_serial is not None: #Check if a new serial number is detected
                with self.serial_lock:
                    self.serial_number = new_serial
                    self.model_id = model_id
                    self.format_data()
                print(f"New USB device detected: {new_serial} Model: {model_id}")
                #Call the callback function to update the text of the serial_entry widget in Options Dialog
                self.update_serial_callback(new_serial,self.model_name)
                #Set the previous serial to prevent repeat barcode prints
                previous_serial = new_serial
            self.loop_count+=1
            loop_count= self.loop_count
            #print(self.loop_count)
            logging.info(f'End of Loop: #{loop_count}')
        logging.info('Stop event set, exiting loop')            
    def format_data(self):
        last_4_chars = self.serial_number[-4:]
        try:
            if last_4_chars in self.serial_check:
                self.model_name = str(self.serial_check[last_4_chars])
            else:
                self.model_name = str(self.model_id_to_name[self.model_id])
        except Exception as Model_Exc:
                print(f"Model Not Found.  Error: {Model_Exc}")
                self.model_name = "MODEL NOT FOUND"
    #When a new serial number is available, convert to zpl and send command to printer
    def print_thread(self):
        while not self.stop_event.is_set():
            #Check main application status
            if not self.root.winfo_exists():  sys.exit()
            with self.serial_lock:
                if self.serial_number is not None:
                    self.format_data()
                    zpl_command = self.string_to_zpl_code39(self.serial_number,self.model_id,self.model_name)
                    self.zebra_print(zpl_command,self.z)
                    print("Printed barcode for:", self.serial_number)
                    self.serial_number = None
                    self.model_id = None  # Clear the serial number, model, and name.  
                    self.model_name = None

    #start 2 threads for the detecting usb and printing labels
    def start_threads(self):
        self.usb_thread = threading.Thread(target=self.usb_detection_thread,args=(self.root,))
        self.print_thread = threading.Thread(target=self.print_thread)
        self.usb_thread.start()
        self.print_thread.start()
        #SomeChangeHEre

    #Set the stop event to signal the threads to exit their loops
    def stop_threads(self):
        self.stop_event.set()
        logging.info("stop_threads self.stop_event is set")
        if self.usb_thread:self.usb_thread.join()
        logging.info("USB Thread Joined")
        if self.print_thread:self.print_thread.join()
        logging.info("Print Thread Joined")
    #Initialize the printer object
    def create_printer(self):
        z = Zebra.Zebra()
        queues = z.getqueues()
        z=Zebra.Zebra(queues[0])
        return z
    #Convert the serial number string to a zpl command
    def string_to_zpl_code39(self,input_string,model_id,model_name):
        print(input_string)
        zpl_code = "^XA^PRC,C,C~SD20^FO275,120^BY3^B3N,N,100,Y,N^FD" + (input_string) + "^FS"
        zpl_code += f"^FO250,350^A0N,50,50^FD Model_id: {model_id}^FS"
        zpl_code += f"^FO250,400^A0N,50,50^FD Model: {model_name}^FS"
        zpl_code += "^XZ"
        return zpl_code
    #Send Command to the printer
    def zebra_print(self, zpl_code, z):
        z.output(zpl_code)

#User Interface:  Includes a start and stop button.      
class OptionsDialog:
    def __init__(self):
        #Check to ensure file tree exists create the path if not
        #data queue for the communication between Threads
        self.data_queue = queue.Queue()
        #GUI Update Queue
        self.gui_queue = queue.Queue()
        
        #Create the GUI
        self.root = Tk()
        self.root.geometry("800x300") #Size of the window in pixels
        self.root.title("AirPod Serial Extractor") #title of the window
        self.root.attributes("-topmost",True) # Keep the window with the quit button on top of all other windowed objects
        
        #Create a frame and center it in the root window
        self.container = Frame(self.root,bg='white')
        self.container.pack(fill='both',expand=True)
        self.frame1=Frame(self.container,bg='red')
        self.frame1.grid(row=0,column=0,sticky= 'nsew')
        self.frame2=Frame(self.container,bg='blue')
        self.frame2.grid(row=0,column=1,sticky='nsew')
        self.container.grid_columnconfigure(0,weight=1,minsize=100)
        self.container.grid_columnconfigure(1,weight=1,minsize=600)
        
        #StringVar for serial Number
        self.serial_number=StringVar()
        #Create and Place the buttons on the Window
        self.start_button = Button(self.frame1, text="Start", command=self.start_threads)
        self.quit_button = Button(self.frame1, text="Quit", command=self.quit_program)
        #Add a text field to show the Serial Number
        self.serial_entry = Entry(self.frame1,textvariable=self.serial_number)
        
        ######MODEL REMOVED FROM VERSION
        #Add a text field for the model name
        #self.model_entry = Entry(self.frame1)
        
        #Create a label for the barcode in Frame 2
        self.barcode_label = Label(self.frame2)
        self.barcode_label.pack()
        #Add Labels for serial entry and model entry
        self.serial_label = Label(self.frame1,text="Serial Number")
        
        ######MODEL REMOVED FROM VERSION
        #self.model_label = Label(self.frame1 ,text="Model Name")
        
        #Add a checkbox to indicate whether the device passed inspection
        self.passed_var = IntVar()
        self.passed_checkbox = Checkbutton(self.frame1, text = "Passed inspection", variable = self.passed_var)

        #Pack all of the Widgets
        self.start_button.grid(row=0, column=0)
        self.quit_button.grid(row=0, column=1)
        #self.save_button.grid(row=0, column=1)
        #self.indicator.grid(row=0, column=2)
        self.frame1.grid_rowconfigure(1,minsize = 50)
        self.serial_label.grid(row=2, column=0)
        self.serial_entry.grid(row=2, column=1)
        self.frame1.grid_rowconfigure(3, minsize = 20)
        
        #####MODEL REMOVED FROM VERSION
        #self.model_label.grid(row=4, column=0)
        #self.model_entry.grid(row=4, column=1)
        
        self.frame1.grid_rowconfigure(5)
        #self.passed_checkbox.grid(row=5, column=1)

        #stringVar for serial Number
        self.serial_number.trace_add('write',self.create_barcode)
        #Modify the Delete Window Protocol
        self.root.protocol("WM_DELETE_WINDOW",self.quit_program)
        
        
        #Call the test barcode
        barcode_text = "Test-Test-Test"
        barcode = Code128(barcode_text, writer=ImageWriter())
        barcode.save('barcode')

        # Open the barcode image
        image = Image.open('barcode.png')

        # Convert the image to PhotoImage format
        photo = ImageTk.PhotoImage(image)

        #Update self.barcode_label
        self.barcode_label.config(image=photo)
        self.barcode_label.image=photo
        
        
    def create_barcode(self, *args):
        #Get Serial
        serial_number = self.serial_entry.get()
        # Generate the barcode
        barcode = Code128(serial_number, writer=ImageWriter())
        barcode.save('barcode')

        # Open the barcode image
        image = Image.open('barcode.png')

        # Convert the image to PhotoImage format
        photo = ImageTk.PhotoImage(image)

        #Update self.barcode_label
        self.barcode_label.config(image=photo)
        self.barcode_label.image=photo
        


    def flash_indicator(self):
        #Change the indicator color back to red
        self.indicator.itemconfig(1, fill = 'red')

    def check_update_gui(self):
        #Check the gui_queue for update
        try:
            message = self.gui_queue.get_nowait()
            if message == 'update':
                #Update the indicator color to green
                self.indicator.itemconfig(1, fill = 'green')
                #Schedule the color change back to red after 1000 ms
                self.root.after(1000, lambda: self.indicator.itemconfig(1,fill = 'red'))
        except queue.Empty:
            pass
        finally:
            #Schedule this function to be called again after 100 ms
            self.root.after(100, self.check_update_gui)

    def add_to_queue(self):
        #get the serial number and inspection result from GUI
        serial_number = self.serial_entry.get()
        model_name = self.model_entry.get()
        passed_inspection = bool(self.passed_var.get())
        #Add the data to the queue
        self.data_queue.put((serial_number,model_name,passed_inspection))
        #Clear the text field and uncheck the checkbox
        self.serial_entry.delete(0, "end")
        self.model_entry.delete(0,"end")
        self.passed_var.set(0)
        logging.info("Data Passed to Queue")

    #function for the start button 
    def start_threads(self):
        #Begin the threads containing the Device Lookup and Serial Label Print
        printer_manager.start_threads()
        logging.info("Print Manager Thread Started")

    #function for the quit button
    def quit_program(self):
        logging.info("Quit Initiated")
        #call the stop_threads function of the printer_manager object
        printer_manager.stop_threads()
        logging.info("print manager stop event set")
        
        self.data_queue.join()
        logging.info("Data Queue Joined")
        #destroy the root window
        logging.info("Closing the GUI")
        self.root.destroy()
        #Close Log Handlers
        for handler in logging.root.handlers[:]:
            logging.info("Closing Logging Handler")
            handler.close()
            logging.root.removeHandler(handler)
        #stop monitoring thread
        monitoring_thread.stop()
        #Exit the program
        print("System Exit")
        sys.exit()

#Function to detect new Air Pod Connections
def detect_new_device(previous_serial=None,stop_event = None):
    while not (stop_event.is_set()):
        start_time = time.time()
        try:
            devices = list(libusb_package.find(find_all=True))
        except Exception as watdafu:
            print(f'there has been a {watdafu}')
            input('pause here,press enter\n')
            return None,None
        #Iterate through all of the listed USB Connections, return the serial number of Air Pod when Found
        for dev in devices:
            #check if device supports the extraction operation before calling
            try:
                if hasattr(dev, 'iProduct') and dev.iProduct is not None:
                    product = usb.util.get_string(dev, dev.iProduct)
                    if 'Case' in str(product):
                        if hasattr(dev, 'iSerialNumber') and dev.iSerialNumber is not None:
                            serial_number =usb.util.get_string(dev, dev.iSerialNumber)
                            model_id = dev.idProduct if hasattr(dev, 'idProduct') else None
                            if serial_number != previous_serial:
                                end_time = time.time()
                                execution_time = end_time - start_time
                                print(execution_time)
                                return serial_number,model_id
            except (NotImplementedError, ValueError, AttributeError) as e:
                pass
                #logging.error(f'ERROR: {e}')
            except usb.core.USBError as er:
                pass
                #logging.error(f'ERROR: {er}')

        end_time = time.time()
        execution_time = end_time - start_time
        #print(execution_time)
    return None,None

#Can Import as Library to access the USBPrinterManager and OptionsDialog classes, and the device detection function
#the Main function of the program, initialize the class objects and enter the mainloop of our GUI
if __name__ == "__main__":
    create_error_log()
    options_dialog = OptionsDialog()
    #Define the Callback function to update the text of the serial_entry widget in OptionsDialog
    def update_serial(serial_number,model_name):
        options_dialog.serial_entry.delete(0,"end")
        options_dialog.serial_entry.insert(0, serial_number)
        
        ######Remove Model From Version
        #options_dialog.model_entry.delete(0,'end')
        #options_dialog.model_entry.insert(0,model_name)
    #Create instance of USBPrinterManager     
    printer_manager = USBPrinterManager(update_serial, options_dialog)
    options_dialog.root.after(100,options_dialog.check_update_gui)
    options_dialog.root.mainloop()
