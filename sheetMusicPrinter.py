# Dependencies:
# Ghostscript: https://www.ghostscript.com/releases/gsdnld.html

#!/usr/bin/env python3
import shutil
import os
import sys
import pathlib
import logging
import argparse
import tkinter as tk
import tkinter.ttk as ttk
import win32print
import ghostscript
import locale
import math


instruments = [
    ["Score", "Partitur"],
    ["Piccolo"],
    ["Flute", "Fløyte"],
    ["Alto flute"],
    ["Oboe"],
    ["Bassoon"],
    ["Clarinet", "Klarinett"],
    ["Alto Clarinet", "Alt Klarinett", "Klarinett Alt"],
    ["Bass Clarinet", "Bass Klarinett", "Klarinett Bass"],
    ["Alto Sax", "Alto Saxophone", "Alt Saxofon", "Saxofon Alt", "Alt Sax", "Sax Alt", "Altsaxofon", "Altsax", "Altsaksofon"],
    ["Tenor Sax", "Tenor Saxophone", "Tenor Saxofon", "Saxofon Tenor", "Tenor Sax", "Sax Tenor", "Tenorsax", "Tenorsaxofon", "Tenorsaksofon"],
    ["Baritone Sax", "Baritone Saxophone", "Baryton Saxofon", "Saxofon Baryton", "Baryton Sax", "Sax Baryton", "Baritonsax", "Barytonsax", "Baritonsaxofon", "Barytonsaxofon", "Barytonsaksofon"],
    ["Contrabassoon"],
    ["Horn"],
    ["Trumpet", "Trompet", "Kornett", "Cornet"],
    ["Trombone"],
    ["Bass Trombone"],
    ["Euphonium"],
    ["Baritone", "Baryton"],
    ["Tuba"],
    ["Bass"],
    ["Timpani"],
    ["Percussion", "Perkusjon", "Drums"],
    ["Harp"],
    ["Piano","Keyboard"],
    ["Choir"],
]

tunings = ["Bb", "Eb", "F"]

# ØHM
besetning_ohm = {
    "Piccolo": 1,
    "Flute": 3,
    "Clarinet": 6,
    "Bass Clarinet": 1,
    "Bassoon": 1,
    "Alto Sax": 3,
    "Tenor Sax": 1,
    "Baritone Sax": 1,
    "Horn": 4,
    "Trumpet": 8,
    "Trombone": 8,
    "Bass Trombone": 1,
    "Euphonium": 4,
    "Baritone": 4,
    "Bassoon" : 1,
    "Tuba": 3,
    "Bass": 1,
    "Timpani": 0,
    "Percussion": 1
}

#FHM
besetning_fhm = {
    "Piccolo": 1,
    "Flute": 3,
    "Clarinet": 4,
    "Bass Clarinet": 1,
    "Alto Sax": 2,
    "Tenor Sax": 1,
    "Baritone Sax": 0,
    "Horn": 3,
    "Trumpet": 6,
    "Trombone": 2,
    "Bass Trombone": 1,
    "Euphonium": 2,
    "Tuba": 1,
    "Bass": 1,
    "Timpani": 0,
    "Percussion": 0
}

#FHM
besetning_fhm_uten_overlapp = {
    "Piccolo": 1,
    "Flute": 2,
    "Clarinet": 3,
    "Bass Clarinet": 0,
    "Alto Sax": 1,
    "Tenor Sax": 1,
    "Baritone Sax": 0,
    "Horn": 2,
    "Trumpet": 4,
    "Trombone": 2,
    "Bass Trombone": 1,
    "Euphonium": 1,
    "Baritone": 1,
    "Tuba": 1,
    "Bass": 1,
    "Timpani": 0,
    "Percussion": 1
}

besetning = besetning_ohm



class sheetMusicPrinter(tk.Tk):
    def __init__(self, path):
        tk.Tk.__init__(self)
        self.library_path = path
        self.num_search_results = len(instruments)
        self.title("Sheet Music Printer")
        self.selected_search_result = "None"
        self.libraryEntries = []
        self.musicfiles = []
        self.files_by_instrument = []
        self.readSheetMusicLibrary(path)
        search_label = tk.Label(text="Søk")
        search_label.grid(row=0, column=0)
        search_entry = tk.Entry()
        search_entry.grid(row=1, column=0)
        var = tk.Variable(value=self.libraryEntries)
        self.search_results = tk.Listbox(self, listvariable=var, height=self.num_search_results, width=50)
        self.search_results.grid(row=2, column=0, rowspan=len(instruments))
        self.search_results.bind('<<ListboxSelect>>', self.search_result_selected)
        self.selected_search_result_label = tk.Label(text=self.selected_search_result)
        self.selected_search_result_label.grid(row=0, column=1)
        self.musicfiles_var = tk.Variable(value=self.musicfiles)
        self.musicfiles_box = tk.Listbox(self, listvariable=self.musicfiles_var, height=self.num_search_results, width=50)
        self.musicfiles_box.grid(row=2, column=1, rowspan=len(instruments))

    def add_widgets_for_besetning(self):
        instrument_label = tk.Label(text="Instrument")
        instrument_label.grid(row=1, column=2)
        instrument_label = tk.Label(text="Antall")
        instrument_label.grid(row=1, column=3)
        print_button = tk.Button(self, text="Print alle", command=self.print_all)
        print_button.grid(row=1, column=4)

        # Clean widgets
        for widget in self.grid_slaves():
            if int(widget.grid_info()["row"]) > 1 and int(widget.grid_info()["column"]) > 1:
                widget.grid_forget()

        logging.debug(self.files_by_instrument)

        # Add new widgets
        for instrument_row in range(len(instruments)):
            if len(self.files_by_instrument[instrument_row]) > 0:
                logging.debug("Files by instrument[{}]: {}".format(instrument_row, self.files_by_instrument[instrument_row]))
                e = tk.Entry(self)
                e.grid(row=instrument_row+2, column=2)
                e.insert(tk.END, instruments[instrument_row][0])
                e = tk.Entry(self)
                e.grid(row=instrument_row+2, column=3)
                if instruments[instrument_row][0] in besetning:
                    e.insert(tk.END, besetning[instruments[instrument_row][0]])
                else:
                    e.insert(tk.END, 0)
                e = tk.Button(self, text="Print", command= lambda r=instrument_row: self.print_one(self.files_by_instrument[r])) # No worky, instrument_row bytter verdi før knappen trykkes på
                e.grid(row=instrument_row+2, column=4)
            elif (instruments[instrument_row][0] in besetning and besetning[instruments[instrument_row][0]] > 0):
                e = tk.Entry(self)
                e.grid(row=instrument_row+2, column=2)
                e.insert(tk.END, instruments[instrument_row][0])
                e = tk.Entry(self)
                e.grid(row=instrument_row+2, column=3)
                e.insert(tk.END, "IKKE FUNNET")
                




    def run(self):
        self.mainloop()

    # Read folders in sheet music library and populate the list of entries
    def readSheetMusicLibrary(self, path):
        libraryEntries = []
        for item in path.iterdir():
            if item.is_dir():
                libraryEntries.append(str(item.name))
        self.libraryEntries = libraryEntries
        return libraryEntries

    def search_result_selected(self, event):
        selected_indices = self.search_results.curselection()
        self.selected_search_result = self.search_results.get(selected_indices[0])
        self.selected_search_result_label.config(text=self.selected_search_result)
        logging.info("{}".format(self.selected_search_result))
        self.read_files_for_selected()

    def read_files_for_selected(self):
        musicfiles = []
        path = self.library_path.joinpath(pathlib.Path(self.selected_search_result))
        logging.info("Selected music path: {}".format(path))
        pdf_files = path.glob('*pdf')
        for item in pdf_files:
            if item.is_file():
                #musicfiles.append(item.name)
                musicfiles.append(item)
        self.musicfiles = musicfiles
        musicfiles_names = []
        for file in musicfiles:
            musicfiles_names.append(file.name)
        
        self.musicfiles_var = tk.Variable(value=musicfiles_names)
        self.musicfiles_box.config(listvariable=self.musicfiles_var)
        self.identify_and_sort_files()
        return musicfiles
    
    def identify_and_sort_files(self):
        files_by_instrument = []
        for i in range(0, len(instruments)):
            files_by_instrument.append([])
        for file in self.musicfiles:
            for i in range(0, len(instruments)):
                for instrument_variation in instruments[i]:
                    if instrument_variation.lower() in str(file.name).lower():
                        # We have found a file matching the instrument, but it might match multiple instruments
                        # Find which is the longest instrument name that matches
                        unique = True
                        for j in range(0, len(instruments)): # iterate over all instruments again
                            for instrument_variation_check in instruments[j]: # iterate over all instrument variations again
                                # If there exists a longer instrument variant name that matches than the current match, don't add it to the list
                                if len(instrument_variation_check) > len(instrument_variation) and instrument_variation_check.lower() in str(file.name).lower():
                                    logging.info("{} > {}".format(instrument_variation_check, instrument_variation))
                                    unique = False
                        if unique:
                            files_by_instrument[i].append(file)
        for instrument in files_by_instrument: # DEBUG
            logging.info(instrument)
        self.files_by_instrument = files_by_instrument
        self.add_widgets_for_besetning()
        return files_by_instrument

    def identify_voice(self):
        for i in range (0, len(instruments)):
            # if we have more than one file for a given instrument there is more than one voice
            if len(self.files_by_instrument[i]) > 1:
                for file in self.files_by_instrument[i]:
                    for voice in range(0, 6): # Magic number WARNING, checking up to 6 numbered voices per instrument
                        if voice in str(file.name).lower().replace(self.selected_search_result, ''): # First remove the title of the music in case that contains numbers
                            pass #TODO

    def print_one(self, files):
        for file in files:
            print("print_one: {}".format(file))
            
            args = [
                "-dPrinted", "-dBATCH", "-dNOPAUSE", "-dNOPROMPT"
                "-q",
                "-dNumCopies#1",
                "-sDEVICE#mswinpr2",
                "-sOutputFile#%printer%{}".format(win32print.GetDefaultPrinter()),
                "{}".format(pathlib.PureWindowsPath(file))
            ]

            encoding = locale.getpreferredencoding()
            args = [a.encode(encoding) for a in args]
            logging.debug(args)
            ghostscript.Ghostscript(*args)
            
            #os.system("'C:\\Program Files\\gs\\gs10.03.0\\bin\\gswin64.exe' -dPrinted -dBATCH -dNOPAUSE -dNOPROMPT-q -dNumCopies#1 -sDEVICE#mswinpr2 -sOutputFile#%printer%Kontor 'C:\\Notearkivtest\\Bandology\\Bandology kornett 1.pdf'")
            #import subprocess
            #subprocess.run("C:\\Program Files\\gs\\gs10.03.0\\bin\\gswin64.exe -dPrinted -dBATCH -dNOPAUSE -dNOPROMPT-q -dNumCopies#1 -sDEVICE#mswinpr2 -sOutputFile#%printer%{} \"{}\"".format(win32print.GetDefaultPrinter(), file))

    def print_all(self):
        for instrument in range(len(instruments)):
            if ( len(self.files_by_instrument[instrument]) > 0 ) and ( instruments[instrument][0] in besetning ) and (besetning[instruments[instrument][0]] > 0):
                logging.debug("WOLOLO!")
                for file in self.files_by_instrument[instrument]:
                    logging.debug("Print: {}".format(file))
                    args = [
                        "-dPrinted", "-dBATCH", "-dNOPAUSE", "-dNOPROMPT"
                        "-q",
                        "-dNumCopies#{}".format(math.ceil(besetning[instruments[instrument][0]] / len(self.files_by_instrument[instrument]) )),
                        "-sDEVICE#mswinpr2",
                        "-sOutputFile#%printer%{}".format(win32print.GetDefaultPrinter()),
                        "{}".format(pathlib.PureWindowsPath(file))
                    ]

                    encoding = locale.getpreferredencoding()
                    args = [a.encode(encoding) for a in args]
                    logging.debug(args)
                    #with ghostscript.Ghostscript(*args) as g:
                    ghostscript.Ghostscript(*args)
                    ghostscript.cleanup()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Tool for printing full or partial sets of sheet music")
    parser.add_argument("--directory", dest="directory", required=False,
                        help="The directory containing the sheet music library")
    parser.add_argument("-d", "--debug", help="logging.info debug info", action="store_const", dest="loglevel", const=logging.DEBUG, default=logging.WARNING)

    args = parser.parse_args()
    #logging.basicConfig(level=args.loglevel)
    logging.basicConfig(level="DEBUG")

    if args.directory is not None:
        path = pathlib.Path(args.directory)
    else:
        #path = pathlib.Path(os.getcwd())
        path = pathlib.Path("G:/.shortcut-targets-by-id/1kVVhV0ucB47it1EKNzr_ZCUaDfKxzU7j/Notearkiv") # ØHM
        #path = pathlib.Path("G:/.shortcut-targets-by-id/1cbc_RBXfrVKZ_WuUhTIRk7ZZMPIbgQmC/Notearkiv") # FHM
        #path = pathlib.Path("C:/Notearkivtest")

    logging.info("Notearkiv path: {}".format(path))

    printer = sheetMusicPrinter(path)
    printer.run()