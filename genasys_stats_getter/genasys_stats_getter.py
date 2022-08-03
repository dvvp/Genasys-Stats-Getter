import tkinter as tk
from tkinter import filedialog as fd
import pandas as pd
import numpy as np
import threading

root = tk.Tk()

root.wm_title("Genasys Test Stats Getter")
root.resizable(False, False)

canvas = tk.Canvas(root, width = 600, height = 300)
canvas.grid(columnspan = 3, rowspan = 3)

# instructions
instructions = tk.Label(root, text = "Select the test files you want to parse", font = "Arial")
instructions.grid(columnspan = 3, column = 0, row = 0)

# sheet names
sheet_names = {
            # 450XL
            "Maximum Output (Wide-Off, Ch1) - PASSED.xlsx" : "max_out_ch1", 
            "Maximum Output (Wide-Off, Ch2) - PASSED.xlsx" : "max_out_ch2",
            "Sensitivity (MP3 Input, Sound Projection Narrow, VB OFF, (Ch1) - PASSED.xlsx" : "sens_nar_vb_off_ch1",
            "Sensitivity (MP3 Input, Sound Projection Narrow, VB ON, (Ch1) - PASSED.xlsx" : "sens_nar_vb_on_ch1",
            "Sensitivity (MP3 Input, Sound Projection Wide, VB OFF, Ch1) - PASSED.xlsx" : "sens_wide_vb_off_ch1",
            "Sensitivity (MP3 Input, Sound Projection Wide, VB OFF, Ch2) - PASSED.xlsx" : "sens_wide_vb_off_ch2",
            "Sensitivity (MP3 Input, Sound Projection Wide, VB ON, Ch1) - PASSED.xlsx" : "sens_wide_vb_on_ch1",

            # 1u amplifier
            "Sensitivity (RCA-1 Input, RCA-2 Output)(Ch1) - PASSED.xlsx" : "sens_rca_rca_ch1", 
            "Sensitivity (RCA-1 Input, RCA-2 Output)(Ch6) - PASSED.xlsx" : "sens_rca_rca_ch6",
            "Sensitivity (RCA-1 Input, RCA-2 Output)(Ch6) - PASSED.xlsx" : "sens_rca_rca_ch6", 
            "Sensitivity (XLR Input, -6dB Muted, (Ch1) - PASSED.xlsx" : "sens_xlr_6dB_ch1", 
            "Sensitivity (XLR Input, -12dB Muted, (Ch1) - PASSED.xlsx" : "sens_xlr_12dB_ch1", 
            "Sensitivity (XLR Input, Ch1) - PASSED.xlsx" : "sens_xlr_ch1", 
            "Sensitivity (XLR Input, Ch2) - PASSED.xlsx" : "sens_xlr_ch2", 
            "Sensitivity (XLR Input, Ch3) - PASSED.xlsx" : "sens_xlr_ch3",
            "Sensitivity (XLR Input, Ch4) - PASSED.xlsx" : "sens_xlr_ch4", 
            "Sensitivity (XLR Input, Mute ON, Ch1) - PASSED.xlsx" : "sens_xlr_mute_on_ch1", 
            "Sensitivity (XLR Input, RCA-2 Output)(Ch6) - PASSED.xlsx" : "sens_xlr_rca_ch6", 
            "Volume Control Function (XLR Input, Ch1) - PASSED.xlsx" : "vol_fn_xlr_ch1"
            }

def open_files():

    global files, wo_id_filtered, sheet_names

    browse_text.set("loading...")
    files = fd.askopenfilenames(title = "Select test files", filetype = [("Excel files", "*.xlsx")])

    if files:

        # removes serial numbers
        wo_id = np.unique(tuple(map(lambda x : x.split("] ")[1], files)))

        # filters extensions for passed tests and tests that are not SNR
        wo_id_filtered = list(filter(lambda test : test[-11:] == "PASSED.xlsx" and test[:3] != "SNR", wo_id))

        instructions['text'] = "Reselect test files or write output"
        write_btn['state'] = tk.NORMAL
    else:
        write_btn['state'] = tk.DISABLED
        instructions['text'] = "Select the test files you want to parse"
        
    browse_text.set("Browse")

def write_files():

    browse_btn['state'] = tk.DISABLED
    write_btn['state'] = tk.DISABLED

    instructions['text'] = "Loading... This might take a while"

    # reads excel files, and finds stats
    try:
        with pd.ExcelWriter('genasys_test_stats.xlsx') as writer:

            for i in wo_id_filtered:
      
                progress_text["text"] = "Writing... " + i.split(" -")[0]
                progress_text.update()

                created_final_df = False

                for j in files:
                    if i in j:
                        if ("Sensitivity" in i or "Volume Control Function" in i):
                            sheet = "RMS Level"
                        else:
                            sheet = "FFT Spectrum"
                        df = pd.read_excel(j, sheet_name = sheet, engine = 'openpyxl')
                        df = df.tail(df.shape[0] - 3)
                        df.rename(columns = {sheet : "X", "Unnamed: 1" : "Y"}, inplace = True)
                        if created_final_df == False:
                            final_df = df.get(["Y"])
                            created_final_df = True
                        else:
                            final_df = pd.concat([final_df, df.get(["Y"])], axis = 1)

                final_df = final_df.assign(X = df.get("X"))
                final_df = final_df.set_index("X")
                maximum = np.max(final_df, axis = 1)
                minimum = np.min(final_df, axis = 1)
                mean = np.mean(final_df, axis = 1)
                median = np.median(final_df, axis = 1)
                sd = np.std(final_df, axis = 1)
                final_df = final_df.assign(MAX = maximum, MIN = minimum, MEAN = mean, MEDIAN = median, SD = sd)
                final_df = final_df.get(["MAX", "MIN", "MEAN", "MEDIAN", "SD"])
                final_df.to_excel(writer, sheet_name = sheet_names[i])
    
    except PermissionError as perm_error:
        instructions["text"] = "Please close 'genasys_test_stats.xlsx' and try again"
        progress_text["text"] = perm_error
        write_btn["state"] = tk.DISABLED
        browse_btn["state"] = tk.NORMAL
        return
    except Exception as e:
        instructions["text"] = "Please check formatting and try again"
        progress_text["text"] = e
        write_btn["state"] = tk.DISABLED
        browse_btn["state"] = tk.NORMAL
        return

    instructions["text"] = "Select the test files you want to parse"
    progress_text["text"] = "Writing completed"
    write_btn['state'] = tk.DISABLED
    browse_btn['state'] = tk.NORMAL

def start_open_thread():
    threading.Thread(target = open_files).start()

def start_write_thread():
    threading.Thread(target = write_files).start()

# browse button
browse_text = tk.StringVar()
browse_btn = tk.Button(root, textvariable = browse_text, command = lambda : start_open_thread(), font = "Arial", bg = "#297fc6", fg = "white", cursor = "hand2", height = 2, width = 15)
browse_text.set("Browse")
browse_btn.grid(column = 0, row = 1)

# write button
write_text = tk.StringVar()
write_btn = tk.Button(root, textvariable = write_text, command = lambda : start_write_thread(), font = "Arial", bg = "#297fc6", fg = "white", cursor = "hand2", height = 2, width = 15, state = tk.DISABLED)
write_text.set("Write")
write_btn.grid(column = 2, row = 1)

# progress text
progress_text = tk.Label(root, text = "", font = "Arial")
progress_text.grid(columnspan = 3, column = 0, row = 2)

root.mainloop()
