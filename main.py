import customtkinter as ctk
import tkinter as tk
from time import time
from openpyxl import Workbook


wb = None

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("650x800")
        self.title("Match Tracker")
        self.resizable(False, False)

        self.frame_for_buttons = None
        self.frame_for_result = None
        self.timer_label = None
        self.record_label = None
        self.timer_running = False
        self.start_time = None
        self.elapsed_time = 0
        self.start_record_time = None
        self.record_times = []
        self.total_recorded_time = 0
        self.record_number = 0
        self.record_start_time = None        

        self.create_frames()
        self.create_buttons()
        self.create_timer_label()
        self.create_record_label()
        self.create_record_time_label()

    def create_frames(self):
        self.frame_for_buttons = ctk.CTkFrame(self)
        self.frame_for_buttons.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.frame_for_result = ctk.CTkFrame(self, bg_color="white")
        self.frame_for_result.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.record_text = tk.Text(self.frame_for_result)
        self.record_text.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.record_text.configure(font=("Arial", 20))

        self.total_recorded_time_label = ctk.CTkLabel(
            self.frame_for_result, 
            text="Total recorded time: 00:00:00", 
            justify="center"
        )
        self.total_recorded_time_label.pack(side=tk.BOTTOM, pady=10, fill=tk.X)

        self.recorded_time_label = ctk.CTkLabel(self.frame_for_result, text="00:00:00")
        self.recorded_time_label.pack(side=tk.BOTTOM)

    def create_buttons(self):
        play_button = ctk.CTkButton(self.frame_for_buttons, text="Play", command=self.play)
        play_button.pack(side=tk.TOP, padx=10, pady=5)

        pause_button = ctk.CTkButton(self.frame_for_buttons, text="Pause", command=self.pause)
        pause_button.pack(side=tk.TOP, padx=10, pady=5)

        stop_button = ctk.CTkButton(self.frame_for_buttons, text="Stop", command=self.stop)
        stop_button.pack(side=tk.TOP, padx=10, pady=5)

        self.record_button = ctk.CTkButton(self.frame_for_buttons, text="Record")
        self.record_button.pack(side=tk.BOTTOM, padx=10, pady=5)
        self.record_button.bind("<Button-1>", self.start_record)
        self.record_button.bind("<ButtonRelease-1>", self.stop_record)

        self.clear_records_button = ctk.CTkButton(self.frame_for_buttons, text="Clear Record", command=self.clear_records)
        self.clear_records_button.pack(side=tk.BOTTOM, padx=10, pady=5)

    def create_timer_label(self):
        self.timer_label = ctk.CTkLabel(self.frame_for_buttons, text="Time: 00:00:00")
        self.timer_label.pack(side=tk.TOP, pady=10)

    def create_record_label(self):
        self.record_time = ctk.CTkLabel(
            self.frame_for_buttons, 
            text="Record Time: 00:00:00", 
            font=("Arial", 15),
            justify="center"
        )
        self.record_time.pack(side=tk.BOTTOM, pady=10)

    def create_record_time_label(self):
        self.record_time_label = ctk.CTkLabel(
            self.frame_for_buttons,
            text="Recorded Intervals: 00:00:00",
            font=("Arial", 15),
            justify="center"
        )
        self.record_time_label.pack(side=tk.BOTTOM, pady=10)

    def play(self):
        if not self.timer_running:
            if self.start_time is None:
                self.start_time = time()
            else:
                self.start_time += time() - self.pause_time
            self.timer_running = True
            self.update_timer()

    def pause(self):
        if self.timer_running:
            self.elapsed_time += time() - self.start_time
            self.timer_running = False
            self.pause_time = time()

    def stop(self):
        if self.timer_running:
            self.timer_running = False
        self.timer_label.configure(text="Time: 00:00:00")
        self.elapsed_time = 0
        self.start_time = None  # Updated this line
            
    def start_record(self, event):
        if not self.start_record_time and self.timer_running:
            self.start_record_time = time()
            self.record_number += 1
            self.record_start_time = time()

    def stop_record(self, event):
        if self.start_record_time:
            elapsed_time = time() - self.start_record_time
            recorded_seconds = int(elapsed_time)

            timer_elapsed = time() - self.start_time
            timer_hours = int(timer_elapsed // 3600)
            timer_minutes = int((timer_elapsed % 3600) // 60)
            timer_seconds = int(timer_elapsed % 60)
            milliseconds = int((elapsed_time % 1) * 1000)
            seconds = int(elapsed_time)

            self.record_time.configure(text=f"Record Time: {seconds:02d}s {milliseconds:03d}ms")
            self.start_record_time = None  # Stop the recording here
            self.record_time_label.configure(text="Recorded Intervals: 00:00:00")

            #Save to excel
            self.save_to_excel(self.record_number, recorded_seconds, timer_hours, timer_minutes, timer_seconds)

            # Get current text in the text widget and append the new recorded time
            self.record_text.get("1.0", tk.END).strip()
            new_record = f"Recorded Intervals: {seconds:02d}:{milliseconds:03d}s"
            self.record_text.insert(tk.END, f"{new_record}\n")

            # Update the total recorded time
            self.total_recorded_time += recorded_seconds
            total_recorded_time_str = f"Total recorded time: {self.total_recorded_time//3600:02d}:{(self.total_recorded_time%3600)//60:02d}:{self.total_recorded_time%60:02d}"
            self.total_recorded_time_label.configure(text=total_recorded_time_str)
            self.record_times.append((recorded_seconds, (timer_hours,timer_minutes, timer_seconds)))

            self.record_text.delete("1.0", tk.END)
            for i, (recorded_seconds, timer_value) in enumerate(self.record_times):
                record_str = f"Record#{i+1} -- {recorded_seconds:02d}s -- (Match Time:{timer_value[0]:2d}:{timer_value[1]:02d}:{timer_value[2]:02d})\n"
                self.record_text.insert(tk.END, record_str)
            
            self.record_start_time = None
            self.recorded_time_label.configure(text="Recorded Intervals: 00:00:00")

    def update_timer(self):
        if self.timer_running:
            elapsed_time = time() - self.start_time
            hours = int(elapsed_time // 3600)
            minutes = int((elapsed_time % 3600) // 60)
            seconds = int(elapsed_time % 60)
            milliseconds = int((elapsed_time % 1) * 1000)

            self.timer_label.configure(text=f"Time: {hours:02d}:{minutes:02d}:{seconds:02d}:{milliseconds:03d}")
            self.after(1, self.update_timer)
        
            if self.record_start_time:
                elapsed_time = time() - self.record_start_time
                seconds = int(elapsed_time)
                milliseconds = int((elapsed_time % 1) * 1000)

                self.record_time_label.configure(text=f"Recorded Intervals: {seconds:02d}:{milliseconds:03d}")

    def update_total_recorded_time_label(self):
        milliseconds = int((self.total_recorded_time %1) * 1000)
        seconds = int(self.total_recorded_time)
        total_recorded_time_str = f"Total recorded time: {seconds:02d}s {milliseconds:03d}ms"
        self.total_recorded_time_label.configure(text=total_recorded_time_str)

    def update_recorded_time_label(self):
        if self.record_start_time:
            elapsed_time = time() - self.record_start_time
            seconds = int(elapsed_time)
            milliseconds = int((elapsed_time % 1) * 1000)

            self.recorded_time_label.configure(text=f"Recorded Intervals: {seconds:02d}:{milliseconds:03d}")
            self.after(1, self.update_recorded_time_label)


    def clear_records(self):
        self.record_text.delete("1.0", tk.END)
        self.record_times = []
        self.total_recorded_time = 0
        self.update_total_recorded_time_label()
        
    def save_to_excel(self, record_number, recorded_seconds, timer_hours, timer_minutes, timer_seconds):
        global wb
        try:
            if wb is None:
                wb = Workbook()
                ws = wb.active
                ws.append(["Record Number","Record Time (ms)","Total Time (HH:MM:SS)"])
            
            ws = wb.active
            raw_data = [record_number, recorded_seconds, f"{timer_hours:02d}:{timer_minutes:02d}:{timer_seconds:02d}"]
            ws.append(raw_data)
            
            wb.save("match_data.xlsx")
            print("Excel saved succesfully")

        except Exception as e:
            print(f"Error saving to excel {e}")


if __name__ == "__main__":
    App().mainloop()
