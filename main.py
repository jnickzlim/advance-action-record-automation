import os
from datetime import datetime
import croniter
import json
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from tkinter import ttk
from pynput import mouse, keyboard
import pyautogui
import threading
import time
import re
from collections import deque
import logging
import functools

def debounce(wait):
    def decorator(fn):
        def debounced(*args, **kwargs):
            def call_it():
                fn(*args, **kwargs)
            try:
                debounced.t.cancel()
            except(AttributeError):
                pass
            debounced.t = threading.Timer(wait, call_it)
            debounced.t.start()
        return debounced
    return decorator

class ActionList:
    def __init__(self, name, sequence=0, interval=0, active=True):
        self.name = name
        self.actions = []
        self.repeat = 1
        self.sequence = sequence
        self.interval = interval
        self.last_executed = 0
        self.executed = 0
        self.active = active

    def add_action(self, action):
        self.actions.append(action)

    def remove_action(self, index):
        del self.actions[index]

    def clear_actions(self):
        self.actions.clear()

class ActionRecorder:
    def __init__(self, root):
        self.root = root

        # Set default window size 
        self.root.geometry("1000x800") # FOR Mac 
        # self.root.geometry("2050x1300") # FOR VM
        
        # Optionally, you can set minimum and maximum window sizes
        self.root.minsize(700, 700)  # Minimum size
        # self.root.maxsize(1200, 1200)  # Maximum size

        self.root.title("Advanced Action Record Automation - v1.0.10")

        # Create a main frame with padding
        self.main_frame = ttk.Frame(self.root, padding=25)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.action_lists = []
        self.cron_jobs = []
        self.current_list = None
        self.recording = False
        self.replaying = False
        self.paused = False
        self.last_time = None
        self.action_repeat_count = 0
        self.dark_mode = tk.BooleanVar(value=True)  # Set to True by default
        self.special_keys = [
            'space',
                 'enter',
               'shift',
                'ctrl',
                 'alt',
                 'tab',
                 'backspace',
                 'delete',
                 'esc',
                 'up',
                 'down',
                 'left',
                 'right',
                 'home',
                 'end',
                 'pageup',
                 'pagedown'
        ] 

        self.configure_logging()
        self.configure_styles()
        self.create_widgets()
        self.create_bottom_frame()
        self.update_move_buttons()  

        self.listener_mouse = mouse.Listener(on_click=self.on_click)
        self.listener_keyboard = keyboard.Listener(on_press=self.on_press)

        self.listener_mouse.start()
        self.listener_keyboard.start()

        # Apply dark mode by default
        self.toggle_theme()
        self.root.after(1000, self.check_and_execute_cron_jobs)
        time.sleep(2)

    def create_bottom_frame(self):
        self.bottom_frame = ttk.Frame(self.root)
        self.bottom_frame.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.current_time_label = ttk.Label(self.bottom_frame, text="Current Time: ")
        self.current_time_label.pack(side=tk.RIGHT, padx=5, pady=10)

    def configure_logging(self):
        logging.basicConfig(filename='action_recorder.log', level=logging.DEBUG)

    def configure_styles(self):
        self.style = ttk.Style()
        self.style.theme_use('clam')

        # Configure dark mode colors
        bg_color = '#2E2E2E'
        fg_color = '#FFFFFF'
        select_color = '#4A4A4A'
        border_color = '#1E1E1E'  # Dark border color

        self.style.configure('TFrame', background=bg_color)
        self.style.configure('TLabel', background=bg_color, foreground=fg_color)
        self.style.configure('TButton', background=bg_color, foreground=fg_color, bordercolor=border_color)
        self.style.map('TButton', background=[('active', select_color)])
        
        self.style.configure('Big.TButton', padding=(10, 5), font=('Helvetica', 12), width=10, bordercolor=border_color)
        self.style.map('Big.TButton', background=[('active', select_color)])

        self.style.configure('Treeview', 
                            background=bg_color, 
                            foreground=fg_color, 
                            fieldbackground=bg_color, 
                            bordercolor=border_color,
                            lightcolor=border_color,
                            darkcolor=border_color)
        self.style.map('Treeview', background=[('selected', select_color)])

        self.style.configure('TNotebook', background=bg_color, bordercolor=border_color)
        self.style.configure('TNotebook.Tab', background=bg_color, foreground=fg_color, bordercolor=border_color)
        self.style.map('TNotebook.Tab', background=[('selected', select_color)])

        self.style.configure('TEntry', background=bg_color, foreground=fg_color, fieldbackground=bg_color, bordercolor=border_color)
        self.style.configure('TCheckbutton', background=bg_color, foreground=fg_color, bordercolor=border_color)

        # Configure scrollbar colors
        self.style.configure('Vertical.TScrollbar', background=bg_color, bordercolor=border_color, arrowcolor=fg_color, troughcolor=bg_color)
        self.style.configure('Horizontal.TScrollbar', background=bg_color, bordercolor=border_color, arrowcolor=fg_color, troughcolor=bg_color)

        # Additional widget configurations
        self.style.configure('TCombobox', background=bg_color, foreground=fg_color, fieldbackground=bg_color, bordercolor=border_color)
        self.style.map('TCombobox', fieldbackground=[('readonly', bg_color)])

        self.style.configure('TSeparator', background=border_color)

        # Ensure all frames use the dark background
        self.style.configure('TFrame', background=bg_color)

        # Update the main window background
        self.root.configure(bg=bg_color)

    def create_widgets(self):
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.record_frame = ttk.Frame(self.notebook)
        self.replay_frame = ttk.Frame(self.notebook)
        self.cron_frame = ttk.Frame(self.notebook)  # Add this line
        self.settings_frame = ttk.Frame(self.notebook)
        
        self.notebook.add(self.record_frame, text="Record")
        self.notebook.add(self.replay_frame, text="Replay")
        self.notebook.add(self.cron_frame, text="Cron Jobs")  # Add this line
        self.notebook.add(self.settings_frame, text="Settings")
        
        self.create_record_widgets()
        self.create_replay_widgets()
        self.create_cron_jobs_widgets()
        self.create_settings_widgets()

    def create_cron_jobs_widgets(self):
        # Row 0: Top buttons
        top_button_frame = ttk.Frame(self.cron_frame)
        top_button_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

        self.import_cron_button = ttk.Button(top_button_frame, text="Import", command=self.import_cron_jobs, style='Big.TButton', width=10)
        self.import_cron_button.pack(side=tk.LEFT, padx=(0, 5))

        self.save_cron_button = ttk.Button(top_button_frame, text="Save", command=self.save_cron_jobs, style='Big.TButton', width=10)
        self.save_cron_button.pack(side=tk.LEFT, padx=5)

        self.clear_cron_button = ttk.Button(top_button_frame, text="Clear", command=self.clear_cron_jobs, style='Big.TButton', width=10)
        self.clear_cron_button.pack(side=tk.LEFT, padx=5)

        # Row 1: Cron Jobs Tree
        self.cron_jobs_tree = ttk.Treeview(self.cron_frame, columns=("No.", "Name", "Actions", "Cron Expression", "Active", "Last Executed"), show='headings')
        self.cron_jobs_tree.heading("No.", text="No.")
        self.cron_jobs_tree.heading("Name", text="Name")
        self.cron_jobs_tree.heading("Actions", text="Actions")
        self.cron_jobs_tree.heading("Cron Expression", text="Cron Expression")
        self.cron_jobs_tree.heading("Active", text="Active")
        self.cron_jobs_tree.heading("Last Executed", text="Last Executed")
        self.cron_jobs_tree.grid(row=1, column=0, padx=(10, 0), pady=10, sticky="nsew")

        # Set column widths
        self.cron_jobs_tree.column("No.", width=30, minwidth=30, anchor='center')
        self.cron_jobs_tree.column("Name", width=100, minwidth=100, anchor='w')
        self.cron_jobs_tree.column("Actions", width=50, minwidth=50, anchor='center')
        self.cron_jobs_tree.column("Cron Expression", width=50, minwidth=50, anchor='center')
        self.cron_jobs_tree.column("Active", width=80, minwidth=80, anchor='center')
        self.cron_jobs_tree.column("Last Executed", width=60, minwidth=60, anchor='center')

        # Button frame
        button_frame = ttk.Frame(self.cron_frame, width=230)
        button_frame.grid(row=1, column=1, padx=(0, 10), pady=10, sticky="nsew")
        button_frame.grid_propagate(False)

        button_frame.grid_columnconfigure(0, weight=1)

        # Create a style for buttons in the button frame
        button_style = ttk.Style()
        button_style.configure('ButtonFrame.TButton', width=20)

        # Buttons in the button frame
        self.execute_cron_var = tk.BooleanVar(value=True)
        self.execute_cron_checkbox = ttk.Checkbutton(button_frame, text="  |   Execute Cron Jobs", variable=self.execute_cron_var)
        self.execute_cron_checkbox.grid(row=0, column=0, padx=5, sticky="ew")

        self.edit_cron_button = ttk.Button(button_frame, text="Edit", command=self.edit_cron_job, state=tk.DISABLED, style='ButtonFrame.TButton')
        self.edit_cron_button.grid(row=1, column=0, pady=5, padx=5, sticky="ew")

        self.duplicate_cron_button = ttk.Button(button_frame, text="Duplicate", command=self.duplicate_cron_job, state=tk.DISABLED, style='ButtonFrame.TButton')
        self.duplicate_cron_button.grid(row=2, column=0, pady=5, padx=5, sticky="ew")

        self.delete_cron_button = ttk.Button(button_frame, text="Delete", command=self.delete_cron_job, state=tk.DISABLED, style='ButtonFrame.TButton')
        self.delete_cron_button.grid(row=3, column=0, pady=5, padx=5, sticky="ew")

        self.activate_cron_button = ttk.Button(button_frame, text="Activate/Deactivate", command=self.toggle_cron_job_active, state=tk.DISABLED, style='ButtonFrame.TButton')
        self.activate_cron_button.grid(row=4, column=0, pady=5, padx=5, sticky="ew")

        self.play_cron_button = ttk.Button(button_frame, text="Play", command=self.play_cron_job, state=tk.DISABLED, style='ButtonFrame.TButton')
        self.play_cron_button.grid(row=5, column=0, pady=5, padx=5, sticky="ew")
        
        self.current_cron_job_label = ttk.Label(top_button_frame, text="|   Currently Playing: None")
        self.current_cron_job_label.pack(side=tk.LEFT, padx=5)

        # Configure grid weights
        self.cron_frame.grid_columnconfigure(0, weight=1)
        self.cron_frame.grid_columnconfigure(1, weight=0)
        self.cron_frame.grid_rowconfigure(1, weight=1)

        # Add a horizontal scrollbar
        h_scroll = ttk.Scrollbar(self.cron_frame, orient="horizontal", command=self.cron_jobs_tree.xview)
        h_scroll.grid(row=2, column=0, columnspan=2, sticky="ew")
        self.cron_jobs_tree.configure(xscrollcommand=h_scroll.set)

        # Bind selection event to enable/disable buttons
        self.cron_jobs_tree.bind('<<TreeviewSelect>>', self.on_cron_job_select)


    def on_cron_job_select(self, event):
        selected = self.cron_jobs_tree.selection()
        if selected:
            self.edit_cron_button.config(state=tk.NORMAL)
            self.duplicate_cron_button.config(state=tk.NORMAL)
            self.delete_cron_button.config(state=tk.NORMAL)
            self.activate_cron_button.config(state=tk.NORMAL)
            self.play_cron_button.config(state=tk.NORMAL)
        else:
            self.edit_cron_button.config(state=tk.DISABLED)
            self.duplicate_cron_button.config(state=tk.DISABLED)
            self.delete_cron_button.config(state=tk.DISABLED)
            self.activate_cron_button.config(state=tk.DISABLED)
            self.play_cron_button.config(state=tk.DISABLED)

    def play_cron_job(self):
        selected = self.cron_jobs_tree.selection()
        if selected:
            item = selected[0]
            original_index = int(self.cron_jobs_tree.item(item, 'tags')[0])  # Get the original index from the tag
            cron_job = self.cron_jobs[original_index]
            
            if self.execute_cron_var.get():
                self.play_cron_button.config(text="Stop")
                self.replaying = True
                threading.Thread(target=self.execute_cron_job_once, args=(cron_job,)).start()
            else:
                messagebox.showinfo("Execution Disabled", "Cron job execution is currently disabled. Enable it using the checkbox.")

    def duplicate_cron_job(self):
        selected = self.cron_jobs_tree.selection()
        if selected:
            item = selected[0]
            index = int(self.cron_jobs_tree.item(item, 'tags')[0])  # Get the original index from the tag
            cron_job = self.cron_jobs[index].copy()
            cron_job['name'] = f"{cron_job['name']} {index}"
            self.cron_jobs.append(cron_job)
            self.update_cron_jobs_list()

    def execute_cron_job_once(self, cron_job):
        self.root.after(0, lambda: self.current_cron_job_label.config(text=f"|   Currently Playing: {cron_job['name']}"))
        self.play_cron_button.config(text="Stop")
        for action in cron_job['actions']:
            if not self.replaying:
                    break
            action_type, action_detail, delay = action
            time.sleep(delay)
            if action_type == 'click':
                x, y = action_detail
                pyautogui.click(x, y)
            elif action_type == 'key':
                # Check if action detail is a single key or not a special key
                if str(action_detail).lower() in [str(key).lower() for key in self.special_keys]:
                    pyautogui.press(action_detail)
                else:
                    pyautogui.write(action_detail)
        
        self.root.after(0, self.reset_play_button)
        self.root.after(0, lambda: self.current_cron_job_label.config(text="|   Currently Playing: None"))

    def reset_play_button(self):
        self.play_cron_button.config(text="Play", state=tk.NORMAL)

    def stop_cron_job_replay(self):
        self.replaying = False
        self.play_cron_button.config(text="Play", command=self.play_cron_job)

    def import_cron_jobs(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                with open(file_path, 'r') as file:
                    data = json.load(file)
                    self.cron_jobs = []
                    for item in data:
                        cron_job = {
                            'name': item['name'],
                            'actions': item['actions'],
                            'cron_expression': item['cron_expression'],
                            'time': item.get('time', '12:00 AM'),  # Default time if not present
                            'active': item.get('active', True),
                            'last_executed': item.get('last_executed', '-')
                        }
                        self.cron_jobs.append(cron_job)
                    self.update_cron_jobs_list()
                messagebox.showinfo("Import Successful", "Cron jobs imported successfully.")
            except Exception as e:
                messagebox.showerror("Import Error", f"Failed to import cron jobs: {str(e)}")
                logging.error(f"Cron job import error: {str(e)}")

    def save_cron_jobs(self):
        if not self.cron_jobs:
            messagebox.showwarning("No Cron Jobs", "There are no cron jobs to save.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                with open(file_path, 'w') as file:
                    json.dump(self.cron_jobs, file, indent=2)
                messagebox.showinfo("Save Successful", "Cron jobs saved successfully.")
            except Exception as e:
                messagebox.showerror("Save Error", f"Failed to save cron jobs: {str(e)}")
                logging.error(f"Cron job save error: {str(e)}")

    def clear_cron_jobs(self):
        if messagebox.askyesno("Clear Cron Jobs", "Are you sure you want to clear all cron jobs?"):
            self.cron_jobs.clear()
            self.update_cron_jobs_list()

    def edit_cron_job(self):
        selected = self.cron_jobs_tree.selection()
        if selected:
            item = selected[0]
            original_index = int(self.cron_jobs_tree.item(item, 'tags')[0])  # Get the original index from the tag
            if 0 <= original_index < len(self.cron_jobs):
                cron_job = self.cron_jobs[original_index]
                dialog = EditCronJobDialog(self.root, cron_job)
                if dialog.result:
                    self.cron_jobs[original_index] = dialog.result
                    self.update_cron_jobs_list()
            else:
                messagebox.showerror("Error", "Invalid cron job index.")

    def delete_cron_job(self):
        selected = self.cron_jobs_tree.selection()
        if selected:
            item = selected[0]
            original_index = int(self.cron_jobs_tree.item(item, 'tags')[0])  # Get the original index from the tag
            del self.cron_jobs[original_index]
            self.update_cron_jobs_list()

    def toggle_cron_job_active(self):
        selected = self.cron_jobs_tree.selection()
        if selected:
            item = selected[0]
            original_index = int(self.cron_jobs_tree.item(item, 'tags')[0])  # Get the original index from the tag
            self.cron_jobs[original_index]['active'] = not self.cron_jobs[original_index]['active']
            self.update_cron_jobs_list()

    def add_to_cron_job(self):
        if self.current_list:
            time_str = simpledialog.askstring("Set Time", "Enter time (HH:MM AM/PM):", parent=self.root)
            if time_str:
                try:
                    # Convert 12-hour format to 24-hour format
                    time_obj = datetime.strptime(time_str, "%I:%M %p")
                    cron_expression = f"{time_obj.minute} {time_obj.hour} * * *"
                    cron_job = {
                        'name': self.current_list.name,
                        'actions': self.current_list.actions,
                        'cron_expression': cron_expression,
                        'time': time_str,  # Store the original time string
                        'active': True,
                        'last_executed': '-'
                    }
                    self.cron_jobs.append(cron_job)
                    self.update_cron_jobs_list()
                    messagebox.showinfo("Cron Job Added", f"Added cron job to run at: {time_str}")
                except ValueError:
                    messagebox.showerror("Invalid Input", "Please enter time in HH:MM AM/PM format")

    def update_cron_jobs_list(self):
        self.cron_jobs_tree.delete(*self.cron_jobs_tree.get_children())
        
        # Sort cron jobs by time
        sorted_jobs = sorted(enumerate(self.cron_jobs), key=lambda x: datetime.strptime(x[1]['time'], "%I:%M %p"))
        
        for i, (original_index, job) in enumerate(sorted_jobs):
            self.cron_jobs_tree.insert("", tk.END, values=(
                i+1, 
                job['name'], 
                len(job['actions']), 
                job['time'],
                "✓" if job['active'] else "✗",
                job.get('last_executed', '-')
            ), tags=(original_index,))  # Store the original index as a tag

    def check_and_execute_cron_jobs(self):
        current_time = datetime.now()
        for cron_job in self.cron_jobs:
            if cron_job['active']:
                job_time = datetime.strptime(cron_job['time'], "%I:%M %p").time()
                if current_time.time().hour == job_time.hour and current_time.time().minute == job_time.minute and current_time.time().second == 0:
                    # Execute cron job actions
                    action_list = ActionList(cron_job['name'])
                    action_list.actions = cron_job['actions']
                    self.replaying = True
                    threading.Thread(target=self.execute_action_list, args=(action_list,)).start()
                    cron_job['last_executed'] = current_time.strftime("%Y-%m-%d %H:%M:%S")
                    self.update_cron_jobs_list()  # Update the UI to reflect the last executed time

        # Schedule the next check in 1 second
        self.root.after(1000, self.check_and_execute_cron_jobs)

        # Update the current time display
        current_time_str = current_time.strftime("%I:%M:%S %p")
        self.current_time_label.config(text=f"Current Time: {current_time_str}")

        # Update the current time display
        def update_time_display():
            current_time_str = current_time.strftime("%I:%M:%S %p")
            self.current_time_label.config(text=f"Current Time: {current_time_str}")
        
        self.root.after(0, update_time_display)
        
    def execute_action_list(self, action_list):
        for action in action_list.actions:
            if not self.replaying:
                    break
            action_type, action_detail, delay = action
            time.sleep(delay)
            if action_type == 'click':
                x, y = action_detail
                pyautogui.click(x, y)
            elif action_type == 'key':
                    print(action_detail)
                    if str(action_detail).lower() in [str(key).lower() for key in self.special_keys]:
                        pyautogui.press(action_detail)
                    else:
                        pyautogui.write(action_detail)

    def create_settings_widgets(self):
        ttk.Label(self.settings_frame, text="Theme:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        ttk.Checkbutton(self.settings_frame, text="Dark Mode", variable=self.dark_mode, command=self.toggle_theme).grid(row=0, column=1, padx=10, pady=10, sticky="w")

        ttk.Label(self.settings_frame, text="Version: v1.0.10").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        ttk.Label(self.settings_frame, text="Last updated: 12/07/2024").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        ttk.Label(self.settings_frame, text="Credit: Concept By Nick Lim | Assisted by Sonet 3.5").grid(row=3, column=0, padx=10, pady=10, sticky="w")

    def toggle_theme(self):
        if self.dark_mode.get():
            self.style.theme_use('clam')
            self.style.configure('.', background='#333333', foreground='white')
            self.style.configure('TNotebook.Tab', background='#555555', foreground='white')
            self.style.map('TNotebook.Tab', background=[('selected', '#777777')])
            self.style.configure('Treeview', background='#333333', fieldbackground='#333333', foreground='white')
            self.style.map('TEntry', fieldbackground=[('!disabled', '#555555')], foreground=[('!disabled', 'white')])
            self.style.map('TCombobox', fieldbackground=[('!disabled', '#555555')], foreground=[('!disabled', 'white')])
        else:
            self.style.theme_use('clam')
            self.style.configure('.', background='#f0f0f0', foreground='black')
            self.style.configure('TNotebook.Tab', background='#e1e1e1', foreground='black')
            self.style.map('TNotebook.Tab', background=[('selected', '#f0f0f0')])
            self.style.configure('Treeview', background='#ffffff', fieldbackground='#ffffff', foreground='black')

    def create_dialog_on_main_thread(self, dialog_func, *args, **kwargs):
        result = None
        def run_dialog():
            nonlocal result
            result = dialog_func(*args, parent=self.root, **kwargs)
        self.root.after(0, run_dialog)
        self.root.wait_window()
        return result

    def create_record_widgets(self):
        # Row 0: Top buttons
        top_button_frame = ttk.Frame(self.record_frame)
        top_button_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

        self.import_button = ttk.Button(top_button_frame, text="Import", command=self.import_actions, style='Big.TButton', width=10)
        self.import_button.pack(side=tk.LEFT, padx=(0, 5))

        self.export_button = ttk.Button(top_button_frame, text="Save", command=self.export_actions, state=tk.DISABLED, style='Big.TButton', width=10)
        self.export_button.pack(side=tk.LEFT, padx=5)

        self.record_button = ttk.Button(top_button_frame, text="Start Recording", command=self.start_recording, style='Big.TButton', width=12)
        self.record_button.pack(side=tk.LEFT, padx=5)

        # Row 1: Action Tree
        self.action_tree = ttk.Treeview(self.record_frame, columns=("Action", "Delay"), show='headings')
        self.action_tree.heading("Action", text="Action")
        self.action_tree.heading("Delay", text="Delay (seconds)")
        self.action_tree.grid(row=1, column=0, padx=(10, 0), pady=10, sticky="nsew")
        self.action_tree.bind('<Double-1>', self.edit_action_in_tree)

        # Button frame
        button_frame = ttk.Frame(self.record_frame, width=230)
        button_frame.grid(row=1, column=1, padx=(0, 10), pady=10, sticky="nsew")
        button_frame.grid_propagate(False)  # This prevents the frame from shrinking

        button_frame.grid_columnconfigure(0, weight=1)

        # Create a style for buttons in the button frame
        button_style = ttk.Style()
        button_style.configure('ButtonFrame.TButton', width=20)

        # Buttons in the button frame
        self.edit_action_button = ttk.Button(button_frame, text="Edit Action", command=self.edit_selected_action, state=tk.DISABLED, style='ButtonFrame.TButton')
        self.edit_action_button.grid(row=0, column=0, pady=5, padx=5, sticky="ew")

        self.delete_button = ttk.Button(button_frame, text="Delete", command=self.delete_action, state=tk.DISABLED, style='ButtonFrame.TButton')
        self.delete_button.grid(row=1, column=0, pady=5, padx=5, sticky="ew")

        self.clear_button = ttk.Button(button_frame, text="Clear", command=self.clear_actions, state=tk.DISABLED, style='ButtonFrame.TButton')
        self.clear_button.grid(row=2, column=0, pady=5, padx=5, sticky="ew")

        self.add_to_replay_button = ttk.Button(button_frame, text="Add to Replay", command=self.add_to_replay, state=tk.DISABLED, style='ButtonFrame.TButton')
        self.add_to_replay_button.grid(row=3, column=0, pady=5, padx=5, sticky="ew")

        self.add_to_cron_job_button = ttk.Button(button_frame, text="Add to Cron Job", command=self.add_to_cron_job, state=tk.DISABLED, style='ButtonFrame.TButton')
        self.add_to_cron_job_button.grid(row=4, column=0, pady=5, padx=5, sticky="ew")

        ttk.Separator(button_frame, orient='horizontal').grid(row=5, column=0, pady=10, sticky="ew")

        self.cycle_count_label = ttk.Label(button_frame, text="Currently in cycle: 0")
        self.cycle_count_label.grid(row=6, column=0, pady=5, padx=5, sticky="w")

        self.reset_cycle_button = ttk.Button(button_frame, text="Reset Cycles", command=self.reset_cycles, style='ButtonFrame.TButton')
        self.reset_cycle_button.grid(row=7, column=0, pady=5, padx=5, sticky="ew")

        self.repeat_var = tk.BooleanVar()
        self.repeat_checkbox = ttk.Checkbutton(button_frame, text="Repeat", variable=self.repeat_var)
        self.repeat_checkbox.grid(row=8, column=0, pady=5, padx=5, sticky="w")

        repeat_frame = ttk.Frame(button_frame)
        repeat_frame.grid(row=9, column=0, pady=5, padx=5, sticky="ew")
        ttk.Label(repeat_frame, text="Repeat for:").pack(side=tk.LEFT)
        self.repeat_count_var = tk.StringVar(value="0")
        self.repeat_count_entry = ttk.Entry(repeat_frame, textvariable=self.repeat_count_var, width=10)
        self.repeat_count_entry.pack(side=tk.LEFT, padx=(5, 0))

        self.play_button = ttk.Button(button_frame, text="Play", command=self.play_recording, state=tk.DISABLED, style='ButtonFrame.TButton')
        self.play_button.grid(row=10, column=0, pady=5, padx=5, sticky="ew")

        # Add this near the end of the method
        self.check_coords_var = tk.BooleanVar()
        self.check_coords_button = ttk.Checkbutton(
            button_frame, 
            text="Check Coordinates", 
            variable=self.check_coords_var, 
            command=self.toggle_coord_check
        )
        self.check_coords_button.grid(row=11, column=0, pady=5, padx=5, sticky="ew")

        # Configure grid weights
        self.record_frame.grid_columnconfigure(0, weight=1)
        self.record_frame.grid_columnconfigure(1, weight=0)
        self.record_frame.grid_rowconfigure(1, weight=1)

        # Add a horizontal scrollbar
        h_scroll = ttk.Scrollbar(self.record_frame, orient="horizontal", command=self.action_tree.xview)
        h_scroll.grid(row=2, column=0, columnspan=2, sticky="ew")
        self.action_tree.configure(xscrollcommand=h_scroll.set)

        # Bind selection event to enable/disable buttons
        self.action_tree.bind('<<TreeviewSelect>>', self.on_action_select)

    def toggle_coord_check(self):
        if self.check_coords_var.get():
            # self.root.withdraw()  # Hide the main window
            self.coord_listener = mouse.Listener(on_click=self.on_coord_check)
            self.coord_listener.start()
        else:
            if hasattr(self, 'coord_listener'):
                self.coord_listener.stop()

    def on_coord_check(self, x, y, button, pressed):
        if pressed:
            self.root.clipboard_clear()
            self.root.clipboard_append(f"({x}, {y})")
            messagebox.showinfo("Coordinates", f"Coordinates ({x}, {y}) copied to clipboard!")
            self.check_coords_var.set(False)  # Uncheck the button
            self.coord_listener.stop()
            self.root.deiconify()  # Show the main window again
            return False  # Stop listener

    def copy_coords_to_clipboard(self, x, y):
        self.root.clipboard_clear()
        self.root.clipboard_append(f"({x}, {y})")
        messagebox.showinfo("Coordinates", f"Coordinates ({x}, {y}) copied to clipboard!")

    def on_action_select(self, event):
        selected = self.action_tree.selection()
        if selected:
            self.edit_action_button.config(state=tk.NORMAL)
        else:
            self.edit_action_button.config(state=tk.DISABLED)

    def reset_cycles(self):
        self.action_repeat_count = 0
        self.cycle_count_label.config(text="Cycles: 0")

    def create_replay_widgets(self):
        # Row 0 (New row for Import, Save, and Clear buttons)
        replay_control_frame = ttk.Frame(self.replay_frame)
        replay_control_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

        self.import_replay_button = ttk.Button(replay_control_frame, text="Import", command=self.import_replay, style='Big.TButton', width=10)
        self.import_replay_button.pack(side=tk.LEFT, padx=(0, 5))

        self.combine_import_button = ttk.Button(replay_control_frame, text="Combine New Import", command=self.combine_new_import, state=tk.DISABLED, style='Big.TButton', width=15)
        self.combine_import_button.pack(side=tk.LEFT, padx=5)

        self.save_replay_button = ttk.Button(replay_control_frame, text="Save", command=self.save_replay, style='Big.TButton', width=10)
        self.save_replay_button.pack(side=tk.LEFT, padx=5)

        self.clear_replay_button = ttk.Button(replay_control_frame, text="Clear", command=self.clear_replay, style='Big.TButton', width=10)
        self.clear_replay_button.pack(side=tk.LEFT, padx=5)

        # Row 1 (Replay Tree and Button Frame)
        self.replay_tree = ttk.Treeview(self.replay_frame, columns=("Sequence", "Name", "Actions", "Repeat", "Interval", "Executed", "Active", "Last Executed"), show='headings')
        self.replay_tree.heading("Sequence", text="No.")
        self.replay_tree.heading("Name", text="Name")
        self.replay_tree.heading("Actions", text="Actions")
        self.replay_tree.heading("Repeat", text="Repeat")
        self.replay_tree.heading("Interval", text="Interval (min)")
        self.replay_tree.heading("Executed", text="Executed")
        self.replay_tree.heading("Active", text="Active")
        self.replay_tree.heading("Last Executed", text="Last Executed")
        self.replay_tree.column("Last Executed", width=150, minwidth=150, anchor='center')
        self.replay_tree.grid(row=1, column=0, padx=(10, 0), pady=10, sticky="nsew")
        self.replay_tree.bind('<ButtonRelease-1>', self.on_replay_tree_click)
        self.replay_tree.bind('<Double-1>', self.on_replay_tree_double_click)

        # Set column widths
        self.replay_tree.column("Sequence", width=30, minwidth=30, anchor='center')
        self.replay_tree.column("Name", width=100, minwidth=100, anchor='w')
        self.replay_tree.column("Actions", width=50, minwidth=50, anchor='center')
        self.replay_tree.column("Repeat", width=50, minwidth=50, anchor='center')
        self.replay_tree.column("Interval", width=80, minwidth=80, anchor='center')
        self.replay_tree.column("Executed", width=60, minwidth=60, anchor='center')
        self.replay_tree.column("Active", width=50, minwidth=50, anchor='center')

        # Button frame
        button_frame = ttk.Frame(self.replay_frame, width=230)
        button_frame.grid(row=1, column=1, padx=(0, 10), pady=10, sticky="nsew")
        button_frame.grid_propagate(False)  # This prevents the frame from shrinking

        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)

        # Create a style for buttons in the button frame
        button_style = ttk.Style()
        button_style.configure('ButtonFrame.TButton', width=20)

        # Buttons in the button frame
        self.move_up_button = ttk.Button(button_frame, text="Move Up", command=self.move_up, state=tk.DISABLED, style='ButtonFrame.TButton')
        self.move_up_button.grid(row=0, column=0, columnspan=2, pady=5, padx=5, sticky="ew")

        self.move_down_button = ttk.Button(button_frame, text="Move Down", command=self.move_down, state=tk.DISABLED, style='ButtonFrame.TButton')
        self.move_down_button.grid(row=1, column=0, columnspan=2, pady=5, padx=5, sticky="ew")

        self.edit_replay_button = ttk.Button(button_frame, text="Edit Replay", command=self.edit_selected_replay, state=tk.DISABLED, style='ButtonFrame.TButton')
        self.edit_replay_button.grid(row=2, column=0, columnspan=2, pady=5, padx=5, sticky="ew")

        self.delete_replay_button = ttk.Button(button_frame, text="Delete", command=self.delete_replay_item, state=tk.DISABLED, style='ButtonFrame.TButton')
        self.delete_replay_button.grid(row=3, column=0, columnspan=2, pady=5, padx=5, sticky="ew")

        self.duplicate_replay_button = ttk.Button(button_frame, text="Duplicate", command=self.duplicate_replay_item, state=tk.DISABLED, style='ButtonFrame.TButton')
        self.duplicate_replay_button.grid(row=4, column=0, columnspan=2, pady=5, padx=5, sticky="ew")

        self.activate_button = ttk.Button(button_frame, text="Activate/Deactivate", command=self.toggle_active, state=tk.DISABLED, style='ButtonFrame.TButton')
        self.activate_button.grid(row=5, column=0, columnspan=2, pady=5, padx=5, sticky="ew")

        ttk.Separator(button_frame, orient='horizontal').grid(row=6, column=0, columnspan=2, pady=5, sticky="ew")

        self.add_to_record_button = ttk.Button(button_frame, text="Add to Record", command=self.add_to_record, state=tk.DISABLED, style='ButtonFrame.TButton')
        self.add_to_record_button.grid(row=7, column=0, columnspan=2, pady=5, padx=5, sticky="ew")

        ttk.Separator(button_frame, orient='horizontal').grid(row=8, column=0, columnspan=2, pady=5, sticky="ew")

        self.status_label = ttk.Label(button_frame, text="Status: Stopped")
        self.status_label.grid(row=9, column=0, columnspan=2, pady=5, padx=5, sticky="w")

        self.current_item_label = ttk.Label(button_frame, text="N/A")
        self.current_item_label.grid(row=10, column=0, columnspan=2, pady=5, padx=5, sticky="w")

        self.cycle_label = ttk.Label(button_frame, text="Cycle: 0")
        self.cycle_label.grid(row=11, column=0, columnspan=2, pady=5, padx=5, sticky="w")

        self.full_cycle_label = ttk.Label(button_frame, text="Full Cycles: 0")
        self.full_cycle_label.grid(row=12, column=0, columnspan=2, pady=5, padx=5, sticky="w")

        self.repeat_all_var = tk.BooleanVar()
        self.repeat_all_checkbox = ttk.Checkbutton(button_frame, text="Repeat All", variable=self.repeat_all_var)
        self.repeat_all_checkbox.grid(row=13, column=0, columnspan=2, pady=5, padx=5, sticky="w")

        self.start_button = ttk.Button(button_frame, text="Start", command=self.start_replay, style='ButtonFrame.TButton')
        self.start_button.grid(row=14, column=0, columnspan=2, pady=5, padx=5, sticky="ew")

        self.pause_button = ttk.Button(button_frame, text="Pause", command=self.pause_replay, state=tk.DISABLED, style='ButtonFrame.TButton')
        self.pause_button.grid(row=15, column=0, pady=5, padx=5, sticky="ew")

        self.stop_button = ttk.Button(button_frame, text="Stop", command=self.stop_replay, state=tk.DISABLED, style='ButtonFrame.TButton')
        self.stop_button.grid(row=15, column=1, pady=5, padx=5, sticky="ew")

        # Configure grid weights
        self.replay_frame.grid_columnconfigure(0, weight=1)
        self.replay_frame.grid_columnconfigure(1, weight=0)
        self.replay_frame.grid_rowconfigure(1, weight=1)

        # Add a horizontal scrollbar
        h_scroll = ttk.Scrollbar(self.replay_frame, orient="horizontal", command=self.replay_tree.xview)
        h_scroll.grid(row=2, column=0, columnspan=2, sticky="ew")
        self.replay_tree.configure(xscrollcommand=h_scroll.set)

    def combine_new_import(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                with open(file_path, 'r') as file:
                    new_data = json.load(file)
                    for item in new_data:
                        action_list = ActionList(item['name'], item['sequence'], item.get('interval', 0))
                        action_list.actions = item['actions']
                        action_list.repeat = item['repeat']
                        action_list.active = item.get('active', True)
                        action_list.last_executed = "-"
                        self.action_lists.append(action_list)
                    self.update_replay_list()
                messagebox.showinfo("Import Successful", "Replay actions combined successfully.")
            except Exception as e:
                messagebox.showerror("Import Error", f"Failed to import and combine replay actions: {str(e)}")
                logging.error(f"Import and combine error: {str(e)}")

    def on_replay_tree_click(self, event):
        self.on_replay_select(event)

    def on_replay_tree_double_click(self, event):
        region = self.replay_tree.identify("region", event.x, event.y)
        if region == "cell":
            column = self.replay_tree.identify_column(event.x)
            if column == "#7":  # The 'Active' column
                self.toggle_active()
            else:
                self.edit_replay_in_tree(event)

    def on_replay_select(self, event):
        selected = self.replay_tree.selection()
        if selected:
            self.delete_replay_button.config(state=tk.NORMAL)
            self.edit_replay_button.config(state=tk.NORMAL)
            self.add_to_record_button.config(state=tk.NORMAL)
            self.activate_button.config(state=tk.NORMAL)
            self.duplicate_replay_button.config(state=tk.NORMAL)
        else:
            self.delete_replay_button.config(state=tk.DISABLED)
            self.edit_replay_button.config(state=tk.DISABLED)
            self.add_to_record_button.config(state=tk.DISABLED)
            self.activate_button.config(state=tk.DISABLED)
            self.duplicate_replay_button.config(state=tk.DISABLED)
        self.update_move_buttons()

    def duplicate_replay_item(self):
        selected = self.replay_tree.selection()
        if selected:
            index = self.replay_tree.index(selected[0])
            original_item = self.action_lists[index]
            new_item = ActionList(f"{original_item.name} (Copy)", original_item.sequence, original_item.interval, original_item.active)
            new_item.actions = original_item.actions.copy()
            new_item.repeat = original_item.repeat
            self.action_lists.insert(index + 1, new_item)
            self.update_replay_list()
            self.replay_tree.selection_set(self.replay_tree.get_children()[index + 1])

    def clear_replay(self):
        if messagebox.askyesno("Clear Replay", "Are you sure you want to clear all replay actions?"):
            self.action_lists.clear()
            self.update_replay_list()

    def add_to_record(self):
        selected = self.replay_tree.selection()
        if selected:
            index = self.replay_tree.index(selected[0])
            action_list = self.action_lists[index]
            self.current_list = ActionList(action_list.name)
            self.current_list.actions = action_list.actions.copy()
            self.update_action_list()
            self.notebook.select(0)  # Switch to the Record tab

    def update_move_buttons(self):
        selected = self.replay_tree.selection()
        if selected:
            index = self.replay_tree.index(selected[0])
            self.move_up_button.config(state=tk.NORMAL if index > 0 else tk.DISABLED)
            self.move_down_button.config(state=tk.NORMAL if index < len(self.action_lists) - 1 else tk.DISABLED)
        else:
            self.move_up_button.config(state=tk.DISABLED)
            self.move_down_button.config(state=tk.DISABLED)

    def move_up(self):
        selected = self.replay_tree.selection()
        if selected:
            index = self.replay_tree.index(selected[0])
            if index > 0:
                self.action_lists[index], self.action_lists[index-1] = self.action_lists[index-1], self.action_lists[index]
                self.action_lists[index].sequence, self.action_lists[index-1].sequence = self.action_lists[index-1].sequence, self.action_lists[index].sequence
                self.update_replay_list()
                self.replay_tree.selection_set(self.replay_tree.get_children()[index-1])

    def move_down(self):
        selected = self.replay_tree.selection()
        if selected:
            index = self.replay_tree.index(selected[0])
            if index < len(self.action_lists) - 1:
                self.action_lists[index], self.action_lists[index+1] = self.action_lists[index+1], self.action_lists[index]
                self.action_lists[index].sequence, self.action_lists[index+1].sequence = self.action_lists[index+1].sequence, self.action_lists[index].sequence
                self.update_replay_list()
                self.replay_tree.selection_set(self.replay_tree.get_children()[index+1])


    def add_edit_buttons(self):
        self.edit_action_button = ttk.Button(self.record_frame, text="Edit Action", command=self.edit_selected_action, style='Big.TButton', width=10)
        self.edit_action_button.grid(row=2, column=3, padx=10, pady=10, sticky="nsew")

        self.edit_replay_button = ttk.Button(self.replay_frame, text="Edit Replay", command=self.edit_selected_replay, style='Big.TButton', width=10)
        self.edit_replay_button.grid(row=1, column=4, padx=10, pady=10, sticky="nsew")

    @debounce(0.3)
    def start_recording(self):
        if not self.recording:
            self.recording = True
            self.current_list = ActionList(f"Recording_{len(self.action_lists) + 1}")
            self.last_time = time.time()
            self.record_button.config(text="Stop Recording")
            self.status_label.config(text="Status: Recording")
        else:
            self.stop_recording()

    def stop_recording(self):
        self.recording = False
        self.record_button.config(text="Record")
        self.status_label.config(text="Status: Stopped")
        self.root.after(0, self.update_action_list)
        self.action_lists.append(self.current_list)
        self.root.after(0, self.update_replay_list)

    def on_click(self, x, y, button, pressed):
        if self.recording and pressed and not self.check_coords_var.get():
            self.root.after(0, self.add_click_action, x, y)

    def add_click_action(self, x, y):
        current_time = time.time()
        delay = current_time - self.last_time
        self.last_time = current_time
        self.current_list.add_action(('click', (x, y), delay))
        self.root.after(0, self.update_action_list)

    def on_press(self, key):
        try:
            if key == keyboard.Key.esc:
                if self.recording:
                    self.stop_recording()
                elif self.replaying:
                    self.stop_replay()
                elif hasattr(self, 'play_cron_button') and self.play_cron_button['text'] == "Stop":
                    self.stop_cron_job_replay()
            elif key == keyboard.Key.home:
                self.start_replay()
            elif key == keyboard.Key.end:
                self.pause_replay()
        except AttributeError:
            pass

        if self.recording:
            self.root.after(0, self.add_key_action, key)

    def add_key_action(self, key):
        current_time = time.time()
        delay = current_time - self.last_time
        self.last_time = current_time
        key_str = self.key_to_string(key)
        self.current_list.add_action(('key', key_str, delay))
        self.root.after(0, self.update_action_list)

    def key_to_string(self, key):
        # Convert key to string for PyAutoGUI
        try:
            if key.char:
                return key.char
        except AttributeError:
            special_keys = {
                keyboard.Key.space: 'space',
                keyboard.Key.enter: 'enter',
                keyboard.Key.shift: 'shift',
                keyboard.Key.ctrl: 'ctrl',
                keyboard.Key.alt: 'alt',
                keyboard.Key.tab: 'tab',
                keyboard.Key.backspace: 'backspace',
                keyboard.Key.delete: 'delete',
                keyboard.Key.esc: 'esc',
                keyboard.Key.up: 'up',
                keyboard.Key.down: 'down',
                keyboard.Key.left: 'left',
                keyboard.Key.right: 'right',
                keyboard.Key.home: 'home',
                keyboard.Key.end: 'end',
                keyboard.Key.page_up: 'pageup',
                keyboard.Key.page_down: 'pagedown'
            }
            return special_keys.get(key, str(key).replace('Key.', ''))

    def update_action_list(self):
        self.action_tree.delete(*self.action_tree.get_children())
        for action in self.current_list.actions:
            action_type, action_detail, delay = action
            if action_type == 'click':
                action_str = f"Click at {action_detail}"
            elif action_type == 'key':
                action_str = f"Key {action_detail}"
            self.action_tree.insert("", tk.END, values=(action_str, f"{delay:.2f}"))

        if self.current_list.actions:
            self.export_button.config(state=tk.NORMAL)
            self.delete_button.config(state=tk.NORMAL)
            self.clear_button.config(state=tk.NORMAL)
            self.add_to_replay_button.config(state=tk.NORMAL)
            self.add_to_cron_job_button.config(state=tk.NORMAL)
            self.play_button.config(state=tk.NORMAL)
        else:
            self.export_button.config(state=tk.DISABLED)
            self.delete_button.config(state=tk.DISABLED)
            self.clear_button.config(state=tk.DISABLED)
            self.add_to_replay_button.config(state=tk.DISABLED)
            self.add_to_cron_job_button.config(state=tk.DISABLED)
            self.play_button.config(state=tk.DISABLED)

    def update_replay_list(self):
        self.replay_tree.delete(*self.replay_tree.get_children())
        for i, action_list in enumerate(self.action_lists):
            action_list.sequence = i  # Update sequence
            item = self.replay_tree.insert("", tk.END, values=(
                action_list.sequence, 
                action_list.name, 
                len(action_list.actions), 
                action_list.repeat, 
                action_list.interval,
                getattr(action_list, 'executed', 0),
                "✓" if action_list.active else "✗"
            ))
            self.replay_tree.item(item, tags=('active' if action_list.active else 'inactive',))
        
        self.replay_tree.tag_configure('active', background='lightgreen')
        self.replay_tree.tag_configure('inactive', background='lightgray')
        
        # Reset button states
        self.delete_replay_button.config(state=tk.DISABLED)
        self.edit_replay_button.config(state=tk.DISABLED)
        self.add_to_record_button.config(state=tk.DISABLED)
        self.activate_button.config(state=tk.DISABLED)
        
        self.update_move_buttons()

    def toggle_active(self):
        selected = self.replay_tree.selection()
        if selected:
            item = selected[0]
            index = self.replay_tree.index(item)
            action_list = self.action_lists[index]
            action_list.active = not action_list.active
            self.update_replay_list()

    def edit_action_in_tree(self, event):
        selected = self.action_tree.selection()
        if not selected:
            return

        item = selected[0]
        index = self.action_tree.index(item)
        
        if event:
            column = self.action_tree.identify_column(event.x)
        else:
            column = '#1'
        
        current_value = self.action_tree.set(item, column)
        
        if column == '#1':  # Action column
            new_value = simpledialog.askstring("Edit Action", "Enter new action:", initialvalue=current_value, parent=self.root)
            if new_value:
                try:
                    action_type, action_detail = self.parse_action_string(new_value)
                    delay = self.current_list.actions[index][2]
                    self.current_list.actions[index] = (action_type, action_detail, delay)
                    self.action_tree.set(item, column, new_value)
                    print(f"Action updated: {self.current_list.actions[index]}")
                except ValueError as e:
                    messagebox.showerror("Invalid Input", str(e), parent=self.root)
        elif column == '#2':  # Delay column
            new_value = simpledialog.askfloat("Edit Delay", "Enter new delay:", initialvalue=float(current_value), parent=self.root)
            if new_value is not None:
                action_type, action_detail = self.current_list.actions[index][:2]
                self.current_list.actions[index] = (action_type, action_detail, new_value)
                self.action_tree.set(item, column, f"{new_value:.2f}")
                print(f"Delay updated: {self.current_list.actions[index]}")

        self.update_action_list()

    def parse_action_string(self, action_str):
        if action_str.startswith("Click at "):
            try:
                coords = eval(action_str.replace("Click at ", ""))
                if isinstance(coords, tuple) and len(coords) == 2:
                    return 'click', coords
                else:
                    raise ValueError
            except:
                raise ValueError("Invalid click coordinates")
        elif action_str.startswith("Key "):
            return 'key', action_str.replace("Key ", "")
        else:
            raise ValueError("Invalid action string format")

    def edit_replay_in_tree(self, event):
        selected = self.replay_tree.selection()
        if not selected:
            return
        item = selected[0]
        index = self.replay_tree.index(item)
        action_list = self.action_lists[index]
        
        # Schedule the dialog creation on the main GUI thread
        self.root.after(0, self.create_edit_dialog, action_list)

    def create_edit_dialog(self, action_list):
        dialog = EditReplayDialog(self.root, action_list)
        if dialog.result:
            action_list.name, action_list.repeat, action_list.sequence, action_list.interval, action_list.active = dialog.result
            self.update_replay_list()

    @debounce(0.3)
    def delete_action(self):
        selected = self.action_tree.selection()
        if selected:
            index = self.action_tree.index(selected[0])
            self.current_list.remove_action(index)
            self.update_action_list()

    @debounce(0.3)
    def clear_actions(self):
        self.current_list.clear_actions()
        self.update_action_list()

    @debounce(0.3)
    def add_to_replay(self):
        if self.current_list:
            new_sequence = max([al.sequence for al in self.action_lists], default=-1) + 1
            self.current_list.sequence = new_sequence
            self.action_lists.append(self.current_list)
            self.update_replay_list()
            self.current_list = None
            self.clear_actions()

    @debounce(0.3)
    def import_actions(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                with open(file_path, 'r') as file:
                    data = json.load(file)
                    self.current_list = ActionList(data['name'])
                    self.current_list.actions = data['actions']
                    self.update_action_list()
            except Exception as e:
                messagebox.showerror("Import Error", f"Failed to import actions: {str(e)}")
                logging.error(f"Import error: {str(e)}")

    @debounce(0.3)
    def export_actions(self):
        if not self.current_list or not self.current_list.actions:
            messagebox.showwarning("No Actions", "There are no actions to export.")
            return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                data = {
                    'name': self.current_list.name,
                    'actions': self.current_list.actions
                }
                with open(file_path, 'w') as file:
                    json.dump(data, file, indent=2)
                messagebox.showinfo("Export Successful", "Actions exported successfully.")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export actions: {str(e)}")
                logging.error(f"Export error: {str(e)}")

    def play_recording(self):
        if not self.current_list or not self.current_list.actions:
            messagebox.showwarning("No Actions", "There are no actions to play.")
            return

        self.replaying = True
        self.paused = False
        self.play_button.config(text="Stop", command=self.stop_recording_playback)
        threading.Thread(target=self.execute_recording_playback).start()

    def stop_recording_playback(self):
        self.replaying = False
        self.play_button.config(text="Play", command=self.play_recording)

    def execute_recording_playback(self):
        repeat_count = int(self.repeat_count_var.get()) if self.repeat_var.get() else 1
        cycles_completed = 0
        
        while self.replaying and (repeat_count == 0 or cycles_completed < repeat_count):
            for action in self.current_list.actions:
                if not self.replaying:
                    break
                action_type, action_detail, delay = action
                time.sleep(delay)
                if action_type == 'click':
                    x, y = action_detail
                    pyautogui.click(x, y)
                elif action_type == 'key':
                    if str(action_detail).lower() in [str(key).lower() for key in self.special_keys]:
                        pyautogui.press(action_detail)
                    else:
                        pyautogui.write(action_detail)

            cycles_completed += 1
            if repeat_count > 0:
                self.root.after(0, lambda: self.cycle_count_label.config(text=f"Cycles: {cycles_completed}/{repeat_count}"))
            else:
                self.root.after(0, lambda: self.cycle_count_label.config(text=f"Cycles: {cycles_completed} (Infinite)"))

        self.root.after(0, self.stop_recording_playback)

    @debounce(0.3)
    def start_replay(self):
        if not self.action_lists:
            messagebox.showwarning("No Actions", "There are no action lists to replay.")
            return

        self.replaying = True
        self.paused = False
        self.status_label.config(text="Status: Running")
        self.start_button.config(state=tk.DISABLED)
        self.pause_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.NORMAL)

        # Reset the 'executed' count and 'last_executed' time for all action lists
        for action_list in self.action_lists:
            action_list.executed = 0
            action_list.last_executed = 0
        self.update_replay_list()

        threading.Thread(target=self.execute_replay).start()

    def execute_replay(self):
        full_cycles = 0
        while self.replaying:
            for action_list in sorted(self.action_lists, key=lambda x: x.sequence):
                if not self.replaying:
                    break
                
                if not action_list.active:
                    continue  # Skip inactive items
                
                current_time = time.time()
                
                if (action_list.executed == 0 or 
                    action_list.interval == 0 or 
                    current_time >= action_list.last_executed + (action_list.interval * 60)):
                    
                    self.action_repeat_count = 0
                    while self.replaying and (self.action_repeat_count < action_list.repeat or action_list.repeat == 0):
                        for i, action in enumerate(action_list.actions):
                            if not self.replaying:
                                break
                            while self.paused:
                                time.sleep(0.1)
                            action_type, action_detail, delay = action
                            time.sleep(delay)
                            if action_type == 'click':
                                x, y = action_detail
                                pyautogui.click(x, y)
                            elif action_type == 'key':
                                # Check if action detail is a single key or not a special key
                                if str(action_detail).lower() in [str(key).lower() for key in self.special_keys]:
                                    pyautogui.press(action_detail)
                                else:
                                    pyautogui.write(action_detail)
                            self.root.after(0, self.update_replay_status, action_list.name, i+1, len(action_list.actions))
                        self.action_repeat_count += 1
                        self.root.after(0, self.update_cycle_count)
                    
                    action_list.last_executed = time.time()
                    action_list.executed += 1
                    self.root.after(0, self.update_replay_list)
                else:
                    print(f"Skipping {action_list.name} due to interval not met")
            
            full_cycles += 1
            self.root.after(0, self.update_full_cycle_count, full_cycles)
            
            if not self.repeat_all_var.get():
                break

        self.root.after(0, self.stop_replay)

    def update_full_cycle_count(self, count):
        self.full_cycle_label.config(text=f"Full Cycles: {count}")

    def update_replay_status(self, name, current, total):
        self.current_item_label.config(text=f" {name}, {current}/{total} [{self.action_repeat_count}]")

    def update_cycle_count(self):
        self.cycle_count_label.config(text=f"Cycles: {self.action_repeat_count}")

    @debounce(0.3)
    def pause_replay(self):
        if self.replaying:
            if not self.paused:
                self.paused = True
                self.status_label.config(text="Status: Paused")
                self.pause_button.config(text="Resume")
            else:
                self.paused = False
                self.status_label.config(text="Status: Running")
                self.pause_button.config(text="Pause")

    def stop_replay(self):
        self.replaying = False
        self.paused = False
        self.status_label.config(text="Status: Stopped")
        self.start_button.config(state=tk.NORMAL)
        self.pause_button.config(state=tk.DISABLED, text="Pause")
        self.stop_button.config(state=tk.DISABLED)
        self.current_item_label.config(text="N/A")
        self.cycle_label.config(text="Cycle: 0")
        # Note: We're not resetting the 'executed' count here

    def close_program(self):
        if messagebox.askokcancel("Quit", "Do you want to quit?", parent=self.root):
            self.root.quit()
            self.root.destroy()
            os._exit(0)  # Force exit

    @debounce(0.3)
    def edit_selected_action(self):
        selected = self.action_tree.selection()
        if selected:
            self.edit_action_in_tree(None)

    @debounce(0.3)
    def edit_selected_replay(self):
        selected = self.replay_tree.selection()
        if selected:
            self.edit_replay_in_tree(None)

    @debounce(0.3)
    def delete_replay_item(self):
        selected = self.replay_tree.selection()
        if selected:
            index = self.replay_tree.index(selected[0])
            del self.action_lists[index]
            self.update_replay_list()

    def save_replay(self):
        if not self.action_lists:
            messagebox.showwarning("No Actions", "There are no action lists to save.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                data = []
                for action_list in self.action_lists:
                    data.append({
                        'name': action_list.name,
                        'actions': action_list.actions,
                        'repeat': action_list.repeat,
                        'sequence': action_list.sequence,
                        'interval': action_list.interval,
                        'active': action_list.active
                    })
                with open(file_path, 'w') as file:
                    json.dump(data, file, indent=2)
                messagebox.showinfo("Save Successful", "Replay list saved successfully.")
            except Exception as e:
                messagebox.showerror("Save Error", f"Failed to save replay list: {str(e)}")
                logging.error(f"Save error: {str(e)}")

    def import_replay(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                with open(file_path, 'r') as file:
                    data = json.load(file)
                    if not isinstance(data, list):
                        raise ValueError("Invalid file format: expected a list of action lists")
                
                    new_action_lists = []
                    for item in data:
                        if not isinstance(item, dict):
                            raise ValueError("Invalid item format: expected a dictionary")
                        
                        name = item.get('name')
                        sequence = item.get('sequence')
                        interval = item.get('interval', 0)
                        actions = item.get('actions')
                        repeat = item.get('repeat')
                        
                        if not all(isinstance(x, (str, int)) for x in [name, sequence, interval, repeat]):
                            raise ValueError("Invalid data types for action list properties")
                        if not isinstance(actions, list):
                            raise ValueError("Invalid actions format: expected a list")
                        
                        action_list = ActionList(name, sequence, interval)
                        action_list.actions = actions
                        action_list.repeat = repeat
                        action_list.active = item.get('active', True)
                        new_action_lists.append(action_list)
                    
                    self.action_lists = new_action_lists
                    self.update_replay_list()
                messagebox.showinfo("Import Successful", "Replay list imported successfully.")
            except json.JSONDecodeError:
                messagebox.showerror("Import Error", "Invalid JSON file")
                logging.error("Import error: Invalid JSON file")
            except ValueError as e:
                messagebox.showerror("Import Error", str(e))
                logging.error(f"Import error: {str(e)}")
            except Exception as e:
                messagebox.showerror("Import Error", f"An unexpected error occurred: {str(e)}")
                logging.error(f"Import error: {str(e)}")

class EditReplayDialog(simpledialog.Dialog):
    def __init__(self, parent, action_list):
        self.action_list = action_list
        self.result = None
        super().__init__(parent, title="Edit Replay Item")

    def body(self, master):
        ttk.Label(master, text="Name:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.name_entry = ttk.Entry(master)
        self.name_entry.grid(row=0, column=1, padx=5, pady=5)
        self.name_entry.insert(0, self.action_list.name)

        ttk.Label(master, text="Repeat:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.repeat_entry = ttk.Entry(master)
        self.repeat_entry.grid(row=1, column=1, padx=5, pady=5)
        self.repeat_entry.insert(0, str(self.action_list.repeat))

        ttk.Label(master, text="Sequence:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.sequence_entry = ttk.Entry(master)
        self.sequence_entry.grid(row=2, column=1, padx=5, pady=5)
        self.sequence_entry.insert(0, str(self.action_list.sequence))

        ttk.Label(master, text="Interval (min):").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.interval_entry = ttk.Entry(master)
        self.interval_entry.grid(row=3, column=1, padx=5, pady=5)
        self.interval_entry.insert(0, str(self.action_list.interval))

        ttk.Label(master, text="Active:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.active_var = tk.BooleanVar(value=self.action_list.active)
        self.active_checkbox = ttk.Checkbutton(master, variable=self.active_var)
        self.active_checkbox.grid(row=4, column=1, padx=5, pady=5)

        self.actions_tree = ttk.Treeview(master, columns=("Action", "Delay"), show='headings')
        self.actions_tree.heading("Action", text="Action")
        self.actions_tree.heading("Delay", text="Delay (seconds)")
        self.actions_tree.grid(row=5, column=0, columnspan=2, padx=5, pady=5)

        for action in self.action_list.actions:
            action_type, action_detail, delay = action
            if action_type == 'click':
                action_str = f"Click at {action_detail}"
            elif action_type == 'key':
                action_str = f"Key {action_detail}"
            self.actions_tree.insert("", tk.END, values=(action_str, f"{delay:.2f}"))

        self.actions_tree.bind('<Double-1>', self.edit_action)

        button_frame = ttk.Frame(master)
        button_frame.grid(row=6, column=0, columnspan=2, pady=5)

        self.delete_action_button = ttk.Button(button_frame, text="Delete Action", command=self.delete_action)
        self.delete_action_button.pack(side=tk.LEFT, padx=5)

        self.duplicate_action_button = ttk.Button(button_frame, text="Duplicate Action", command=self.duplicate_action)
        self.duplicate_action_button.pack(side=tk.LEFT, padx=5)

        self.set_action_button = ttk.Button(button_frame, text="Set", command=self.set_action)
        self.set_action_button.pack(side=tk.LEFT, padx=5)

        return self.name_entry
    
    def set_action(self):
        def on_action(x, y, button, pressed):
            if pressed:
                self.new_action = ('click', (x, y), 0)
                listener.stop()

        def on_press(key):
            self.new_action = ('key', self.key_to_string(key), 0)
            listener.stop()

        self.new_action = None
        with mouse.Listener(on_click=on_action) as listener:
            with keyboard.Listener(on_press=on_press) as k_listener:
                listener.join()
                k_listener.join()

        if self.new_action:
            selected = self.actions_tree.selection()
            if selected:
                item = selected[0]
                action_type, action_detail, _ = self.new_action
                if action_type == 'click':
                    action_str = f"Click at {action_detail}"
                elif action_type == 'key':
                    action_str = f"Key {action_detail}"
                self.actions_tree.item(item, values=(action_str, "0.00"))

    def show(self):
        self.root.after(100, self._show)  # Add a small delay before showing the dialog

    def _show(self):
        try:
            self.wait_visibility()
            self.grab_set()
            self.wait_window(self)
        except tk.TclError:
            pass  # Ignore the TclError if the window was deleted

    def delete_action(self):
        selected = self.actions_tree.selection()
        if selected:
            self.actions_tree.delete(selected[0])

    def duplicate_action(self):
        selected = self.actions_tree.selection()
        if selected:
            item = selected[0]
            values = self.actions_tree.item(item, 'values')
            self.actions_tree.insert("", self.actions_tree.index(item) + 1, values=values)

    def edit_action(self, event):
        item = self.actions_tree.selection()[0]
        column = self.actions_tree.identify_column(event.x)
        
        if column == '#1':  # Action column
            current_value = self.actions_tree.set(item, column)
            new_value = simpledialog.askstring("Edit Action", "Enter new action:", initialvalue=current_value)
            if new_value:
                self.actions_tree.set(item, column, new_value)
        elif column == '#2':  # Delay column
            current_value = self.actions_tree.set(item, column)
            new_value = simpledialog.askfloat("Edit Delay", "Enter new delay:", initialvalue=float(current_value))
            if new_value is not None:
                self.actions_tree.set(item, column, f"{new_value:.2f}")

    def apply(self):
        try:
            name = self.name_entry.get()
            repeat = int(self.repeat_entry.get())
            sequence = int(self.sequence_entry.get())
            interval = int(self.interval_entry.get())
            active = self.active_var.get()
            
            if repeat < 0 or sequence < 0 or interval < 0:
                raise ValueError("Values must be non-negative integers.")
            
            # Update the action_list with the new values
            self.action_list.name = name
            self.action_list.repeat = repeat
            self.action_list.sequence = sequence
            self.action_list.interval = interval
            self.action_list.active = active
            
            # Update actions from the tree
            self.action_list.actions.clear()
            for item in self.actions_tree.get_children():
                action_str, delay_str = self.actions_tree.item(item, 'values')
                if action_str.startswith("Click at "):
                    action_type = 'click'
                    action_detail = eval(action_str.replace("Click at ", ""))
                elif action_str.startswith("Key "):
                    action_type = 'key'
                    action_detail = action_str.replace("Key ", "")
                delay = float(delay_str)
                self.action_list.actions.append((action_type, action_detail, delay))
            
            self.result = (name, repeat, sequence, interval, active)
        except ValueError as e:
            messagebox.showerror("Invalid Input", str(e))
            self.result = None

class EditCronJobDialog(simpledialog.Dialog):
    def __init__(self, parent, cron_job):
        self.cron_job = cron_job
        self.result = None
        super().__init__(parent, title="Edit Cron Job")

    def body(self, master):
        ttk.Label(master, text="Name:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.name_entry = ttk.Entry(master)
        self.name_entry.grid(row=0, column=1, padx=5, pady=5)
        self.name_entry.insert(0, self.cron_job['name'])

        ttk.Label(master, text="Time (HH:MM AM/PM):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.time_entry = ttk.Entry(master)
        self.time_entry.grid(row=1, column=1, padx=5, pady=5)
        self.time_entry.insert(0, self.cron_job['time'])

        ttk.Label(master, text="Active:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.active_var = tk.BooleanVar(value=self.cron_job['active'])
        self.active_checkbox = ttk.Checkbutton(master, variable=self.active_var)
        self.active_checkbox.grid(row=2, column=1, padx=5, pady=5)

        self.actions_tree = ttk.Treeview(master, columns=("Action", "Delay"), show='headings')
        self.actions_tree.heading("Action", text="Action")
        self.actions_tree.heading("Delay", text="Delay (seconds)")
        self.actions_tree.grid(row=3, column=0, columnspan=2, padx=5, pady=5)

        for action in self.cron_job['actions']:
            action_type, action_detail, delay = action
            if action_type == 'click':
                action_str = f"Click at {action_detail}"
            elif action_type == 'key':
                action_str = f"Key {action_detail}"
            self.actions_tree.insert("", tk.END, values=(action_str, f"{delay:.2f}"))

        self.actions_tree.bind('<Double-1>', self.edit_action)

        button_frame = ttk.Frame(master)
        button_frame.grid(row=4, column=0, columnspan=2, pady=5)

        self.delete_action_button = ttk.Button(button_frame, text="Delete Action", command=self.delete_action)
        self.delete_action_button.pack(side=tk.LEFT, padx=5)

        self.duplicate_action_button = ttk.Button(button_frame, text="Duplicate Action", command=self.duplicate_action)
        self.duplicate_action_button.pack(side=tk.LEFT, padx=5)

        self.add_action_button = ttk.Button(button_frame, text="Add Action", command=self.add_action)
        self.add_action_button.pack(side=tk.LEFT, padx=5)

        return self.name_entry

    def edit_action(self, event):
        item = self.actions_tree.selection()[0]
        column = self.actions_tree.identify_column(event.x)
        
        if column == '#1':  # Action column
            current_value = self.actions_tree.set(item, column)
            new_value = simpledialog.askstring("Edit Action", "Enter new action:", initialvalue=current_value)
            if new_value:
                self.actions_tree.set(item, column, new_value)
        elif column == '#2':  # Delay column
            current_value = self.actions_tree.set(item, column)
            new_value = simpledialog.askfloat("Edit Delay", "Enter new delay:", initialvalue=float(current_value))
            if new_value is not None:
                self.actions_tree.set(item, column, f"{new_value:.2f}")

    def delete_action(self):
        selected = self.actions_tree.selection()
        if selected:
            self.actions_tree.delete(selected[0])

    def duplicate_action(self):
        selected = self.actions_tree.selection()
        if selected:
            item = selected[0]
            values = self.actions_tree.item(item, 'values')
            self.actions_tree.insert("", self.actions_tree.index(item) + 1, values=values)

    def add_action(self):
        action_type = simpledialog.askstring("Add Action", "Enter action type (click/key):")
        if action_type in ['click', 'key']:
            if action_type == 'click':
                x = simpledialog.askinteger("Add Click Action", "Enter X coordinate:")
                y = simpledialog.askinteger("Add Click Action", "Enter Y coordinate:")
                if x is not None and y is not None:
                    action_str = f"Click at ({x}, {y})"
            else:
                key = simpledialog.askstring("Add Key Action", "Enter key:")
                if key:
                    action_str = f"Key {key}"
            
            delay = simpledialog.askfloat("Add Action", "Enter delay (seconds):", initialvalue=0.0)
            if delay is not None:
                self.actions_tree.insert("", tk.END, values=(action_str, f"{delay:.2f}"))

    def validate(self):
        try:
            name = self.name_entry.get().strip()
            time_str = self.time_entry.get().strip()
            
            if not name or not time_str:
                raise ValueError("Name and Time cannot be empty.")
            
            # Validate time format
            datetime.strptime(time_str, "%I:%M %p")
            
            return True
        except ValueError as e:
            messagebox.showerror("Invalid Input", str(e))
            return False

    def apply(self):
        name = self.name_entry.get().strip()
        time_str = self.time_entry.get().strip()
        active = self.active_var.get()
        
        # Convert time to cron expression
        time_obj = datetime.strptime(time_str, "%I:%M %p")
        cron_expression = f"{time_obj.minute} {time_obj.hour} * * *"
        
        # Update actions from the tree
        actions = []
        for item in self.actions_tree.get_children():
            action_str, delay_str = self.actions_tree.item(item, 'values')
            if action_str.startswith("Click at "):
                action_type = 'click'
                action_detail = eval(action_str.replace("Click at ", ""))
            elif action_str.startswith("Key "):
                action_type = 'key'
                action_detail = action_str.replace("Key ", "")
            delay = float(delay_str)
            actions.append((action_type, action_detail, delay))
        
        self.result = {
            'name': name,
            'cron_expression': cron_expression,
            'time': time_str,
            'active': active,
            'actions': actions
        }

def main():
    root = tk.Tk()
    app = None
    try:
        app = ActionRecorder(root)
        root.protocol("WM_DELETE_WINDOW", app.close_program)
        root.mainloop()
    except Exception as e:
        print(f"An error occurred: {e}")
        if app:
            app.close_program()
        else:
            root.quit()
            root.destroy()
    finally:
        print("Application closed.")

if __name__ == "__main__":
    main()
