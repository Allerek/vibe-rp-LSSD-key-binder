import tkinter as tk
from tkinter import ttk
from pynput import keyboard
from pynput.keyboard import Controller
import json
import os
from pathlib import Path
import sys
import ctypes
import win32gui
import win32con

def hide_console():
    # Hide console window
    console_window = win32gui.GetForegroundWindow()
    win32gui.ShowWindow(console_window, win32con.SW_HIDE)

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    # Re-run the program with admin rights
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    sys.exit()

# Hide console window on startup
hide_console()

class KeyBinderGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Switcher kanałów LSSD")
        self.root.geometry("300x300")
        
        self.keyboard_controller = Controller()
        self.key_buttons = []
        self.current_binding = None
        self.bindings = {}
        
        # Pre-defined commands table
        self.command_table = [
            {"name": "A-TAC-1", "command": "t /rc 511"},
            {"name": "A-TAC-2", "command": "t /rc 512"},
            {"name": "A-TAC-3", "command": "t /rc 513"},
            {"name": "M-AID-1", "command": "t /rc 510"},
            {"name": "M-AID-2", "command": "t /rc 520"},
            {"name": "M-AID-3", "command": "t /rc 530"},
            {"name": "Powrót", "command": "t /rc 692 "}
        ]
        
        # Create save directory if it doesn't exist
        self.save_dir = Path(os.getenv('APPDATA')) / 'lssdswitch'
        self.save_dir.mkdir(parents=True, exist_ok=True)
        self.save_file = self.save_dir / 'keybinds.json'
        
        # Load saved bindings if they exist
        self.load_bindings()
        self.create_widgets()
        
        # Re-initialize bindings after loading
        for i, btn in enumerate(self.key_buttons):
            if i in self.bindings:
                key = self.bindings[i]["key"]
                key_str = self.get_key_string(key)
                btn.config(text=key_str)
        
    def create_widgets(self):
        # Create header labels
        header_frame = ttk.Frame(self.root)
        header_frame.pack(pady=5, padx=10, fill='x')
        ttk.Label(header_frame, text="Kanał", width=15).pack(side='left')
        ttk.Label(header_frame, text="Klawisz", width=20).pack(side='left', padx=5)
        
        # Create key binding rows based on command table
        for i, command_data in enumerate(self.command_table):
            frame = ttk.Frame(self.root)
            frame.pack(pady=5, padx=10, fill='x')
            
            # Command name label
            name_label = ttk.Label(frame, text=command_data["name"], width=15)
            name_label.pack(side='left')
            
            # Bind button
            btn = ttk.Button(frame, text="Wciśnij aby zbindować", width=20)
            btn.config(command=lambda b=btn, idx=i: self.start_binding(b, idx))
            btn.pack(side='left', padx=5)
            self.key_buttons.append(btn)
            
    def start_binding(self, button, idx):
        if self.current_binding is not None:
            return
            
        button.config(text="Wciśnij klawisz...")
        self.current_binding = (button, idx)
        
        # Start key listener
        self.listener = keyboard.Listener(on_press=self.on_key_press)
        self.listener.start()
        
    def get_key_string(self, key):
        key_str = str(key)
        if hasattr(key, 'vk'):  # Regular keys
            if key.vk == 96:  # Numpad 0
                key_str = "Numpad 0"
            elif 97 <= key.vk <= 105:  # Numpad 1-9
                key_str = f"Numpad {key.vk - 96}"
        elif hasattr(key, 'name'):  # Special keys
            key_str = key.name.title()
        return key_str
        
    def on_key_press(self, key):
        if self.current_binding is None:
            return
            
        button, idx = self.current_binding
        
        # Convert key to string representation
        key_str = self.get_key_string(key)
            
        # Update button text and save binding
        button.config(text=key_str)
        self.bindings[idx] = {
            "key": key,
            "command": self.command_table[idx]["command"]
        }
        
        # Save bindings to file
        self.save_bindings()
        
        # Stop listening
        self.current_binding = None
        self.listener.stop()
        
    def save_bindings(self):
        # Convert bindings to serializable format
        save_data = {}
        for idx, binding in self.bindings.items():
            key = binding["key"]
            key_data = {
                "vk": key.vk if hasattr(key, "vk") else None,
                "char": key.char if hasattr(key, "char") else None,
                "name": key.name if hasattr(key, "name") else None
            }
            save_data[str(idx)] = {
                "key": key_data,
                "command": binding["command"]
            }
            
        # Save to file
        with open(self.save_file, 'w') as f:
            json.dump(save_data, f)
            
    def load_bindings(self):
        if not self.save_file.exists():
            return
            
        try:
            with open(self.save_file, 'r') as f:
                save_data = json.load(f)
                
            for idx_str, binding_data in save_data.items():
                idx = int(idx_str)
                key_data = binding_data["key"]
                
                # Reconstruct key object
                if key_data["vk"] is not None:
                    key = keyboard.KeyCode(vk=key_data["vk"], char=key_data["char"])
                elif key_data["char"] is not None:
                    key = keyboard.KeyCode.from_char(key_data["char"])
                else:
                    key = keyboard.Key[key_data["name"]]
                    
                self.bindings[idx] = {
                    "key": key,
                    "command": binding_data["command"]
                }
        except Exception as e:
            print(f"Error loading keybinds: {e}")
        
    def type_command(self, command):
        # Type the command and press enter
        for char in command:
            self.keyboard_controller.press(char)
            self.keyboard_controller.release(char)
        self.keyboard_controller.press(keyboard.Key.enter)
        self.keyboard_controller.release(keyboard.Key.enter)
        
    def run(self):
        # Start listener for executing commands
        def on_press(key):
            for idx, binding in self.bindings.items():
                if isinstance(key, keyboard.KeyCode) and isinstance(binding["key"], keyboard.KeyCode):
                    if key.vk == binding["key"].vk:
                        self.type_command(binding["command"])
                    
        self.command_listener = keyboard.Listener(on_press=on_press)
        self.command_listener.start()
        
        self.root.mainloop()
        
    def get_bindings(self):
        return self.bindings

if __name__ == "__main__":
    binder = KeyBinderGUI()
    binder.run()
