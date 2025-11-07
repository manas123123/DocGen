import win32com.client
import time
import tkinter as tk
from tkinter import messagebox, scrolledtext
import pandas as pd
import os
from tkinter import filedialog
import openpyxl
from tkinter import ttk
import sys
import shutil
from datetime import datetime

# Hide console window on Windows
if sys.platform == "win32":
    try:
        import ctypes
        # Hide the console window
        kernel32 = ctypes.windll.kernel32
        user32 = ctypes.windll.user32
        # Get the console window handle
        hwnd = kernel32.GetConsoleWindow()
        if hwnd:
            # Hide the console window
            user32.ShowWindow(hwnd, 0)  # 0 = SW_HIDE
    except Exception:
        # If hiding fails, continue anyway
        pass

# Try to import tkcalendar, fallback if not available
try:
    from tkcalendar import Calendar
    CALENDAR_AVAILABLE = True
except ImportError:
    CALENDAR_AVAILABLE = False

# Try to import tkinterdnd2, fallback if not available
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False
    TkinterDnD = tk.Tk  # Fallback to regular Tk
    DND_FILES = None


class RoundedFrame(tk.Canvas):
    """A modern frame with smooth rounded corners and optional shadow"""
    def __init__(self, parent, bg_color, radius=15, border_color=None, border_width=0, 
                 shadow=False, shadow_color='#000000', shadow_offset=2, **kwargs):
        # Get parent's background color to match
        try:
            parent_bg = parent.cget('bg')
        except:
            parent_bg = '#1e1e1e'  # Default dark theme background
        
        # Set canvas background to match parent
        kwargs['bg'] = parent_bg
        tk.Canvas.__init__(self, parent, highlightthickness=0, **kwargs)
        self.config(bg=parent_bg)
        self.radius = max(0, radius)  # Ensure non-negative
        self.bg_color = bg_color
        self.border_color = border_color or bg_color
        self.border_width = border_width
        self.shadow = shadow
        self.shadow_color = shadow_color
        self.shadow_offset = shadow_offset
        self.bind("<Configure>", self._draw)
        
    def _draw(self, event=None):
        self.delete("all")
        width = self.winfo_width()
        height = self.winfo_height()
        print(f"Canvas width: {width}, height: {height}")
        if width > 1 and height > 1:
            max_radius = min(width, height) / 2
            radius = min(self.radius, max_radius)

        # Draw a background rectangle first to fill any corner gaps 👇
            self.create_rectangle(0, 0, width, height, fill=self['bg'], outline='')

        # Draw shadow if enabled
            if self.shadow:
                shadow_x1 = width - self.shadow_offset
                shadow_y1 = height - self.shadow_offset
                shadow_x2 = width
                shadow_y2 = height
                shadow_fill = '#0a0a0a'
                self.create_rounded_rectangle_polygon(
                    shadow_x1, shadow_y1, shadow_x2, shadow_y2, radius,
                    fill=shadow_fill, outline=''
                )

        # Draw main rounded rectangle
            self.create_rounded_rectangle_polygon(
                0, 0, width, height, radius,
                fill=self.bg_color,
                outline=self.border_color,
                width=self.border_width
            )
    
    def create_rounded_rectangle_polygon(self, x1, y1, x2, y2, radius, **kwargs):
        """Create a smooth rounded rectangle with consistent radius and clean rendering"""
        fill_color = kwargs.pop('fill', '')
        outline_color = kwargs.pop('outline', '')
        width = kwargs.pop('width', 1)
        alpha = kwargs.pop('alpha', 1.0)
        
        # Ensure radius is valid and consistent
        radius = min(radius, min((x2 - x1) / 2, (y2 - y1) / 2))
        radius = max(0, radius)  # Ensure non-negative
        
        # Fill the rounded rectangle with smooth corners
        # Use a polygon approach for better corner coverage
        # Draw center rectangles first with slight overlap into corner areas
        # Horizontal center strip (extends slightly into corners for overlap)
        self.create_rectangle(x1 - 1, y1 + radius, x2 + 1, y2 - radius,
                             fill=fill_color, outline='', width=0)
        # Vertical center strip (extends slightly into corners for overlap)
        self.create_rectangle(x1 + radius, y1 - 1, x2 - radius, y2 + 1,
                             fill=fill_color, outline='', width=0)
        
        # Draw corner arcs on top to create rounded corners
        # Top-left corner arc (quarter circle)
        self.create_arc(x1, y1, x1 + radius*2, y1 + radius*2,
                       start=90, extent=90, fill=fill_color, outline='', style='pieslice')
        # Top-right corner arc
        self.create_arc(x2 - radius*2, y1, x2, y1 + radius*2,
                       start=0, extent=90, fill=fill_color, outline='', style='pieslice')
        # Bottom-right corner arc
        self.create_arc(x2 - radius*2, y2 - radius*2, x2, y2,
                       start=270, extent=90, fill=fill_color, outline='', style='pieslice')
        # Bottom-left corner arc
        self.create_arc(x1, y2 - radius*2, x1 + radius*2, y2,
                       start=180, extent=90, fill=fill_color, outline='', style='pieslice')
        
        # Draw border outline with smooth, consistent corners
        if outline_color and width > 0:
            # Use arc style for smooth corner borders
            # Top-left corner border
            self.create_arc(x1, y1, x1 + radius*2, y1 + radius*2,
                          start=90, extent=90, outline=outline_color, width=width, style='arc')
            # Top-right corner border
            self.create_arc(x2 - radius*2, y1, x2, y1 + radius*2,
                          start=0, extent=90, outline=outline_color, width=width, style='arc')
            # Bottom-right corner border
            self.create_arc(x2 - radius*2, y2 - radius*2, x2, y2,
                          start=270, extent=90, outline=outline_color, width=width, style='arc')
            # Bottom-left corner border
            self.create_arc(x1, y2 - radius*2, x1 + radius*2, y2,
                          start=180, extent=90, outline=outline_color, width=width, style='arc')
            
            # Draw straight edges with proper alignment
            border_half = width / 2
            self.create_line(x1 + radius, y1 + border_half, x2 - radius, y1 + border_half,
                           fill=outline_color, width=width, capstyle=tk.ROUND)
            self.create_line(x2 - border_half, y1 + radius, x2 - border_half, y2 - radius,
                           fill=outline_color, width=width, capstyle=tk.ROUND)
            self.create_line(x2 - radius, y2 - border_half, x1 + radius, y2 - border_half,
                           fill=outline_color, width=width, capstyle=tk.ROUND)
            self.create_line(x1 + border_half, y2 - radius, x1 + border_half, y1 + radius,
                           fill=outline_color, width=width, capstyle=tk.ROUND)


class DocumentGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Document Generator")
        self.root.geometry("750x850")
        self.root.configure(bg="#1e1e1e")
        
        # Variables
        self.df = None
        self.file_path = None
        self.output_directory = None
        self.today = datetime.now()
        
        # Date variables - will be set from calendar
        self.selected_date = self.today
        
        # Status text widget
        self.status_text = None
        self.calendar_widget = None
        self.date_entry = None  # For fallback when calendar not available
        
        self.setup_ui()
    
    def create_rounded_button(self, parent, text, command, bg_color="#3a4a5c", fg_color="#e0e0e0", 
                             active_bg="#4a5a6c", radius=12, **kwargs):
        """Create a button with rounded corners"""
        # Get parent background for the rounded frame
        try:
            parent_bg = parent.cget('bg')
        except:
            parent_bg = '#2d2d2d'  # Default frame background
        
        btn_frame = RoundedFrame(parent, bg_color=bg_color, radius=radius, border_width=0)
        btn_frame.pack(**kwargs)
        
        btn = tk.Button(
            btn_frame,
            text=text,
            command=command,
            bg=bg_color,
            fg=fg_color,
            activebackground=active_bg,
            activeforeground=fg_color,
            relief=tk.FLAT,
            borderwidth=0,
            cursor="hand2",
            highlightthickness=0
        )
        btn.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        # Update button frame on hover
        def on_enter(e):
            btn_frame.bg_color = active_bg
            btn.config(bg=active_bg)
            btn_frame._draw()
        
        def on_leave(e):
            btn_frame.bg_color = bg_color
            btn.config(bg=bg_color)
            btn_frame._draw()
        
        btn_frame.bind("<Enter>", on_enter)
        btn_frame.bind("<Leave>", on_leave)
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        
        return btn_frame
        
    def setup_ui(self):
        # Configure ttk style for dark theme
        style = ttk.Style()
        style.theme_use('default')
        style.configure("TProgressbar",
                       background='#3a4a5c',
                       troughcolor='#2d2d2d',
                       borderwidth=0,
                       lightcolor='#4a5a6c',
                       darkcolor='#4a5a6c')
        
        # Main container with padding
        main_frame = tk.Frame(self.root, bg="#1e1e1e", padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = tk.Label(
            main_frame, 
            text="Document Generator", 
            font=("Arial", 24, "bold"),
            bg="#1e1e1e",
            fg="#e0e0e0"
        )
        title_label.pack(pady=(0, 30))
        
        # Horizontal container for Step 1 and Step 2
        steps_container = tk.Frame(main_frame, bg="#1e1e1e")
        steps_container.pack(fill=tk.BOTH, expand=True, pady=(0, 15), padx=5)
        steps_container.grid_columnconfigure(0, weight=50)  # Browse Files: 75%
        steps_container.grid_columnconfigure(1, weight=50)   # Date Selection: 25%
        steps_container.grid_rowconfigure(0, weight=1)
        
        # Step 1: File Selection Frame - Modern card style with rounded corners and soft shadow
        file_frame_container = RoundedFrame(steps_container, bg_color="#2d2d2d", radius=15, 
                                           border_color="#3a3a3a", border_width=1,
                                           shadow=True, shadow_color="#000000", shadow_offset=3)
        file_frame_container.grid(row=0, column=0, sticky="nsew", padx=(0, 7.5))
        # Ensure the canvas expands properly
        file_frame_container.config(width=1, height=1)  # Reset to allow proper expansion
        
        file_frame = tk.Frame(file_frame_container, bg="#2d2d2d")
        file_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Title label for the frame
        frame_title = tk.Label(
            file_frame,
            text="Step 1: Select Master Sheet",
            font=("Arial", 12, "bold"),
            bg="#2d2d2d",
            fg="#d0d0d0"
        )
        frame_title.pack(anchor=tk.W, pady=(0, 15))
        
        # Enhanced drag and drop area - Modern card style with rounded corners
        drop_container = RoundedFrame(file_frame, bg_color="#3a4a5c", radius=12, border_color="#4a5a6c", border_width=2, cursor="hand2")
        drop_container.pack(fill=tk.BOTH, expand=True, pady=10, padx=5)
        
        # Make entire container clickable
        drop_container.bind("<Button-1>", lambda e: self.browse_file())
        
        # Create hover handlers that will update both container and label
        def on_drop_enter(e):
            drop_container.bg_color = "#4a5a6c"
            drop_container.border_color = "#5a6a7c"
            drop_container._draw()
            # Update label background if it exists
            if hasattr(drop_container, '_drop_label'):
                drop_container._drop_label.config(bg="#4a5a6c")
        def on_drop_leave(e):
            drop_container.bg_color = "#3a4a5c"
            drop_container.border_color = "#4a5a6c"
            drop_container._draw()
            # Update label background if it exists
            if hasattr(drop_container, '_drop_label'):
                drop_container._drop_label.config(bg="#3a4a5c")
        drop_container.bind("<Enter>", on_drop_enter)
        drop_container.bind("<Leave>", on_drop_leave)
        
        if DND_AVAILABLE:
            # Register drop target on container
            drop_container.drop_target_register(DND_FILES)
            drop_container.dnd_bind('<<Drop>>', self.on_file_drop)
            
            drop_label = tk.Label(
                drop_container,
                text="📁 Browse Files",
                font=("Arial", 14, "bold"),
                bg="#3a4a5c",
                fg="#e0e0e0",
                cursor="hand2"
            )
            drop_label.pack(fill=tk.BOTH, expand=True, padx=40, pady=50)
            # Store reference to label in container for hover handlers
            drop_container._drop_label = drop_label
            
            # Make label clickable and also register drag-drop for redundancy
            drop_label.drop_target_register(DND_FILES)
            drop_label.dnd_bind('<<Drop>>', self.on_file_drop)
            drop_label.bind("<Button-1>", lambda e: self.browse_file())
            drop_label.bind("<Enter>", lambda e: (drop_label.config(bg="#4a5a6c"), on_drop_enter(e)))
            drop_label.bind("<Leave>", lambda e: (drop_label.config(bg="#3a4a5c"), on_drop_leave(e)))
        else:
            drop_label = tk.Label(
                drop_container,
                text="📁 Browse Files",
                font=("Arial", 14, "bold"),
                bg="#3a4a5c",
                fg="#e0e0e0",
                cursor="hand2"
            )
            drop_label.pack(fill=tk.BOTH, expand=True, padx=40, pady=50)
            # Store reference to label in container for hover handlers
            drop_container._drop_label = drop_label
            drop_label.bind("<Button-1>", lambda e: self.browse_file())
            drop_label.bind("<Enter>", lambda e: (drop_label.config(bg="#4a5a6c"), on_drop_enter(e)))
            drop_label.bind("<Leave>", lambda e: (drop_label.config(bg="#3a4a5c"), on_drop_leave(e)))
        
        # Browse button - Modern rounded button
        browse_btn_frame = self.create_rounded_button(
            file_frame,
            text="Browse Files",
            command=self.browse_file,
            bg_color="#3a4a5c",
            fg_color="#e0e0e0",
            active_bg="#4a5a6c",
            radius=10,
            pady=8
        )
        browse_btn = browse_btn_frame.winfo_children()[0]  # Get the actual button
        browse_btn.config(font=("Arial", 10, "bold"), padx=25, pady=10)
        
        # Selected file label
        self.file_label = tk.Label(
            file_frame,
            text="No file selected",
            font=("Arial", 9),
            bg="#2d2d2d",
            fg="#a0a0a0",
            wraplength=250
        )
        self.file_label.pack(pady=5)
        
        # Step 2: Date Selection Frame - Modern card style with rounded corners and soft shadow
        date_frame_container = RoundedFrame(steps_container, bg_color="#2d2d2d", radius=15,
                                           border_color="#3a3a3a", border_width=1,
                                           shadow=True, shadow_color="#000000", shadow_offset=3)
        date_frame_container.grid(row=0, column=1, sticky="nsew", padx=(7.5, 0))
        # Ensure the canvas expands properly
        date_frame_container.config(width=1, height=1)  # Reset to allow proper expansion
        
        date_frame = tk.Frame(date_frame_container, bg="#2d2d2d")
        date_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Title label for the frame
        date_frame_title = tk.Label(
            date_frame,
            text="Step 2: Select Date",
            font=("Arial", 12, "bold"),
            bg="#2d2d2d",
            fg="#d0d0d0"
        )
        date_frame_title.pack(anchor=tk.W, pady=(0, 15))
        
        # Calendar widget container - centered
        calendar_container = tk.Frame(date_frame, bg="#2d2d2d")
        calendar_container.pack(fill=tk.BOTH, expand=True, pady=10)
        
        if CALENDAR_AVAILABLE:
            # Use tkcalendar widget with nice styling
            self.calendar_widget = Calendar(
                calendar_container,
                selectmode='day',
                year=self.today.year,
                month=self.today.month,
                day=self.today.day,
                date_pattern='yyyy-mm-dd',
                font=("Arial", 11),
                background='#2d2d2d',
                foreground='#e0e0e0',
                selectbackground='#3a4a5c',
                selectforeground='white',
                normalbackground='#2d2d2d',
                normalforeground='#e0e0e0',
                weekendbackground='#353535',
                weekendforeground='#e0e0e0',
                headersbackground='#1a1a1a',
                headersforeground='#e0e0e0',
                borderwidth=2,
                relief=tk.RAISED,
                cursor="hand2"
            )
            self.calendar_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        else:
            # Fallback to simple date entry if calendar not available
            fallback_frame = tk.Frame(calendar_container, bg="#2d2d2d")
            fallback_frame.pack()
            
            tk.Label(fallback_frame, text="Date (YYYY-MM-DD):", font=("Arial", 10), bg="#2d2d2d", fg="#e0e0e0").pack(side=tk.LEFT, padx=5)
            self.date_entry = tk.Entry(fallback_frame, font=("Arial", 10), width=15, bg="#353535", fg="#e0e0e0", insertbackground="#e0e0e0")
            self.date_entry.pack(side=tk.LEFT, padx=5)
            self.date_entry.insert(0, self.today.strftime("%Y-%m-%d"))
            
            tk.Label(
                fallback_frame,
                text="⚠ tkcalendar not installed. Install with: pip install tkcalendar",
                font=("Arial", 8),
                bg="#2d2d2d",
                fg="#e74c3c"
            ).pack(side=tk.LEFT, padx=10)
        
        # Generate button - Modern rounded button
        self.generate_btn_frame = self.create_rounded_button(
            main_frame,
            text="Generate Documents",
            command=self.start_generation,
            bg_color="#3a4a5c",
            fg_color="#e0e0e0",
            active_bg="#4a5a6c",
            radius=12,
            pady=20
        )
        self.generate_btn = self.generate_btn_frame.winfo_children()[0]  # Get the actual button
        self.generate_btn.config(font=("Arial", 12, "bold"), padx=35, pady=14, state=tk.DISABLED)
        
        # Step 3: Status and Progress Frame - Modern card style with rounded corners and soft shadow
        status_frame_container = RoundedFrame(main_frame, bg_color="#2d2d2d", radius=15,
                                             border_color="#3a3a3a", border_width=1,
                                             shadow=True, shadow_color="#000000", shadow_offset=3)
        status_frame_container.pack(fill=tk.BOTH, expand=True, pady=(0, 10), padx=5)
        
        status_frame = tk.Frame(status_frame_container, bg="#2d2d2d")
        status_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Title label for the frame
        status_frame_title = tk.Label(
            status_frame,
            text="Status & Progress",
            font=("Arial", 12, "bold"),
            bg="#2d2d2d",
            fg="#d0d0d0"
        )
        status_frame_title.pack(anchor=tk.W, pady=(0, 15))
        
        # Progress bar with label
        progress_container = tk.Frame(status_frame, bg="#2d2d2d")
        progress_container.pack(fill=tk.X, pady=(0, 10))
        
        self.progress_label = tk.Label(
            progress_container,
            text="Ready",
            font=("Arial", 10),
            bg="#2d2d2d",
            fg="#a0a0a0"
        )
        self.progress_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_container,
            variable=self.progress_var,
            maximum=100,
            length=500,
            mode='determinate'
        )
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.progress_percent = tk.Label(
            progress_container,
            text="0%",
            font=("Arial", 10, "bold"),
            bg="#2d2d2d",
            fg="#d0d0d0",
            width=5
        )
        self.progress_percent.pack(side=tk.RIGHT, padx=(10, 0))
        
        # Modern status log area with better styling
        log_container = tk.Frame(status_frame, bg="#252525", relief=tk.SUNKEN, borderwidth=1)
        log_container.pack(fill=tk.BOTH, expand=True)
        
        # Header for log area
        log_header = tk.Frame(log_container, bg="#3a3a3a", height=30)
        log_header.pack(fill=tk.X)
        log_header.pack_propagate(False)
        
        tk.Label(
            log_header,
            text="Activity Log",
            font=("Arial", 9, "bold"),
            bg="#3a3a3a",
            fg="#c0c0c0"
        ).pack(side=tk.LEFT, padx=10, pady=5)
        
        # Status text area with modern styling
        self.status_text = scrolledtext.ScrolledText(
            log_container,
            height=8,
            font=("Segoe UI", 9),
            bg="#2d2d2d",
            fg="#e0e0e0",
            wrap=tk.WORD,
            relief=tk.FLAT,
            borderwidth=0,
            padx=10,
            pady=8,
            selectbackground="#3a4a5c",
            selectforeground="white",
            insertbackground="#e0e0e0"
        )
        self.status_text.pack(fill=tk.BOTH, expand=True, padx=1, pady=(0, 1))
        self.status_text.config(state=tk.DISABLED)
        
        # Configure tags for different message types
        self.status_text.tag_config("info", foreground="#a0a0a0", font=("Segoe UI", 9))
        self.status_text.tag_config("success", foreground="#4caf50", font=("Segoe UI", 9, "bold"))
        self.status_text.tag_config("error", foreground="#e74c3c", font=("Segoe UI", 9, "bold"))
        self.status_text.tag_config("warning", foreground="#f39c12", font=("Segoe UI", 9))
        self.status_text.tag_config("timestamp", foreground="#888888", font=("Segoe UI", 8))
        
        self.log_status("Welcome! Please select a master sheet file to begin.", "info")
        
    def log_status(self, message, msg_type="info"):
        """Add a message to the status text area with modern styling"""
        if self.status_text:
            self.status_text.config(state=tk.NORMAL)
            timestamp = datetime.now().strftime("%H:%M:%S")
            
            # Determine icon and tag based on message type
            icons = {
                "info": "ℹ",
                "success": "✓",
                "error": "✗",
                "warning": "⚠"
            }
            icon = icons.get(msg_type, "•")
            
            # Insert timestamp
            self.status_text.insert(tk.END, f"[{timestamp}] ", "timestamp")
            
            # Insert icon and message with appropriate tag
            self.status_text.insert(tk.END, f"{icon} ", msg_type)
            self.status_text.insert(tk.END, f"{message}\n", msg_type)
            
            self.status_text.see(tk.END)
            self.status_text.config(state=tk.DISABLED)
            self.root.update()
    
    def update_progress(self, value, label_text=None):
        """Update progress bar and label"""
        self.progress_var.set(value)
        if label_text:
            self.progress_label.config(text=label_text)
        self.progress_percent.config(text=f"{int(value)}%")
        self.root.update()
    
    def on_file_drop(self, event):
        """Handle file drop event"""
        files = self.root.tk.splitlist(event.data)
        if files:
            file_path = files[0]
            self.load_file(file_path)
    
    def browse_file(self):
        """Open file dialog to select master sheet"""
        file_path = filedialog.askopenfilename(
            title="Select Master Sheet",
            filetypes=[("Excel Files", "*.xlsx;*.xls"), ("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if file_path:
            self.load_file(file_path)
    
    def load_file(self, file_path):
        """Load and validate the selected file"""
        try:
            self.file_path = file_path
            file_name = os.path.basename(file_path)
            self.file_label.config(text=f"✓ Selected: {file_name}", fg="#4caf50")
            self.log_status(f"Loading file: {file_name}", "info")
            
            # Read the file
            if file_path.endswith(".xlsx") or file_path.endswith(".xls"):
                self.df = pd.read_excel(file_path, header=1)
            elif file_path.endswith(".csv"):
                self.df = pd.read_csv(file_path, header=1)
            else:
                messagebox.showerror("Error", "Unsupported file format. Please select an Excel or CSV file.")
                self.log_status("Unsupported file format.", "error")
                return
            
            # Initialize dataframe
            self.df = self.initialize_df(self.df)
            self.log_status(f"Master sheet loaded successfully! ({len(self.df)} records found)", "success")
            self.generate_btn.config(state=tk.NORMAL)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error loading file: {str(e)}")
            self.log_status(f"Error loading file: {str(e)}", "error")
            self.file_label.config(text="Error loading file", fg="#e74c3c")
    
    def start_generation(self):
        """Start the document generation process"""
        if self.df is None:
            messagebox.showwarning("Warning", "Please select a master sheet first.")
            return
        
        # Get selected date from calendar
        if CALENDAR_AVAILABLE and self.calendar_widget:
            try:
                selected_date_str = self.calendar_widget.get_date()
                self.selected_date = datetime.strptime(selected_date_str, "%Y-%m-%d")
            except Exception as e:
                messagebox.showerror("Error", f"Invalid date selected: {str(e)}")
                return
        else:
            # Fallback to entry field
            try:
                date_str = self.date_entry.get()
                self.selected_date = datetime.strptime(date_str, "%Y-%m-%d")
            except Exception as e:
                messagebox.showerror("Error", f"Invalid date format. Please use YYYY-MM-DD: {str(e)}")
                return
        
        day = self.selected_date.day
        month = self.selected_date.month
        year = self.selected_date.year
        
        # Disable generate button
        self.generate_btn.config(state=tk.DISABLED)
        
        # Update dataframe with selected date
        self.df['Day'] = day
        self.df['Month'] = month
        self.df['Year'] = year
        
        month_map = {
            1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
            7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
        }
        self.df['Month'] = self.df['Month'].map(month_map)
        
        self.log_status(f"Starting document generation for {year}-{month:02d}-{day:02d}...", "info")
        self.log_status(f"Processing {len(self.df)} documents...", "info")
        self.update_progress(0, "Initializing...")
        
        # Run generation in a separate thread-like manner
        self.root.after(100, lambda: self.generate_document(self.df))
    
    def generate_document(self, df):
        """Generate documents from the dataframe"""
        try:
            # Select output directory
            self.log_status("Please select the output directory...", "info")
            output_file = filedialog.askdirectory(title="Select Output Directory")
            
            if not output_file:
                self.log_status("Generation cancelled: No output directory selected.", "warning")
                self.generate_btn.config(state=tk.NORMAL)
                return
            
            self.output_directory = output_file
            self.log_status(f"Output directory: {output_file}", "success")
            
            # Open Microsoft Word application
            self.log_status("Opening Microsoft Word...", "info")
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            # Template file path
            if getattr(sys, 'frozen', False):
                template_file = os.path.join(sys._MEIPASS, "template.docx")
            else:
                template_file = r"C:\Users\agraw\OneDrive\Desktop\CODE\Work\template.docx"
            
            if not os.path.exists(template_file):
                messagebox.showerror("Error", f"Template file not found at {template_file}")
                self.log_status(f"Template file not found at {template_file}", "error")
                word.Quit()
                self.generate_btn.config(state=tk.NORMAL)
                return
            
            output_file_copy = output_file
            progress_counter = 0
            total_rows = len(df)
            
            self.log_status("Generating documents...", "info")
            self.update_progress(0, "Generating...")

            for index, row in df.iterrows():
                # Copy template
                temp_template_path = os.path.join(output_file, "temp_template.docx")
                temp_template_path = os.path.abspath(os.path.normpath(temp_template_path))
                shutil.copy(template_file, temp_template_path)
                
                if not os.path.exists(temp_template_path):
                    self.log_status(f"Temporary template file not found at {temp_template_path}", "error")
                    continue
                
                # Open document
                output_file = output_file_copy
                try:
                    doc = word.Documents.Open(temp_template_path)
                except Exception as e:
                    self.log_status(f"Error opening file: {e}", "error")
                    continue
                
                # Prepare replacements
                name = f"{row['Student']}"
                id = str(row['ID'])
                date = f"{row['Month']} {row['Day']:02d}, {row['Year']}"
                course = f"{row['Course_Name']} {row['Course_Code']} {row['Course_Section']}"
                
                if not name or not id or not date:
                    self.log_status(f"Skipping row {index + 1}: Missing required fields", "warning")
                    doc.Close()
                    continue
                
                output_file = self.output_file_generator(row["Month"], row["Year"], row["Centre"], row["Room"], output_file, row["Day"])
                
                # Replace content
                replacements = {"{{Name}}": name, "{{ID}}": id, "{{Date}}": date, "{{Course}}": course}
                self.replace_table_cell_content_in_header(doc, replacements)
                
                # Build filename
                first_name = row['First_Name']
                last_name_initial = row['Last_Name'][0] if isinstance(row['Last_Name'], str) and len(row['Last_Name']) > 0 else ''
                file_name = f"{first_name}.{last_name_initial}.{row['Month']}.{row['Day']}.{row['Year']}.{row['Course_Name']}.{row['Course_Code']}.{row['Course_Section']}.docx"
                output_file_with_timestamp = os.path.join(output_file, file_name)
                
                # Handle duplicates
                counter = 1
                original_output_file = output_file_with_timestamp
                while os.path.exists(output_file_with_timestamp):
                    output_file_with_timestamp = f"{original_output_file[:-5]}_{counter}.docx"
                    counter += 1
                
                # Update progress
                progress_counter += 1
                progress = (progress_counter / total_rows) * 100
                self.update_progress(progress, f"Processing {progress_counter}/{total_rows}...")
                
                # Save document
                try:
                    doc.SaveAs(output_file_with_timestamp, FileFormat=16)
                    self.log_status(f"Saved: {file_name} ({progress_counter}/{total_rows})", "success")
                except Exception as e:
                    self.log_status(f"Error saving file {file_name}: {e}", "error")
                finally:
                    doc.Close()
                
                # Delete temporary template
                if os.path.exists(temp_template_path):
                    os.remove(temp_template_path)

            # Quit Word
            word.Quit()
            
            self.log_status(f"Successfully generated {progress_counter} documents!", "success")
            self.update_progress(100, "Complete!")
            
            # Ask if user wants to process another sheet
            self.ask_another_sheet()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error during generation: {str(e)}")
            self.log_status(f"Error: {str(e)}", "error")
            self.update_progress(0, "Error occurred")
        finally:
            self.generate_btn.config(state=tk.NORMAL)
    
    def ask_another_sheet(self):
        """Ask user if they want to process another sheet"""
        result = messagebox.askyesno(
            "Generation Complete",
            "Documents generated successfully!\n\nWould you like to process another sheet?",
            icon="question"
        )
        
        if result:
            # Reset for another sheet
            self.df = None
            self.file_path = None
            self.file_label.config(text="No file selected", fg="#a0a0a0")
            self.progress_var.set(0)
            self.status_text.config(state=tk.NORMAL)
            self.status_text.delete(1.0, tk.END)
            self.status_text.config(state=tk.DISABLED)
            self.generate_btn.config(state=tk.DISABLED)
            self.update_progress(0, "Ready")
            self.log_status("Ready for a new master sheet. Please select a file to begin.", "info")
        else:
            self.root.quit()
    
    def replace_table_cell_content_in_header(self, doc, replacements):
        """Replace table content in the header of the Word document"""
        section = doc.Sections(1)
        primary_header = section.Headers(1)
        first_page_header = section.Headers(2)
        even_page_header = section.Headers(3)
        
        for header in [primary_header, first_page_header, even_page_header]:
            if header.Range.Tables.Count > 0:
                table = header.Range.Tables(1)
                for row in table.Rows:
                    for cell in row.Cells:
                        text = cell.Range.Text.strip()
                        for placeholder, value in replacements.items():
                            if placeholder in text:
                                cell.Range.Text = text.replace(placeholder, value)
                break
    
    def output_file_generator(self, month, year, centre, room, file, day=None):
        """Generate output file path"""
        month_full_map = {
            "Jan": "January", "Feb": "February", "Mar": "March", "Apr": "April",
            "May": "May", "Jun": "June", "Jul": "July", "Aug": "August",
            "Sep": "September", "Oct": "October", "Nov": "November", "Dec": "December"
        }
        month_full = month_full_map.get(month, month)
        
        if day is None:
            day = 1
        
        path = os.path.join(file, f"{month_full} {day}", "Word Documents Completed")
        os.makedirs(path, exist_ok=True)
        return path
    
    def initialize_df(self, df):
        """Initialize dataframe by cleaning and extracting necessary columns"""
        df.drop(df.columns[[0, 1, 3, 6, 7, 5, 8, 10, 11]], axis=1, inplace=True)
        df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
        df.dropna(inplace=True)
        df["Word"] = df.iloc[:, -2]
        df["ID"] = df.iloc[:, -2]
        df.drop(df.columns[3], axis=1, inplace=True)
        df.drop(df.columns[3], axis=1, inplace=True)
        df["Course_Name"] = df["Course"].str.split(" ", expand=True)[0]
        df["Course_Code"] = df["Course"].str.split(" ", expand=True)[1]
        df["Course_Section"] = df["Course"].str.split(" ", expand=True)[3]
        
        first_names = []
        last_names = []
        df["Room Booking"] = df["Room Booking"].str.replace(r'\s+', ' ', regex=True)
        
        for name in df['Student']:
            name_parts = name.split()
            first_name = name_parts[0]
            last_name = name_parts[-1]
            first_names.append(first_name)
            last_names.append(last_name)
        
        df['First_Name'] = first_names
        df['Last_Name'] = last_names
        df.drop(columns=["Word"], inplace=True)
        df['ID'] = df['ID'].astype(int)
        
        buildings = []
        rooms = []
        
        for booking in df["Room Booking"]:
            parts = booking.split()
            building = ''.join(filter(str.isdigit, parts[-2])) if len(parts) > 1 else ""
            room = parts[-1] if len(parts) > 0 else ""
            buildings.append(building)
            rooms.append(room)
        
        df['Centre'] = buildings
        df['Room'] = rooms
        df.drop(df.columns[2], axis=1, inplace=True)
        df.drop(df.columns[0], axis=1, inplace=True)
        
        return df


def main():
    # Use TkinterDnD if available, otherwise regular Tk
    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    
    app = DocumentGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
