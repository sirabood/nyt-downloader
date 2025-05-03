import os
import threading
import tkinter as tk, tkinter.ttk as ttk
from tkinter import filedialog, messagebox, scrolledtext
import yt_dlp as youtube_dl
import queue
import shutil
import sys
import json
from tktooltip import ToolTip
import pandas as pd
from pystray import Icon as TrayIcon, MenuItem as item, Menu
from PIL import Image
import settingsWindow

def check_and_install_ffmpeg():
    """Check if FFmpeg is installed; if not, provide terminal commands for installation."""
    if shutil.which("ffmpeg") is None:  # Check if FFmpeg exists in system path
        messagebox.showinfo(
            "FFmpeg Not Found",
            "FFmpeg is not installed.\n"
            "It is required for this app to work properly.\n\n"
            "To install FFmpeg, run the following commands in your terminal or command prompt:\n\n"
            "For Windows:\n"
            "winget install ffmpeg\n\n"
            "For macOS:\n"
            "brew install ffmpeg\n\n"
            "For Linux (Debian/Ubuntu):\n"
            "sudo apt update && sudo apt install ffmpeg\n\n"
            "Restart the app after installation."
        )
        sys.exit(0)  # Exit the script to allow user to install FFmpeg
    else:
        print("✅ FFmpeg is already installed.")

# Run the check at script start
check_and_install_ffmpeg()

class DownloaderApp:
    def __init__(self, master):
        self.master = master
        master.title("NYT Downloader")
        try:
            master.iconbitmap("Assets/logo.ico")
        except:
            pass
        master.state("zoomed")  # Fullscreen
        master.protocol("WM_DELETE_WINDOW", self.close_application)
        self.download_directory = ""
        self.queue = []  # Store download items
        self.handy_widgets = {}
        self.is_loading = False  # Flag to prevent recursive loading
        self.preset_dir = "settings_presets"
        self.download_queue = queue.Queue()  # Queue for managing tasks
        self.max_threads = 4  # Limit concurrent downloads
        self.preset_dir = "settings_presets"
        try:
            with open(f"{self.preset_dir}/default.json", "r") as f:
                self.global_ydl_opts = json.load(f)  # Load global options from JSON file
        except FileNotFoundError:
            self.global_ydl_opts = {}  # Initialize with empty options if file not found
        self.export_image = tk.PhotoImage(file="Assets/export.png")
        self.clear_image = tk.PhotoImage(file="Assets/clear.png")
        self.search_image = tk.PhotoImage(file="Assets/search.png")
        self.refresh_image = tk.PhotoImage(file="Assets/refresh.png")
        self.settings_image = tk.PhotoImage(file="Assets/settings.png")
        self.import_image = tk.PhotoImage(file="Assets/import.png")
        self.format_options = {
            "All Available Formats": "all",
            "Best (Video+Audio)": "best",
            "Best Video+Audio (Muxed)": "bv*+ba/best",
            "Best ≤720p": "bestvideo[height<=720]+bestaudio/best[height<=720]",
            "Best ≤480p": "bestvideo[height<=480]+bestaudio/best[height<=480]",
            "Worst Video+Audio (Muxed)": "worst",
            "Best Video Only": "bestvideo",
            "Worst Video Only": "worstvideo",
            "Best Video (any stream)": "bestvideo*",
            "Worst Video (any stream)": "worstvideo*",
            "Best Audio Only": "bestaudio",
            "Worst Audio Only": "worstaudio",
            "Best Audio (any stream)": "bestaudio*",
            "Worst Audio (any stream)": "worstaudio*",
            "Smallest File": "best -S +size,+br,+res,+fps",
        }

        self.video_exts = ["mp4", "webm", "mkv", "flv", "avi"]
        self.audio_exts = ["mp3", "m4a", "aac", "wav", "ogg", "opus"]
        self.categories = [
            "sponsor", "intro", "outro", "selfpromo", "interaction",
            "music_offtopic", "preview", "filler", "exclusive_access",
            "poi_highlight", "poi_nonhighlight"
        ]
        self.create_widgets()
        self.master.bind("<F2>", self.open_settings_window_for_selected)  # F2 Key for settings

    def create_widgets(self):
        # Main layout: Three panels
        self.panedwindow = tk.PanedWindow(self.master, orient=tk.HORIZONTAL, sashwidth=5)
        self.panedwindow.pack(fill=tk.BOTH, expand=1)

        # Left panel: Input & Options
        self.left_main_frame = tk.Frame(self.panedwindow, width=300, bg="lightgray")
        self.panedwindow.add(self.left_main_frame)

        # Add a vertical scrollbar to the left panel
        self.left_scrollbar = tk.Scrollbar(self.left_main_frame, orient=tk.VERTICAL, bg="#169976")
        self.left_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Create a canvas to enable scrolling
        self.left_canvas = tk.Canvas(self.left_main_frame, bg="lightgray", yscrollcommand=self.left_scrollbar.set)
        self.left_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Configure the scrollbar to scroll the canvas
        self.left_scrollbar.config(command=self.left_canvas.yview,)

        # Create a frame inside the canvas to hold the widgets
        self.left_frame = tk.Frame(self.left_canvas, bg="lightgray")
        self.left_canvas.create_window((0, 0), window=self.left_frame, anchor="nw")  # Anchor to the top-left corner

        # Bind the canvas to resize and scroll properly
        def update_scroll_region(event=None):
            self.left_canvas.configure(scrollregion=self.left_canvas.bbox("all"))
            if self.left_frame.winfo_height() <= self.left_canvas.winfo_height():
                self.left_canvas.unbind_all("<MouseWheel>")
                self.left_scrollbar.pack_forget()
            else:
                #re-enable Scrollbar and bind mouse wheel event
                self.left_canvas.bind_all("<MouseWheel>", lambda e: self.left_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))
                self.left_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.master.bind("<Configure>", update_scroll_region)
        self.left_canvas.bind("<Configure>", lambda e: self.left_canvas.itemconfig(self.left_canvas.find_withtag("all")[0], width=e.width))

        self.main_right_frame = tk.PanedWindow(self.panedwindow, width=300, orient=tk.VERTICAL, bg="lightgray")
        self.main_right_frame.pack(fill=tk.BOTH, expand=1)
        self.panedwindow.add(self.main_right_frame)
        # Right panel: Download Queue
        self.right_frame = tk.Frame(self.main_right_frame, width=300, bg="lightgray")
        self.main_right_frame.add(self.right_frame)
        # Bottom panel: Logs
        self.bottom_frame = tk.Frame(self.main_right_frame, bg="lightgray")
        self.main_right_frame.add(self.bottom_frame)

        """Improved Bottom Panel with Colored Logs and Enhanced UI"""

        log_frame = tk.LabelFrame(self.bottom_frame, text="📜 Download Logs", font=("Arial", 10, "bold"))
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Top bar for log controls
        log_controls_frame = tk.Frame(log_frame, bg="lightgray")
        log_controls_frame.pack(fill=tk.X, padx=5, pady=5)

        # Export button to save logs to a file
        export_button = tk.Button(log_controls_frame, command=self.export_logs,image=self.export_image,relief="flat",bg="lightgray")
        export_button.pack(side=tk.LEFT, padx=5, pady=5)
        ToolTip(export_button, msg="Export logs to a file can be useful for sharing problems.\n Shift to verbose output in Full settings for detailed logs.", delay=0.5)

        # Clear logs button
        clear_logs_button = tk.Button(log_controls_frame, command=lambda: self.log_text.delete("1.0", tk.END),image=self.clear_image,relief="flat",bg="lightgray")
        clear_logs_button.pack(side=tk.LEFT, padx=5, pady=5)
        ToolTip(clear_logs_button, msg="Clear all logs from the log window.\nThis action cannot be undone, so ensure you have exported the logs if needed.", delay=0.5)

        # Search logs entry
        search_entry = tk.Entry(log_controls_frame, font=("Arial", 8))
        search_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        search_entry.bind("<Return>", lambda e: search_logs() if self.master.focus_get() == search_entry else None)  # Bind Enter key to search function
        ToolTip(search_entry, msg="Press Enter to search logs after typing.It highlights Text into yellow.", delay=0.5)
        def search_logs():
            """Highlight search results in the log text."""
            self.log_text.tag_remove("highlight", "1.0", tk.END)  # Remove previous highlights
            query = search_entry.get().strip()
            if query:
                start_pos = "1.0"
            while True:
                start_pos = self.log_text.search(query, start_pos, stopindex=tk.END, nocase=True)
                if not start_pos:
                    break
                end_pos = f"{start_pos}+{len(query)}c"
                self.log_text.tag_add("highlight", start_pos, end_pos)
                start_pos = end_pos
            self.log_text.tag_config("highlight", background="yellow", foreground="black")

        search_button = tk.Button(log_controls_frame, command=search_logs,image=self.search_image,borderwidth=0,relief="flat",bg="lightgray")
        search_button.pack(side=tk.LEFT, padx=5, pady=5)

        # Log text area with scroll
        self.log_text = scrolledtext.ScrolledText(log_frame, bg="#222", fg="white", font=("Consolas", 10), wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        # Enable Ctrl+A to select all text in the log area
        self.log_text.bind("<Control-a>", lambda e: self.log_text.tag_add("sel", "1.0", tk.END) or "break")        
        # Block other editing but allow specific key combinations
        self.log_text.bind("<Key>", lambda e: None if e.state & 4 and e.keysym in ('a', 'A', 'c', 'C') else "break")
    
        #           ------------ left panel widgets --------------

        # Section: YouTube Input
        input_frame = tk.LabelFrame(self.left_frame, text="🎥 Input", font=("Arial", 10, "bold"))
        input_frame.pack(fill=tk.X, padx=10, pady=5,expand=True)
        
        # Search entry field
        self.query_entry = tk.Entry(input_frame, width=45, font=("Arial", 10))
        self.query_entry.pack(pady=5, padx=10, fill=tk.X, expand=True)
        ToolTip(self.query_entry, msg="Enter URL of any supported websites(check yt_dlp docs) or Search Query in YouTube.\nPress Enter to add to queue.", delay=0.5)
        # ✅ Bind Enter Key to "Add to Queue" function
        self.query_entry.bind("<Return>", self.add_to_queue)
        #A note label
        note_label = tk.Label(input_frame, text="Please use (Any Stream) format for other than ytsites.", font=("Arial", 8), fg="gray").pack(pady=5, padx=10, fill=tk.X)
        add_button = tk.Button(input_frame, text="➕ Add to Queue", command=self.add_to_queue, bg="#28a745", fg="white", font=("Arial", 10, "bold"))
        add_button.pack(pady=5, padx=10, fill=tk.X)
        ToolTip(add_button, msg="Click to add the URL or Search Query to the download queue.\n Note: It just adds to queue and not start downloads.\n Press Strat Download Manually.", delay=0.5)

        # Section: Load Spreadsheet/CSV
        csv_frame = tk.LabelFrame(self.left_frame, text="📂 Load Spreadsheet", font=("Arial", 10, "bold"))
        csv_frame.pack(fill=tk.X, padx=10, pady=5)
        select_file_button = tk.Button(csv_frame, text="📄 Select Spreadsheet File", command=self.load_spreadsheet_threaded, bg="#007bff", fg="white", font=("Arial", 10))
        select_file_button.pack(pady=5, padx=10, fill=tk.X)
        ToolTip(select_file_button, msg="Load CSV (Comma-separated Values) or Spreadsheet file containing URLs or Search Queries.\nIt will add all entries to the download queue.", delay=0.5)
        # Label for file format instructions
        format_label = tk.Label(csv_frame, text="Supported Formats:\n- CSV: Comma-separated URLs or Queries\n- Excel: Columns 'Query' or 'URL' (Optional: 'Preset')", font=("Arial", 8), fg="gray", justify="left")
        format_label.pack(pady=5, padx=10, fill=tk.X)
        # Section: Download Location
        dir_frame = tk.LabelFrame(self.left_frame, text="📁 Download Directory", font=("Arial", 10, "bold"))
        dir_frame.pack(fill=tk.X, padx=10, pady=5)
        dir_button = tk.Button(dir_frame, text="📂 Select Folder", command=self.select_directory, bg="#ff9800", fg="white", font=("Arial", 10))
        dir_button.pack(pady=5, padx=10, fill=tk.X)
        self.dir_label = tk.Label(dir_frame, text="No directory selected", font=("Arial", 9), fg="gray")
        self.dir_label.pack(pady=5)
        ToolTip(dir_button, msg="Click to select the folder where you want to save the downloaded files.\nIt will be used as default download location.\n You can edit download location for single file even after queueing in File and Naming option Tab", delay=0.5)

        #Handy Settings Option
        handy_frame = tk.LabelFrame(self.left_frame, text="🛠️ Handy Settings", font=("Arial", 10, "bold"))
        handy_frame.pack(fill=tk.X, padx=10, pady=5)

        # Preset Dropdown
        tk.Label(handy_frame, text="Select Preset:", font=("Arial", 10)).grid(row=0, column=0, padx=5, sticky="w")
        self.preset_var = tk.StringVar()
        self.preset_dropdown = ttk.Combobox(handy_frame, textvariable=self.preset_var, state="readonly")
        self.preset_dropdown.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        self.preset_dropdown.bind("<<ComboboxSelected>>", self.on_preset_selected)
        files = [f[:-5] for f in os.listdir(self.preset_dir) if f.endswith(".json")]
        self.preset_dropdown["values"] = files + ["New/Unsaved"]
        if "default" in files:
            self.preset_dropdown.set("default")
        else:
            self.preset_dropdown.set("New/Unsaved")
        ToolTip(self.preset_dropdown, msg="Select a preset to load its settings.\nYou can also create new presets from the settings window.", delay=0.5)

        #Format dropdown
        tk.Label(handy_frame, text="Select Format:", font=("Arial", 10)).grid(row=1, column=0, padx=5, sticky="w")
         # Sync dropdown and entry
        def update_format_entry(event=None):
            selected_label = self.format_var.get()
            # Adapt final extension options
            if "Audio" in selected_label and not "Video" in selected_label:
                self.handy_widgets['final_ext']['values'] = ["original"] + self.audio_exts
                self.handy_widgets['final_ext'].set("original")  # Set default to "original"
            elif "Video" in selected_label or "Quality" in selected_label or "MP4" in selected_label:
                self.handy_widgets['final_ext']['values'] = ["original"] + self.video_exts
                self.handy_widgets['final_ext'].set("original")  # Set default to "original"
            else:
                self.handy_widgets['final_ext']['values'] = ["original"] + self.video_exts + self.audio_exts  # fallback
                self.handy_widgets['final_ext'].set("original")  # Set default to "original"

        # Format Dropdown
        self.format_var = tk.StringVar(value="Best Video+Audio (Muxed)")
        self.handy_widgets["format_dropdown"] = ttk.Combobox(master=handy_frame, textvariable=self.format_var, values=list(self.format_options.keys()),state="readonly")
        self.handy_widgets["format_dropdown"].grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        self.handy_widgets['format_dropdown'].bind("<<ComboboxSelected>>", update_format_entry)
        ToolTip(self.handy_widgets["format_dropdown"], msg="Select a format for the download.\n Please select Any stream format if face format error.\n Also you can manually type any legal format in settings Window.\nIt will automatically adapt the final extension options.", delay=0.5)

        #final Extension Dropdown
        tk.Label(handy_frame, text="Select Extension:", font=("Arial", 10)).grid(row=2, column=0, padx=5, sticky="w")
        final_ext_options = ["original"] + self.video_exts
        self.final_ext_var = tk.StringVar(value="original")
        self.handy_widgets["final_ext"] = ttk.Combobox(master=handy_frame, textvariable=self.final_ext_var, values=final_ext_options, state="readonly")
        self.handy_widgets["final_ext"].grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        ToolTip(self.handy_widgets["final_ext"], msg="Select the final file extension for the download.\n It will be used as default extension for all downloads.\n Custom formats will be converted on your Machine and that is Resource taking process.", delay=0.5)

        #Embeded Subtitles
        self.embed_subtitles_var = tk.BooleanVar(value=False)
        self.handy_widgets["embedsubtitles"] = tk.Checkbutton(handy_frame, text="Embed Subtitles", variable=self.embed_subtitles_var, font=("Arial", 10))
        self.handy_widgets["embedsubtitles"].grid(row=3, column=0, padx=5, sticky="w")
        ToolTip(self.handy_widgets["embedsubtitles"], msg="Embed subtitles into the video file.\n Full Control over it in Settings Window.", delay=0.5)

        #Embeded Thumbnail
        self.embed_thumbnail_var = tk.BooleanVar(value=False)
        self.handy_widgets["embedthumbnail"] = tk.Checkbutton(handy_frame, text="Embed Thumbnail", variable=self.embed_thumbnail_var, font=("Arial", 10))
        self.handy_widgets["embedthumbnail"].grid(row=3, column=1, padx=5, sticky="w")
        ToolTip(self.handy_widgets["embedthumbnail"], msg="Embed thumbnail into the video file.", delay=0.5)

        #writeautomaticsub checkbox
        self.writeautomaticsub_var = tk.BooleanVar(value=False)
        self.handy_widgets["writeautomaticsub"] = tk.Checkbutton(handy_frame, text="Include Auto subs", variable=self.writeautomaticsub_var, font=("Arial", 10))
        self.handy_widgets["writeautomaticsub"].grid(row=4, column=0, padx=5, sticky="w")
        ToolTip(self.handy_widgets["writeautomaticsub"], msg="Include auto-generated subtitles in the download.\n Full Control over it in Settings Window.", delay=0.5)

        #Metadata Checkbox
        self.metadata_var = tk.BooleanVar(value=False)
        self.handy_widgets["addmetadata"] = tk.Checkbutton(handy_frame, text="Add Metadata", variable=self.metadata_var, font=("Arial", 10))
        self.handy_widgets["addmetadata"].grid(row=4, column=1, padx=5, sticky="w")
        ToolTip(self.handy_widgets["addmetadata"], msg="Add metadata to the downloaded file.\n Metadata includes title,author, etc. data. ", delay=0.5)

        #providing two radiobuttons for remove all Sponserblock or mark all sponserblock and one for Other then listing other selected as Label
        tk.Label(handy_frame, text="Sponserblock:", font=("Arial", 10)).grid(row=5, column=0, padx=5, sticky="w")
        self.sponserblock_var = tk.StringVar(value="mark")
        self.mark_all_radio = tk.Radiobutton(handy_frame, text="Mark All", variable=self.sponserblock_var, value="mark", font=("Arial", 10))
        self.mark_all_radio.grid(row=6, column=0, padx=5, sticky="w")
        ToolTip(self.mark_all_radio, msg="Mark all sponserblock segments in the video.\n Full Control over it in Settings Window.", delay=0.5)
        self.remove_all_radio = tk.Radiobutton(handy_frame, text="Remove All", variable=self.sponserblock_var, value="remove", font=("Arial", 10))
        self.remove_all_radio.grid(row=6, column=1, padx=5, sticky="w")
        ToolTip(self.remove_all_radio, msg="Remove all sponserblock segments from the video.\n Full Control over it in Settings Window.", delay=0.5)
        self.Other_radio = tk.Radiobutton(handy_frame, text="Custom", variable=self.sponserblock_var, value="Other", font=("Arial", 10))
        self.Other_radio.grid(row=6, column=1, sticky="e")
        ToolTip(self.Other_radio, msg="Custom Sponserblock settings.\n If not any preset selected and not opted anything,\n in Settings window it sets to do nothing.", delay=0.5)

        #binding save_handy_settings_into_global to all handy_settings widgets
        self.format_var.trace_add("write", lambda *args: self.save_handy_settings_into_global())
        self.final_ext_var.trace_add("write", lambda *args: self.save_handy_settings_into_global())
        self.embed_thumbnail_var.trace_add("write", lambda *args: self.save_handy_settings_into_global())
        self.metadata_var.trace_add("write", lambda *args: self.save_handy_settings_into_global())
        self.sponserblock_var.trace_add("write", lambda *args: self.save_handy_settings_into_global())
        self.embed_subtitles_var.trace_add("write", lambda *args: self.save_handy_settings_into_global())
        self.writeautomaticsub_var.trace_add("write", lambda *args: self.save_handy_settings_into_global())

        # Load settings from global_ydl_opts into handy widgets
        if self.global_ydl_opts != {}:
            self.load_handy_settings()
        else:
            self.save_handy_settings_into_global()

        # Save settings to global_ydl_opts button
        tk.Label(handy_frame, text="If not autosaved :", font=("Arial", 10)).grid(row=9, column=0, padx=5, sticky="w")
        tk.Button(handy_frame, text="💾 Save Settings", command=self.save_handy_settings_into_global, bg="#007bff", fg="white", font=("Arial", 10)).grid(row=9, column=1, padx=10, pady=5, sticky="ew")
        # ⚙️ Full Settings Button
        self.settings_button = tk.Button(handy_frame, text="⚙️ Full Settings", command=self.open_settings_window, bg="#007bff", fg="white", font=("Arial", 10))
        ToolTip(self.settings_button, msg="Click to open the full settings window.\n You can customize various options for downloads.\n It will be saved as default template.\n press F2 as shortcut can edit individual task settings via it also.", delay=0.5)
        self.settings_button.grid(row=10, column=0,columnspan=2, padx=5, pady=5,sticky="ew")

        # Section: Actions
        actions_frame = tk.LabelFrame(self.left_frame, text="⚡ Actions", font=("Arial", 10, "bold"))
        actions_frame.pack(fill=tk.X, padx=10, pady=5)

        # Start Download Button
        download_button = tk.Button(actions_frame, text="🚀 Start Downloads", command=self.start_downloads, bg="#dc3545", fg="white", font=("Arial", 12, "bold"))
        download_button.pack(pady=10, padx=10, fill=tk.X)
        ToolTip(download_button, msg="Click to start the downloads in the queue.\n It will start downloading all the items in the queue.\n It starts 4 concurrent downloads for now.", delay=0.5)

        # Minimize to Tray Button
        bg_button = tk.Button(actions_frame, text="🛡️ Run in Background", bg="#444", fg="white", command=self.minimize_to_tray)
        bg_button.pack(pady=10, padx=10, fill=tk.X)
        ToolTip(bg_button, msg="Click to minimize the app to the system tray.\n You can restore it from the tray icon.", delay=0.5)

        """Right Panel with a Styled Table for Queue"""
        queue_frame = tk.LabelFrame(self.right_frame, text="📥 Download Queue", font=("Arial", 10, "bold"))
        queue_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Styled Table (Treeview)
        self.queue_table = ttk.Treeview(queue_frame, columns=("Index","Status", "Query","Progress"), show="headings", height=18,selectmode="extended",)

        # Define column headings
        self.queue_table.heading("Index", text="Sr. No.", anchor="center")
        self.queue_table.heading("Query", text="📄 Title / Query", anchor="w")
        self.queue_table.heading("Status", text="🔄 Status", anchor="center")
        self.queue_table.heading("Progress", text="Progress", anchor="center")

        # Define column widths
        self.queue_table.column("Index", width=25, anchor="center")  # Small column for numbering
        self.queue_table.column("Query", width=250, anchor="w")
        self.queue_table.column("Status", width=100, anchor="center")
        self.queue_table.column("Progress", width=150, anchor="center")  # ✅ Progress column

        # Add Scrollbar
        scrollbar = ttk.Scrollbar(queue_frame, orient=tk.VERTICAL, command=self.queue_table.yview)
        self.queue_table.configure(yscroll=scrollbar.set)
        
        #add horizontal scrollbar
        h_scrollbar = ttk.Scrollbar(queue_frame, orient=tk.HORIZONTAL, command=self.queue_table.xview)
        self.queue_table.configure(xscroll=h_scrollbar.set)

        # Add some button controls for queue
        button_frame = tk.Frame(queue_frame)
        button_frame.grid(row=2, column=0, padx=5, pady=0, sticky="ew")

        self.refresh_button = tk.Button(button_frame, image=self.refresh_image, command=self.update_queue_listbox_threadsafe, borderwidth=0, relief="flat")
        self.refresh_button.pack(side=tk.LEFT, padx=10)
        ToolTip(self.refresh_button, msg="Refresh\nRefresh the queue list although it refreshes automatically. You can press it to manually refresh.\nIt will not remove any item from the queue.", delay=0.5)

        self.export_button = tk.Button(button_frame, image=self.export_image, command=self.export_queue, borderwidth=0, relief="flat")
        self.export_button.pack(side=tk.LEFT, padx=10)
        ToolTip(self.export_button, msg="Export\nExport the pending downloads with all settings saved into JSON format.\nThis can be useful for sharing or using it later via the Import button.", delay=0.5)

        self.import_button = tk.Button(button_frame, image=self.import_image, command=self.import_queue, borderwidth=0, relief="flat")
        self.import_button.pack(side=tk.LEFT, padx=10)
        ToolTip(self.import_button, msg="Import\nImport a JSON file containing download items.\nIt will add all entries to the download queue.", delay=0.5)

        self.clear_button = tk.Button(button_frame, image=self.clear_image, command=self.clear_queue, borderwidth=0, relief="flat")
        self.clear_button.pack(side=tk.LEFT, padx=10)
        ToolTip(self.clear_button, msg="Clear\nClear all items from the download queue if none are selected, or clear only the selected items.\nNote: This action cannot be undone.", delay=0.5)

        self.settings_button = tk.Button(button_frame, image=self.settings_image, command=self.open_settings_window_for_selected, borderwidth=0, relief="flat")
        self.settings_button.pack(side=tk.LEFT, padx=10)
        ToolTip(self.settings_button, msg="⚙️ Settings (F2)\nOpen settings for the selected item(s) in the queue.\nIf more than one item is selected, it changes settings for all in common.\nIf no item is selected, it opens Global settings.\nYou can customize various options for specific downloads.", delay=0.5)
        
        # Grid the Treeview and Scrollbar properly
        self.queue_table.grid(row=0, column=0, sticky="nsew")   # Fill available space
        scrollbar.grid(row=0, column=1, sticky="ns")   # Attach scrollbar to the right
        h_scrollbar.grid(row=1, column=0, sticky="ew")  # Attach horizontal scrollbar to the bottom
        tk.Label(queue_frame, text="🔍 Click to select individual , Shift+Click to select range,Ctrl + click to select multiple randomly, Ctrl+A to select all", font=("Arial", 10), fg="gray",justify='center').grid(row=3, column=0, sticky="nsew", padx=5, pady=0)
        # Make the frame expandable
        queue_frame.columnconfigure(0, weight=1)
        queue_frame.rowconfigure(0, weight=1)

        # ✅ Bind all selection methods
        self.queue_table.bind("<Shift-Button-1>", self.on_shift_click)  # Shift + Click
        self.queue_table.bind("<B1-Motion>", self.on_drag_select)  # Drag Selection
        self.queue_table.bind("<Control-a>", self.select_all)  # Ctrl + A
        self.queue_table.bind("<Delete>", self.del_key)  # Delete Key
        self.queue_table.bind("backspace", self.del_key)  # Backspace Key

    def save_handy_settings_into_global(self):
        """Save settings from the handy widgets into global_ydl_opts."""
        if self.is_loading:
            return
        #get postprocessors from global_ydl_opts
        postprocessors = self.global_ydl_opts.get('postprocessors', [])
        selected_label = self.format_var.get()
        value = self.format_options.get(selected_label, "Custom")
        selected_ext = self.final_ext_var.get()
        embed_thumbnail = self.embed_thumbnail_var.get()
        add_metadata = self.metadata_var.get()
        # Sponserblock
        sponserblock = self.sponserblock_var.get()
        #embeded Subs
        embed_subtitles = self.embed_subtitles_var.get()
        #writeautomaticsub
        writeautomaticsub = self.writeautomaticsub_var.get()

        # Update global options based on selection
        self.global_ydl_opts['format_dropdown'] = selected_label
        self.global_ydl_opts['format'] = value
        if selected_ext == "original":
            self.global_ydl_opts['final_ext'] = None
        else:
            self.global_ydl_opts['final_ext'] = selected_ext     
        self.global_ydl_opts['embedthumbnail'] = embed_thumbnail
        if embed_thumbnail:
            self.global_ydl_opts['writethumbnail'] = True
            self.global_ydl_opts['embedthumbnail'] = True
        else:
            self.global_ydl_opts['writethumbnail'] = False
            self.global_ydl_opts['embedthumbnail'] = False
        self.global_ydl_opts['addmetadata'] = add_metadata
        if embed_subtitles:
            self.global_ydl_opts['writesubtitles'] = True
            self.global_ydl_opts['embedsubtitles'] = True
        else:
            self.global_ydl_opts['writesubtitles'] = False
            self.global_ydl_opts['embedsubtitles'] = False
        if writeautomaticsub:
            self.global_ydl_opts['writeautomaticsub'] = True
            self.global_ydl_opts['embedsubtitles'] = True
        else:
            self.global_ydl_opts['writeautomaticsub'] = False
            self.global_ydl_opts['embedsubtitles'] = False

        # --- 2. Update Postprocessors ---    
        def update_postprocessor(postprocessors, new_pp):
            """ Replace or add a postprocessor dict with the same 'key'. """
            for i, pp in enumerate(postprocessors):
                if pp.get("key") == new_pp.get("key"):
                    postprocessors[i] = new_pp  # Replace
                    return
            postprocessors.append(new_pp)  # Add if not found

        # FFmpegMetadata (Metadata + Chapters)
        if add_metadata:
            meta_pp = {"key": "FFmpegMetadata"}
            meta_pp["add_metadata"] = True
            meta_pp["add_chapters"] = True
            update_postprocessor(postprocessors, meta_pp)
        else:
            for i, pp in enumerate(postprocessors):
                if pp.get("key") == "FFmpegMetadata":
                    postprocessors.pop(i)
                    break
        if selected_ext:
            # FFmpegExtractAudio (Audio Extraction)
            if selected_ext in self.audio_exts:
                # Infer preferredquality from format dropdown
                quality = "5"  # default medium
                if "best" in value:
                    quality = "0"
                elif "smallest" in value:
                    quality = "9"
                update_postprocessor(postprocessors, {
                    "key": "FFmpegExtractAudio",
                    "preferredcodec": selected_ext,
                    "preferredquality": quality
                })
                # Remove FFmpegVideoConvertor if it exists it saves the time of conversion
                for i, pp in enumerate(postprocessors):
                    if pp.get("key") == "FFmpegVideoConvertor":
                        postprocessors.pop(i)
                        break
            # FFmpegVideoConvertor (Video Conversion) in desired format
            elif selected_ext in self.video_exts:
                update_postprocessor(postprocessors, {
                    "key": "FFmpegVideoConvertor",
                    "preferedformat": selected_ext
                })
            else:
                # If the selected extension is not in audio or video, remove any existing FFmpegExtractAudio and FFmpegVideoConvertor postprocessor
                for i, pp in enumerate(postprocessors):
                    if pp.get("key") == "FFmpegExtractAudio":
                        postprocessors.pop(i)
                        break
                    if pp.get("key") == "FFmpegVideoConvertor":
                        postprocessors.pop(i)
                        break
        # EmbedThumbnail in last to avoid conflicts with other postprocessors
        if embed_thumbnail:
            update_postprocessor(postprocessors,{"key": "EmbedThumbnail"})
        else:
            for i, pp in enumerate(postprocessors):
                if pp.get("key") == "EmbedThumbnail":
                    postprocessors.pop(i)
                    break
        # EmbedSubtitles in last to avoid conflicts with other postprocessors
        if embed_subtitles:
            update_postprocessor(postprocessors,{'key': 'FFmpegEmbedSubtitle'})
        else:
            for i, pp in enumerate(postprocessors):
                if pp.get("key") == "FFmpegEmbedSubtitle":
                    postprocessors.pop(i)
                    break

        # Sponsorblock
        if sponserblock == "mark":
            self.global_ydl_opts['sponsorblock_mark'] = self.categories
            self.global_ydl_opts['sponsorblock_remove'] = []
            self.global_ydl_opts['add_chapters'] = True
            update_postprocessor(postprocessors, {
            "key": "SponsorBlock",
            "categories": self.categories,
            "when": 'after_filter',
            })
            update_postprocessor(postprocessors, {
            "key": "ModifyChapters",
            "sponsorblock_chapter_title": self.global_ydl_opts.get('sponsorblock_chapter_title','[SponsorBlock]: %(category_names)l')
            })
            update_postprocessor(postprocessors, {
            "key": "FFmpegMetadata",
            "add_metadata": add_metadata,
            "add_chapters": True,
            })

        elif sponserblock == "remove":
            self.global_ydl_opts['sponsorblock_remove'] = self.categories
            self.global_ydl_opts['add_chapters'] = True
            if self.global_ydl_opts.get('sponsorblock_mark', None):
                self.global_ydl_opts['sponsorblock_mark'] = []

            update_postprocessor(postprocessors, {
            "key": "SponsorBlock",
            "categories": self.categories,
            "when": 'after_filter',
            })
            update_postprocessor(postprocessors, {
            "key": "ModifyChapters",
            'remove_chapters_patterns': [],
            'remove_ranges': [],
            "sponsorblock_chapter_title": '[SponsorBlock]: %(category_names)l',
            'remove_sponsor_segments': self.categories
            })
            update_postprocessor(postprocessors, {
            "key": "FFmpegMetadata",
            "add_metadata": add_metadata,
            "add_chapters": True,
            })
        elif sponserblock == "Other":
            if self.global_ydl_opts.get('sponsorblock_mark', None) == self.categories:
                self.global_ydl_opts['sponsorblock_mark'] = []
                # remove also from postprocessors.
                for i, pp in enumerate(postprocessors):
                    if pp.get("key") == "SponsorBlock":
                        postprocessors.pop(i)
                        break
                    if pp.get("key") == "ModifyChapters":
                            if "sponsorblock_chapter_title" in pp:
                                pp.pop("sponsorblock_chapter_title")
                                break
            if self.global_ydl_opts.get('sponsorblock_remove', None) == self.categories:
                self.global_ydl_opts['sponsorblock_remove'] = []
                # remove also from postprocessors.
                for i, pp in enumerate(postprocessors):
                    if pp.get("key") == "SponsorBlock":
                        postprocessors.pop(i)
                        break
                    if pp.get("key") == "ModifyChapters":
                        if "remove_ranges" in pp:
                            pp.pop("remove_ranges")

                        if "remove_chapters_patterns" in pp:
                            pp.pop("remove_chapters_patterns")

                        if "sponsorblock_chapter_title" in pp:
                                pp.pop("sponsorblock_chapter_title")

                        if "remove_sponsor_segments" in pp:
                           pp.pop("remove_sponsor_segments")
                           break

        # Attach postprocessors if any exist
        if postprocessors:
            self.global_ydl_opts['postprocessors'] = postprocessors
        self.mark_as_unsaved_if_modified()  # Mark as unsaved if modified
    # A function to laad widgets value from global_ydl_opts to handy widgets
    def load_handy_settings(self):
        """Load settings from global_ydl_opts into handy widgets."""
        if self.is_loading:
            return
        try:
            self.is_loading = True
            self.format_var.set(self.global_ydl_opts.get('format_dropdown', "Best (Video+Audio)"))
            self.final_ext_var.set(self.global_ydl_opts.get('final_ext', "original"))
            self.embed_thumbnail_var.set(self.global_ydl_opts.get('embedthumbnail', False))
            self.metadata_var.set(self.global_ydl_opts.get('addmetadata', False))
            self.embed_subtitles_var.set(self.global_ydl_opts.get('embedsubtitles', False))
            self.writeautomaticsub_var.set(self.global_ydl_opts.get('writeautomaticsub', False))
            
            # Sponserblock
            sb_mark = self.global_ydl_opts.get('sponsorblock_mark', None)
            sb_remove = self.global_ydl_opts.get('sponsorblock_remove', None)
            if sb_mark == self.categories:
                self.sponserblock_var.set("mark")
            elif sb_remove == self.categories:
                self.sponserblock_var.set("remove")
            else:
                self.sponserblock_var.set("Other")
                #add what is in other to the label of handy settings
                if not sb_mark:
                    sb_mark = "None"
                if not sb_remove:
                    sb_remove = "None"

        except Exception as e:
            self.log(f"Error loading settings: {e}", level="error")
        finally:
            self.is_loading = False 
    def mark_as_unsaved_if_modified(self):
        current_name = self.preset_var.get()
        if current_name in ["New/Unsaved", ""]:
            return  # Already unsaved
        try:
            with open(os.path.join(self.preset_dir, f"{current_name}.json")) as f:
                saved_opts = json.load(f)
            if saved_opts != self.global_ydl_opts:
                self.preset_var.set("New/Unsaved")
        except:
            self.preset_var.set("New/Unsaved")

    def on_preset_selected(self, event=None):
        name = self.preset_var.get()
        if name == "New/Unsaved":
            return
        path = os.path.join(self.preset_dir, f"{name}.json")
        try:
            with open(path, "r") as f:
                self.is_loading = True
                self.ydl_opts = json.load(f)
                self.global_ydl_opts = self.ydl_opts.copy()  # Update global options
                self.load_handy_settings()
            self.log(f"✅ Loaded preset: {name}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load preset: {e}")
        finally:
            self.is_loading = False

    def open_settings_window(self,initial_opts=None,global_change=True):
        if initial_opts is None:
            initial_opts = self.global_ydl_opts
        settings = settingsWindow.show_settings_window(self.master, initial_opts=initial_opts, global_change=global_change)
        if settings != initial_opts and global_change:
            self.global_ydl_opts = settings
            self.load_handy_settings()
            self.log("🔧 Settings Updated and saved to default template.")
        elif global_change and settings == initial_opts:
            self.log("⚠️ No settings were changed.")
        else:
            return settings  # Return settings for local use
    #function to open settings window for all selected items in the queue
    def open_settings_window_for_selected(self,event=None):
        selected_items = self.queue_table.selection()
        if selected_items:
            if len(selected_items) > 1:
                new_ydl_opts = self.open_settings_window(initial_opts=self.global_ydl_opts, global_change=False)  # Open settings window
                for item in selected_items:
                    index = int(self.queue_table.item(item, "values")[0]) - 1
                    task = self.queue[index]
                    if task["status"] == "Downloading":
                        self.log(f"⚠️ Cannot modify settings for downloading item: {task['query']}", level="warning")
                        continue
                    elif task["status"] == "Completed" or task["status"] == "Failed" or task["status"] == "Queued":
                        try:
                            task["ydl_opts"] = new_ydl_opts  # Update task-specific options
                            self.log(f"🔧 Settings Updated for {task['query']}"
                                     , level="success")
                        except Exception as e:
                            self.log(f"⚠️ Error updating settings: {e}", level="error")
                    else:
                        self.log(f"⚠️ Cannot modify settings for item: {task['query']}", level="warning")
            if len(selected_items) == 1:
                item = selected_items[0]
                index = int(self.queue_table.item(item, "values")[0]) - 1
                task = self.queue[index]
                if task["status"] == "Downloading":
                    self.log(f"⚠️ Cannot modify settings for downloading item: {task['query']}", level="warning")
                    return
                if task["status"] == "Queued" or task["status"] == "Completed" or task["status"] == "Failed":
                    new_ydl_opts = self.open_settings_window(initial_opts=task['ydl_opts'], global_change=False)
                    if new_ydl_opts != task['ydl_opts']:
                        try:
                            task["ydl_opts"] = new_ydl_opts  # Update task-specific options
                            self.log(f"🔧 Settings Updated for {task['query']}", level="success")
                        except Exception as e:
                            self.log(f"⚠️ Error updating settings: {e}", level="error")
                    elif new_ydl_opts == task['ydl_opts']:
                        self.log("⚠️ No settings were changed.")
        else:
            new_ydl_opts=self.open_settings_window(initial_opts=self.global_ydl_opts,global_change=True)  # Open settings window

    def on_shift_click(self, event):
        """Allows Shift + Click to select a range of items in Treeview."""
        selected = self.queue_table.selection()  # Get selected items
        last_clicked = self.queue_table.identify_row(event.y)  # Identify clicked row

        if not selected:
            return  # No selection to extend from

        if last_clicked:
            # ✅ Store the last clicked index as the selection anchor
            self.last_selected_index = self.queue_table.index(last_clicked)

            first_index = self.queue_table.index(selected[0])  # First selected item
            last_index = self.queue_table.index(last_clicked)  # Last clicked item

            # ✅ Select all items in range
            for i in range(min(first_index, last_index), max(first_index, last_index) + 1):
                self.queue_table.selection_add(self.queue_table.get_children()[i])  # ✅ Add to selection

    def on_drag_select(self, event):
        """Allows Drag Selection in Treeview using Mouse Motion."""
        row_id = self.queue_table.identify_row(event.y)
        if row_id:
            self.queue_table.selection_add(row_id)  # ✅ Add item under cursor to selection
    def select_all(self, event=None):
        """Selects all items in Treeview only if it's focused."""
        if self.queue_table == self.master.focus_get():  # ✅ Check if Treeview is focused
            self.queue_table.selection_set(self.queue_table.get_children())  # ✅ Select all
            return "break"  # ✅ Prevents default text selection behavior

    def del_key(self, event):
        """Delete key to remove selected items from the queue."""
        if self.master.focus_get() == self.queue_table:
            self.clear_queue()
    def add_to_queue(self,event=None):
        """Adds a video or full playlist to the queue (fetch videos first)."""
        query = self.query_entry.get().strip()
        if not query and self.master.focus_get() == self.query_entry:
            messagebox.showwarning("Warning", "Please enter a YouTube link or search query.")
            return
        else:
            pass
    
        # Check if the link is a playlist
        if "playlist?" in query or "list=" in query:
            def fetch_playlist():
                self.log(f"Fetching videos from playlist: {query}")

                ydl_opts = {
                    'quiet': True,
                    'extract_flat': True,  # ✅ Fetch only metadata (don't download)
                    'force_generic_extractor': True,
                }

                try:
                    with youtube_dl.YoutubeDL(ydl_opts) as ydl:
                        playlist_info = ydl.extract_info(query, download=False)

                    # Check if we got a valid playlist response
                    if 'entries' in playlist_info:
                        for video in playlist_info['entries']:
                            if 'url' in video and 'title' in video:
                                task = {
                                    "query": video['url'],
                                    "status": "Queued",
                                    "ydl_opts": self.global_ydl_opts.copy(),  # Copy global options
                                }
                                self.queue.append(task)
                                self.log(f"Added video to queue: {video['title']}")
                        self.log(f"Added {len(playlist_info['entries'])} videos to the queue.")
                    else:
                        self.log("⚠️ No videos found in playlist!", level="warning")
                    self.update_queue_listbox_threadsafe()
                except Exception as e:
                    self.log(f"❌ Error fetching playlist: {e}", level="error")
                    self.update_queue_listbox_threadsafe()
            # Start a new thread to fetch the playlist
            threading.Thread(target=fetch_playlist, daemon=True).start()
    
        else:
            # Normal single video handling
            task = {"query": query, "status": "Queued", "ydl_opts": self.global_ydl_opts.copy()}
            self.queue.append(task)
    
        self.update_queue_listbox_threadsafe()
        self.query_entry.delete(0, tk.END)
        self.log(f"✅ Added to queue: {query}")
        self.query_entry.focus_set()
    
    def clear_queue(self):
        """Clears selected items from queue; if none selected, clears everything."""
        selected_items = self.queue_table.selection()  # Get selected items from Treeview

        if selected_items:
            indices = []
            for item in selected_items:
                try:
                    if self.queue_table.item(item, "values")[1] == "Downloading":
                        query = self.queue_table.item(item, "values")[3]
                        self.log(f"⚠️ Cannot remove downloading item: {query}", level="warning")
                        continue
                    else:
                    # Assume the first column holds the serial number
                        index = int(self.queue_table.item(item, "values")[0]) - 1
                        indices.append(index)
                except Exception as e:
                    self.log(f"Error retrieving index for item {item}: {e}", level="error")
            # Remove duplicates and sort indices in descending order
            indices = sorted(set(indices), reverse=True)
            for index in indices:
                del self.queue[index]
            self.update_queue_listbox_threadsafe()
            self.log(f"✅ Removed {len(indices)} selected items from queue.")
        else:
            confirmation = messagebox.askyesno("Warning", "Do you want to clear the entire queue?")
            if confirmation:
                active_downloads = any(task["status"] == "Downloading" for task in self.queue)
                if active_downloads:
                    ask = messagebox.askyesnocancel("Warning", "Some downloads are still running and you cannot cancel them!\n\n"
                    "press 'Yes' to keep the downloads running in background and clear them from the queue\n"
                    "press 'No' to clear the other queue except leaving the downloads running\n"
                    "press 'Cancel' to keep the queue as it is")
                    if ask:
                        self.queue.clear()
                        self.update_queue_listbox_threadsafe()
                        self.log("✅ Cleared the entire queue while dowloading items running in background.")
                    elif ask == False:
                        for task in self.queue:
                            if task["status"] != "Downloading":
                                self.queue.remove(task)
                        self.update_queue_listbox_threadsafe()
                        self.log("✅ Cleared the entire queue except the Downloading items.")
                    elif ask == None:
                        pass
                else:
                    self.queue.clear()
                    self.update_queue_listbox_threadsafe()
                    self.log("✅ Cleared the entire queue.")
            else:
                pass
    def load_spreadsheet_threaded(self):
        def load_spreadsheet_worker():
            file_path = filedialog.askopenfilename(filetypes=[
                ("Spreadsheet Files", "*.csv *.xlsx *.xls"),
            ])
            if not file_path:
                return

            try:
                ext = os.path.splitext(file_path)[-1].lower()

                valid_count = 0
                if ext == ".csv":
                    # Only parse comma-separated queries, use global_ydl_opts
                    with open(file_path, 'r', encoding='utf-8') as f:
                        raw_content = f.read()
                    queries = [q.strip() for q in raw_content.split(",") if q.strip()]

                    for query in queries:
                        task = {
                            "query": query,
                            "status": "Queued",
                            "ydl_opts": self.global_ydl_opts.copy()
                        }
                        self.queue.append(task)
                        self.update_queue_listbox_threadsafe()
                        valid_count += 1

                elif ext in (".xlsx", ".xls"):
                    # Structured rows: one query per row, optional preset
                    df = pd.read_excel(file_path)
                    df.columns = [col.strip().lower() for col in df.columns]

                    for _, row in df.iterrows():
                        query = row.get('query') or row.get('url')
                        if isinstance(query, int) or isinstance(query, float):
                            query = str(query)
                        elif not isinstance(query, str) or not query.strip():
                            continue

                        query = query.strip()
                        preset = str(row.get('preset')).strip() if row.get('preset') else ""

                        # Load preset settings or fall back
                        if preset and preset.lower() != "default":
                            preset_path = os.path.join(self.preset_dir, f"{preset}.json")
                            if os.path.exists(preset_path):
                                with open(preset_path, "r") as f:
                                    ydl_opts = json.load(f)
                            else:
                                preset = "New/Unsaved"
                                ydl_opts = self.global_ydl_opts.copy()
                        else:
                            ydl_opts = self.global_ydl_opts.copy()

                        task = {
                            "query": query,
                            "status": "Queued",
                            "ydl_opts": ydl_opts
                        }
                        self.queue.append(task)
                        self.log(f"Added {query} with preset {preset} to queue.")
                        self.update_queue_listbox_threadsafe()
                        valid_count += 1

                else:
                    messagebox.showerror("Unsupported Format", "Please select a CSV or Excel file.")
                    return

                if valid_count == 0:
                    messagebox.showerror("No Valid Entries", "No valid queries or URLs found.")
                else:
                    self.log(f"✅ Loaded {valid_count} entries from {os.path.basename(file_path)}.")

            except Exception as e:
                messagebox.showerror("Error", f"Failed to process file:\n{e}")

        threading.Thread(target=load_spreadsheet_worker, daemon=True).start()


    #export function to export log into file
    def export_logs(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
        if not file_path:
            return
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                for line in self.log_text.get("1.0", tk.END).splitlines():
                    f.write(line + "\n")
            self.log(f"✅ Log exported to {file_path}")
        except Exception as e:
            self.log(f"❌ Error exporting log: {e}", level="error")
    def select_directory(self):
        directory = filedialog.askdirectory(title="Select Download Directory")
        if directory:
            self.download_directory = directory
            self.dir_label.config(text=directory)

    def start_downloads(self):
        """Start downloads using a fixed number of worker threads"""
        if not self.download_directory:
            tk.messagebox.showwarning("Warning", "Please select a download directory first.")
            return
        if not self.queue:
            tk.messagebox.showwarning("Warning", "The download queue is empty.")
            return

        # Put tasks in the download queue
        for task in self.queue:
            if task["status"] == "Queued":
                self.download_queue.put(task)

        # Start only 4 worker threads
        for _ in range(self.max_threads):
            t = threading.Thread(target=self.worker, daemon=True)
            t.start()
    def worker(self):
        """Worker thread to process downloads from the queue"""
        while not self.download_queue.empty():
            task = self.download_queue.get()  # Get next task
            self.download_task(task)  # Run the download
            self.download_queue.task_done()  # Mark as completed

    def download_task(self, task):
        """Download video/audio with metadata and proper error handling."""

        task["status"] = "Downloading"
        task["progress"] = "0% | --:--"
        self.update_queue_listbox_threadsafe()
        self.log(f"🔄 Starting download: {task['query']}")

        query = task["query"]

        # If the query is a search term, prepend ytsearch:
        if not (query.startswith("http://") or query.startswith("https://")):
            query = f"ytsearch:{query}"
        ydl_opts = task['ydl_opts']  # Use task-specific options

        #setting directory
        try:
            custom_path=task["ydl_opts"]['custom_file_path']
            if custom_path:
                ydl_opts['outtmpl'] = os.path.join(custom_path, '%(title)s.%(ext)s')
        except KeyError:
            ydl_opts['outtmpl'] = os.path.join(self.download_directory, '%(title)s.%(ext)s')
        except Exception as e:
            self.log(f"Setting default directroy to task {task} because of {e}", level="error")
            ydl_opts['outtmpl'] = os.path.join(self.download_directory, '%(title)s.%(ext)s')
        # #add download specific values
        ydl_opts.update({
            'progress_hooks': [lambda d: self.progress_hook(d, task)],
            'logger': self,
            'noplaylist': True  # Ensures only a single video is downloaded, not a playlist
            })
        #check if current ydl_opts has final_ext = orignial if yes then remove it from ydl_opts
        if ydl_opts.get('final_ext') == "original":
            ydl_opts.pop('final_ext')
        #managing commnets extraction : 
        if task["ydl_opts"].get('extract_comments', False):
            ydl_opts['extract_comments'] = True
            ydl_opts['writeinfojson'] = True        
        try:
            with youtube_dl.YoutubeDL(ydl_opts) as ydl:
                ydl.download([query])
            task["status"] = "Completed"
            task["progress"] = "100% | Done"            
        except Exception as e:
            task["status"] = f"Failed"
            task["progress"] = f"❌ Error"
            self.log(f"❌ Error downloading {query}: {e}", level="error")


        self.update_queue_listbox_threadsafe()
    def extract_comments_postprocessor(self, d):
        """Post-process comments extraction, including replies."""
        def extract_comments():
            if d['status'] == 'finished' and d.get('filename'):
                    with open(info_json_path, 'r', encoding='utf-8') as f:
                        info = json.load(f)
                    comments = info.get('comments', [])
                    if comments:
                        # Extract Comments and Replies
                        formatted_comments = []
                        for idx, comment in enumerate(comments, start=1):
                            text = comment.get('text', 'No text available')
                            author = comment.get('author', 'Unknown Author')
                            timestamp = comment.get('timestamp', 'Unknown Timestamp')
                            formatted_comments.append(f"{idx}. [{timestamp}] {author}: {text}")

                            # Handle replies if available
                            replies = comment.get('replies', [])
                            for reply_idx, reply in enumerate(replies, start=1):
                                reply_text = reply.get('text', 'No text available')
                                reply_author = reply.get('author', 'Unknown Author')
                                reply_timestamp = reply.get('timestamp', 'Unknown Timestamp')
                                formatted_comments.append(
                                    f"    ↳ {idx}.{reply_idx} [{reply_timestamp}] {reply_author}: {reply_text}"
                                )

                        # Save comments and replies to a file or process them as needed
                        comments_file_path = os.path.splitext(d['filename'])[0] + "_comments.txt"
                        with open(comments_file_path, 'w', encoding='utf-8') as cf:
                            for formatted_comment in formatted_comments:
                                cf.write(formatted_comment + "\n")
                        self.log(f"✅ Extracted {len(comments)} comments (including replies) to {comments_file_path}")            # Extract comments from the info.json file
                    else:
                        self.log("⚠️ No comments found in the video.", level="warning")
        info_json_path = os.path.splitext(d['filename'])[0] + ".info.json"
        # handling .f in filename to get the correct path of info.json file
        if ".f" in d['filename']:
            info_json_path = os.path.splitext(d['filename'])[0] + "info.json"
        if os.path.exists(info_json_path):
            extract_comments()
        else:
            #ask filepath manually if info.json file is not found
            info_json_path = filedialog.askopenfilename(title=f"info.json for {d.get('filename',"latest")}", filetypes=[("JSON Files", "*.json")])
            if info_json_path:
                extract_comments()
            else:
                self.log("❌ info.json file not found for comments Extraction.", level="error")

    def progress_hook(self, d, task):
        """Update progress percentage & ETA in Treeview."""
        if d['status'] == 'downloading':
            percent = d.get('_percent_str', '0%').strip()
            eta = d.get('_eta_str', '--:--')
            # Update task progress
            task["progress"] = f"{percent} | {eta}"
            self.update_queue_listbox_threadsafe()
        elif d['status'] == 'finished':
            # Call the comments extraction postprocessor if enabled
            if task["ydl_opts"].get('getcomments', False):
                self.extract_comments_postprocessor(d)
            task["status"] = "Completed"
            task["progress"] = "100% | Done"
            self.update_queue_listbox_threadsafe()
    def update_queue_listbox_threadsafe(self):
        def update_queue_listbox():
            existing_iids = set(self.queue_table.get_children())
            expected_iids = set()
    
            for i, task in enumerate(self.queue):
                iid = f"item_{i}"
                expected_iids.add(iid)
    
                status = task["status"]
                progress = task.get("progress", "0% | --:--")
                values = (i + 1, status, task["query"], progress)
    
                # Determine color tag
                color = "gray"
                if status == "Downloading":
                    color = "blue"
                elif status == "Completed":
                    color = "green"
                elif status.startswith("Failed"):
                    color = "red"
    
                if self.queue_table.exists(iid):
                    old_values = self.queue_table.item(iid, "values")
                    if old_values != values:
                        self.queue_table.item(iid, values=values, tags=(color,))
                else:
                    self.queue_table.insert("", "end", iid=iid, values=values, tags=(color,))
    
            # Remove orphaned rows
            for iid in existing_iids - expected_iids:
                self.queue_table.delete(iid)
    
            # Apply tag styles (once)
            self.queue_table.tag_configure("gray", foreground="gray")
            self.queue_table.tag_configure("blue", foreground="blue")
            self.queue_table.tag_configure("green", foreground="green")
            self.queue_table.tag_configure("red", foreground="red")
    
        self.master.after(10, update_queue_listbox)
    #function to load queue values and tasks from a json file
    def import_queue(self):
        file_path = filedialog.askopenfilename(title="Select Queue File", filetypes=[("JSON Files", "*.json")])
        if not file_path:
            return
        try:
            with open(file_path, "r") as f:
                self.queue = json.load(f)
                self.update_queue_listbox_threadsafe()
            self.log(f"✅ Loaded queue from {file_path}")
        except Exception as e:
            self.log(f"❌ Error loading queue: {e}", level="error")
    #function to export queue values and tasks to a json file
    def export_queue(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON Files", "*.json")])
        if not file_path:
            return
        try:
            with open(file_path, "w") as f:
                self.queue = [task for task in self.queue if task["status"] == "Queued"]
                json.dump(self.queue, f, indent=4)
            self.log(f"✅ Exported queue to {file_path}")
        except Exception as e:
            self.log(f"❌ Error exporting queue: {e}", level="error")

    #function to log messages into the log window
    def log(self, message, level="info"):
        """Improved Logging with Colors"""
        colors = {"info": "white", "success": "lightgreen", "warning": "orange", "error": "red"}
        
        self.log_text.insert(tk.END, message + "\n", level)
        self.log_text.tag_config(level, foreground=colors.get(level, "white"))
        if self.master.focus_get() == self.log_text:
            pass  # Don't auto-scroll if the log window is focused
        else:
            self.log_text.see(tk.END)  # Auto-scroll to the end
    #Functions required for youtube_dl logger
    def debug(self, msg): self.log("[DEBUG] " + msg, "info") 
    def success(self, msg): self.log("[SUCCESS] " + msg, "success") 
    def warning(self, msg): self.log("[WARNING] " + msg, "warning")
    def error(self, msg): self.log("[ERROR] " + msg, "error")
    def critical(self, msg): self.log("[CRITICAL] " + msg, "error")

    def close_application(self):
        """Gracefully closes the application, ensuring no active downloads are interrupted."""

        # ✅ Step 1: Check if any downloads are in progress
        active_downloads = any(task["status"] == "Downloading" for task in self.queue)

        if active_downloads:
            # Step 2: Prompt user for confirmation
            confirm = tk.messagebox.askyesnocancel(
                "Downloads in Progress",
                "⚠️ Some downloads are still running!\n\n"
                "Do you want to store queue and exit?"
            )

            if confirm is None:  # User clicked "Cancel"
                return  # Do nothing

            if confirm:  # User clicked "Yes" → Cancel all active downloads
                self.log("🔴 Storing queue and exiting...")
                #if any completed downloads, remove them from the queue
                self.queue = [task for task in self.queue if task["status"] != "Completed"]
                #call export_queue function to save the queue to file
                self.export_queue()           
                self.update_queue_listbox_threadsafe()
                self.log("🔴 Exiting application...")
                #stop any threads if running
                for thread in threading.enumerate():
                    if thread is not threading.current_thread():
                        thread.join(timeout=0.1)
                sys.exit()

            else:  #ser clicked "No" → Don't exit
                self.log("🔴 Exiting without saving")
                for thread in threading.enumerate():
                    if thread is not threading.current_thread():
                        thread.join(timeout=0.1)
                sys.exit()
        # ✅ Step 3: Perform Cleanup Before Exit
        self.log("🔴 Closing application...")
    
    # ✅ Step 4: Destroy the application window
        self.master.destroy()
        sys.exit()

    def minimize_to_tray(self):
        self.master.withdraw()  # Hide main window
        self.show_tray_icon()

    def show_tray_icon(self):
        image = Image.open("Assets/logo.ico")
        menu = Menu(
            item('Show App', self.restore_window,default=True),
            item('Exit', self.close_application),
        )
        self.tray_icon = TrayIcon("YT-DLP GUI", image, "Running in Background", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def restore_window(self, icon, item):
        self.master.after(0, self.master.deiconify)
        self.master.after(0, self.master.focus_force)
        self.master.state('zoomed')
        self.tray_icon.stop()
if __name__ == "__main__":
        root = tk.Tk()
        app = DownloaderApp(root)
        root.mainloop()
