import configparser
from pptx.util import Pt
import os
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox, MULTIPLE, Toplevel, Canvas, Scrollbar, Menu, StringVar
from tkinter.ttk import Progressbar, Label, Entry, Button, Checkbutton
from PIL import Image, ImageTk, ImageOps
from pptx import Presentation
from pptx.dml.color import RGBColor
import subprocess
from pathlib import Path
from io import BytesIO
import threading

class ConfigManager:
    def __init__(self, config_path='config.ini'):
        self.config_path = config_path
        self.config = configparser.ConfigParser()
        self.load_config()

    def load_config(self):
        if not os.path.exists(self.config_path):
            self.create_default_config()
        self.config.read(self.config_path)

        required_options = {
            'order_of_service': 'Opening Prayer, Song, Song, Giving, Song, Lords Supper, Song, Scripture Reading, Sermon, Song, Announcements, Song, Closing Prayer',
            'default_song_search_path': '/home/church/Music/PaperlessHymnal(Vol.1-13)/4x3',
            'add_song_title_slides': 'true',
            'announcement_slides_path': '/home/church/Desktop/Announcement slides.pptx',
            'temp_directory': '/tmp/ppt_combiner_temp',
            'song_title_slide_title': 'Let us sing:\n',
            'title_slide_trim_chars': '-,_,.ppt',
            'default_directory': '/home/church/Music/PaperlessHymnal(Vol.1-13)/4x3/',
            'unoconv_path': '/usr/bin/unoconv'
        }

        if 'Settings' not in self.config:
            self.config['Settings'] = {}

        updated = False
        for option, default_value in required_options.items():
            if option not in self.config['Settings']:
                self.config['Settings'][option] = default_value
                updated = True

        if updated:
            with open(self.config_path, 'w') as configfile:
                self.config.write(configfile)

    def create_default_config(self):
        self.config['Settings'] = {
            'order_of_service': 'Opening Prayer, Song, Song, Giving, Song, Lords Supper, Song, Scripture Reading, Sermon, Song, Announcements, Song, Closing Prayer',
            'default_song_search_path': '/home/church/Music/PaperlessHymnal(Vol.1-13)/4x3',
            'add_song_title_slides': 'true',
            'announcement_slides_path': '/home/church/Desktop/Announcement slides.pptx',
            'temp_directory': '/tmp/ppt_combiner_temp',
            'song_title_slide_title': 'Let us sing:\n',
            'title_slide_trim_chars': '-,_,.ppt',
            'default_directory': '/home/church/Music/PaperlessHymnal(Vol.1-13)/4x3/',
            'unoconv_path': '/usr/bin/unoconv'
        }
        with open(self.config_path, 'w') as configfile:
            self.config.write(configfile)

    def get(self, key):
        return self.config['Settings'].get(key)

    def set(self, key, value):
        self.config['Settings'][key] = value
        with open(self.config_path, 'w') as configfile:
            self.config.write(configfile)

class OptionsPage:
    def __init__(self, root, config_manager):
        self.root = root
        self.config_manager = config_manager
        self.options_dialog = Toplevel(root)
        self.options_dialog.title("Settings")

        # Configuration options
        Label(self.options_dialog, text="Order of Service:").grid(row=0, column=0, padx=10, pady=5, sticky='w')
        self.order_of_service_entry = Entry(self.options_dialog, width=50)
        self.order_of_service_entry.insert(0, self.config_manager.get('order_of_service'))
        self.order_of_service_entry.grid(row=0, column=1, padx=10, pady=5)

        Label(self.options_dialog, text="Default Song Search Path:").grid(row=1, column=0, padx=10, pady=5, sticky='w')
        self.default_song_search_path_entry = Entry(self.options_dialog, width=50)
        self.default_song_search_path_entry.insert(0, self.config_manager.get('default_song_search_path'))
        self.default_song_search_path_entry.grid(row=1, column=1, padx=10, pady=5)

        Label(self.options_dialog, text="Announcement Slides Path:").grid(row=2, column=0, padx=10, pady=5, sticky='w')
        self.announcement_slides_path_entry = Entry(self.options_dialog, width=50)
        self.announcement_slides_path_entry.insert(0, self.config_manager.get('announcement_slides_path'))
        self.announcement_slides_path_entry.grid(row=2, column=1, padx=10, pady=5)

        Label(self.options_dialog, text="Temporary Directory:").grid(row=3, column=0, padx=10, pady=5, sticky='w')
        self.temp_directory_entry = Entry(self.options_dialog, width=50)
        self.temp_directory_entry.insert(0, self.config_manager.get('temp_directory'))
        self.temp_directory_entry.grid(row=3, column=1, padx=10, pady=5)

        Label(self.options_dialog, text="Song Title Slide Title:").grid(row=4, column=0, padx=10, pady=5, sticky='w')
        self.song_title_slide_title_entry = Entry(self.options_dialog, width=50)
        self.song_title_slide_title_entry.insert(0, self.config_manager.get('song_title_slide_title'))
        self.song_title_slide_title_entry.grid(row=4, column=1, padx=10, pady=5)

        Label(self.options_dialog, text="Title Slide Trim Characters:").grid(row=5, column=0, padx=10, pady=5, sticky='w')
        self.title_slide_trim_chars_entry = Entry(self.options_dialog, width=50)
        self.title_slide_trim_chars_entry.insert(0, self.config_manager.get('title_slide_trim_chars'))
        self.title_slide_trim_chars_entry.grid(row=5, column=1, padx=10, pady=5)

        Label(self.options_dialog, text="unoconv Path:").grid(row=6, column=0, padx=10, pady=5, sticky='w')
        self.unoconv_path_entry = Entry(self.options_dialog, width=50)
        self.unoconv_path_entry.insert(0, self.config_manager.get('unoconv_path'))
        self.unoconv_path_entry.grid(row=6, column=1, padx=10, pady=5)

        self.add_song_title_slides_var = StringVar(value=self.config_manager.get('add_song_title_slides'))

        Label(self.options_dialog, text="Add Song Title Slides:").grid(row=7, column=0, padx=10, pady=5, sticky='w')
        Checkbutton(self.options_dialog, text="Enable", variable=self.add_song_title_slides_var, onvalue='true', offvalue='false').grid(row=7, column=1, padx=10, pady=5, sticky='w')

        Button(self.options_dialog, text="Save", command=self.save_settings).grid(row=8, column=0, columnspan=2, pady=10)

    def save_settings(self):
        self.config_manager.set('order_of_service', self.order_of_service_entry.get())
        self.config_manager.set('default_song_search_path', self.default_song_search_path_entry.get())
        self.config_manager.set('announcement_slides_path', self.announcement_slides_path_entry.get())
        self.config_manager.set('temp_directory', self.temp_directory_entry.get())
        self.config_manager.set('song_title_slide_title', self.song_title_slide_title_entry.get())
        self.config_manager.set('title_slide_trim_chars', self.title_slide_trim_chars_entry.get())
        self.config_manager.set('unoconv_path', self.unoconv_path_entry.get())
        self.config_manager.set('add_song_title_slides', self.add_song_title_slides_var.get())
        self.options_dialog.destroy()

class PPTCombinerApp:
    def __init__(self, root, config_manager):
        self.root = root
        self.root.title("Church Service Organizer")
        
        self.config_manager = config_manager
        self.order_of_service = config_manager.get('order_of_service').split(', ')
        self.default_directory = config_manager.get('default_song_search_path')
        self.unoconv_path = config_manager.get('unoconv_path')
        self.add_song_title_slides = config_manager.get('add_song_title_slides') == 'true'
        self.song_title_slide_title = config_manager.get('song_title_slide_title')
        self.title_slide_trim_chars = config_manager.get('title_slide_trim_chars').split(',')
        self.all_ppt_files = []
        self.filtered_ppt_files = []
        self.playlist = []

        self.setup_ui()
        self.load_order_of_service()  # Load the order of worship elements into the playlist manager window
        self.search_files()  # Run the search in the default directory on boot
    
    def setup_ui(self):
        # Create menu
        menu = Menu(self.root)
        self.root.config(menu=menu)
        
        options_menu = Menu(menu, tearoff=0)
        menu.add_cascade(label="Options", menu=options_menu)
        options_menu.add_command(label="Settings", command=self.open_options_dialog)

        frame = tk.Frame(self.root)
        frame.pack(fill=tk.BOTH, expand=True)

        left_frame = tk.Frame(frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        right_frame = tk.Frame(frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Directory selection
        self.dir_label = tk.Label(left_frame, text="Select Directory:", anchor="w")
        self.dir_label.pack(fill="x")
        
        self.dir_entry = tk.Entry(left_frame, width=50)
        self.dir_entry.insert(0, self.default_directory)
        self.dir_entry.pack(fill="x")
        
        self.browse_button = tk.Button(left_frame, text="Browse", command=self.browse_directory)
        self.browse_button.pack(fill="x")
        
        # Search input
        self.search_label = tk.Label(left_frame, text="Search:", anchor="w")
        self.search_label.pack(fill="x")
        
        self.search_entry = tk.Entry(left_frame, width=50)
        self.search_entry.pack(fill="x")
        self.search_entry.bind('<KeyRelease>', self.filter_files)
        
        # Add to playlist button
        self.add_to_playlist_button = tk.Button(left_frame, text="Add to Playlist", command=self.add_to_playlist)
        self.add_to_playlist_button.pack(fill="x")
        
        # File list
        self.file_listbox = Listbox(left_frame, selectmode=tk.BROWSE, width=100, height=20)
        self.file_listbox.pack(fill="both", expand=True)
        self.file_listbox.bind('<Double-1>', self.add_to_playlist)
        self.file_listbox.bind('<Control-1>', self.preview_slide)
        
        # Playlist
        self.playlist_label = tk.Label(right_frame, text="Playlist:", anchor="w")
        self.playlist_label.pack(fill="x")
        
        self.playlist_listbox = Listbox(right_frame, selectmode=tk.BROWSE, width=100, height=20)
        self.playlist_listbox.pack(fill="both", expand=True)

        self.move_up_button = tk.Button(right_frame, text="Move Up", command=self.move_up)
        self.move_up_button.pack(fill="x")

        self.move_down_button = tk.Button(right_frame, text="Move Down", command=self.move_down)
        self.move_down_button.pack(fill="x")

        self.remove_button = tk.Button(right_frame, text="Remove", command=self.remove_from_playlist)
        self.remove_button.pack(fill="x")
        
        # Save and Convert button
        self.save_convert_button = tk.Button(right_frame, text="Save and Convert", command=self.save_and_convert)
        self.save_convert_button.pack(fill="x")

    def load_order_of_service(self):
        for item in self.order_of_service:
            if item.lower() == "song":
                self.playlist.append(item)
                self.playlist_listbox.insert(tk.END, item)
            else:
                self.playlist.append(item)
                self.playlist_listbox.insert(tk.END, item)

    def browse_directory(self):
        directory = filedialog.askdirectory(initialdir=self.default_directory)
        if directory:
            self.dir_entry.delete(0, tk.END)
            self.dir_entry.insert(0, directory)
            self.search_files()  # Update the search results when a new directory is selected
    
    def search_files(self):
        directory = self.dir_entry.get()
        if not os.path.isdir(directory):
            messagebox.showerror("Error", "Invalid directory")
            return
        
        self.all_ppt_files = self.recursive_search(directory)
        self.filtered_ppt_files = self.all_ppt_files
        self.update_file_listbox(self.filtered_ppt_files)
    
    def recursive_search(self, directory):
        matches = []
        for root, _, files in os.walk(directory):
            for file in files:
                if file.endswith(('.ppt', '.pptx')):
                    matches.append(os.path.join(root, file))
        return sorted(matches)  # Sort the results
    
    def filter_files(self, event):
        query = self.search_entry.get().lower()
        self.filtered_ppt_files = [file for file in self.all_ppt_files if query in os.path.basename(file).lower()]
        self.update_file_listbox(self.filtered_ppt_files)
    
    def update_file_listbox(self, files):
        self.file_listbox.delete(0, tk.END)
        for file in files:
            self.file_listbox.insert(tk.END, os.path.basename(file))
    
    def add_to_playlist(self, event=None):
        if event:
            selected_indices = [self.file_listbox.nearest(event.y)]
        else:
            selected_indices = self.file_listbox.curselection()
        
        for index in selected_indices:
            file = self.filtered_ppt_files[index]
            if "song" in [item.lower() for item in self.playlist]:
                song_index = [i for i, item in enumerate(self.playlist) if item.lower() == "song"][0]
                self.playlist[song_index] = file
                self.playlist_listbox.delete(song_index)
                self.playlist_listbox.insert(song_index, os.path.basename(file))
    
    def move_up(self):
        selected_indices = self.playlist_listbox.curselection()
        for index in selected_indices:
            if index > 0:
                self.playlist[index], self.playlist[index-1] = self.playlist[index-1], self.playlist[index]
                self.playlist_listbox.delete(0, tk.END)
                for item in self.playlist:
                    self.playlist_listbox.insert(tk.END, os.path.basename(item))
                self.playlist_listbox.select_set(index-1)

    def move_down(self):
        selected_indices = self.playlist_listbox.curselection()
        for index in reversed(selected_indices):
            if index < len(self.playlist) - 1:
                self.playlist[index], self.playlist[index+1] = self.playlist[index+1], self.playlist[index]
                self.playlist_listbox.delete(0, tk.END)
                for item in self.playlist:
                    self.playlist_listbox.insert(tk.END, os.path.basename(item))
                self.playlist_listbox.select_set(index+1)
    
    def remove_from_playlist(self):
        selected_indices = self.playlist_listbox.curselection()
        for index in reversed(selected_indices):
            del self.playlist[index]
        self.playlist_listbox.delete(0, tk.END)
        for item in self.playlist:
            self.playlist_listbox.insert(tk.END, os.path.basename(item))
    
    def save_and_convert(self):
        desktop_path = str(Path.home() / "Desktop")
        save_path = filedialog.asksaveasfilename(initialdir=desktop_path, defaultextension=".pptx", filetypes=[("PowerPoint files", "*.pptx")])
        if not save_path:
            return

        # Create a new window for the loading indicator
        loading_window = Toplevel(self.root)
        loading_window.title("Saving and Converting")

        progress = Progressbar(loading_window, mode='indeterminate')
        progress.pack(pady=20)
        progress.start()

        def perform_conversion():
            self.combine_presentations(save_path)
            progress.stop()
            loading_window.destroy()

        threading.Thread(target=perform_conversion).start()
    
    def convert_ppt_to_pptx(self, file_path):
        if file_path.endswith('.pptx'):
            return file_path  # No conversion needed
        pptx_path = file_path.replace('.ppt', '.pptx')
        try:
            subprocess.run([self.unoconv_path, '-f', 'pptx', file_path], check=True)
        except FileNotFoundError:
            messagebox.showerror("Error", "unoconv not found. Please install unoconv to convert .ppt files.")
            return None
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Error", f"Conversion failed: {e}")
            return None
        return pptx_path
    
    def combine_presentations(self, save_path):
        combined_ppt = Presentation()
        temp_directory = self.config_manager.get('temp_directory')

        if not os.path.exists(temp_directory):
            os.makedirs(temp_directory)

        for item in self.playlist:
            if os.path.exists(item):  # This is a song file
                pptx_file = self.convert_ppt_to_pptx(item)
                if not pptx_file:
                    return

                # Add the song title slide if the option is enabled
                if self.add_song_title_slides:
                    song_title = os.path.basename(item).rsplit('.', 1)[0]
                    for char in self.title_slide_trim_chars:
                        song_title = song_title.rstrip(char)

                    title_slide = combined_ppt.slides.add_slide(combined_ppt.slide_layouts[5])
                    title_text_box = title_slide.shapes.add_textbox(0, 0, combined_ppt.slide_width, combined_ppt.slide_height)
                    title_text_frame = title_text_box.text_frame
                    title_text_frame.clear()
                    p = title_text_frame.paragraphs[0]
                    p.text = f"{self.song_title_slide_title}\n{song_title}"
                    p.font.size = Pt(44)
                    p.font.bold = True
                    p.font.color.rgb = RGBColor(0, 0, 0)
                    title_text_box.text_anchor = "middle"  # Center text vertically
                    title_text_box.left = int((combined_ppt.slide_width - title_text_box.width) / 2)  # Center text horizontally

                # Add the actual song slides
                ppt = Presentation(pptx_file)
                for slide in ppt.slides:
                    self.add_slide_to_presentation(combined_ppt, slide)

            else:  # This is a service item
                slide = combined_ppt.slides.add_slide(combined_ppt.slide_layouts[5])
                text_box = slide.shapes.add_textbox(0, 0, combined_ppt.slide_width, combined_ppt.slide_height)
                text_frame = text_box.text_frame
                text_frame.text = item
                text_frame.paragraphs[0].font.size = Pt(44)
                text_frame.paragraphs[0].font.bold = True
                text_box.text_anchor = "middle"  # Center text vertically
                text_box.left = int((combined_ppt.slide_width - text_box.width) / 2)  # Center text horizontally

        combined_ppt.save(save_path)
        messagebox.showinfo("Success", f"Combined presentation saved to {save_path}")

    def add_slide_to_presentation(self, combined_ppt, slide):
        slide_copy = combined_ppt.slides.add_slide(combined_ppt.slide_layouts[5])

        for shape in slide.shapes:
            if shape.has_text_frame:
                text_box = slide_copy.shapes.add_textbox(shape.left, shape.top, shape.width, shape.height)
                text_frame = text_box.text_frame
                text_frame.clear()
                for paragraph in shape.text_frame.paragraphs:
                    new_paragraph = text_frame.add_paragraph()
                    new_paragraph.text = paragraph.text
                    new_paragraph.font.color.rgb = RGBColor(0, 0, 0)
                    new_paragraph.font.size = paragraph.font.size
                    new_paragraph.font.bold = paragraph.font.bold
                    new_paragraph.font.italic = paragraph.font.italic
                    new_paragraph.font.underline = paragraph.font.underline
                    new_paragraph.font.name = paragraph.font.name

            elif shape.shape_type == 13:  # Picture
                image_stream = shape.image.blob
                image_file = BytesIO(image_stream)
                image = Image.open(image_file)
                temp_image_path = os.path.join(self.config_manager.get('temp_directory'), "temp_image.png")
                image.save(temp_image_path)
                slide_copy.shapes.add_picture(temp_image_path, shape.left, shape.top, shape.width, shape.height)

    def preview_slide(self, event):
        selected_index = self.file_listbox.nearest(event.y)
        selected_file = self.filtered_ppt_files[selected_index]
        file_name = os.path.basename(selected_file)

        preview_window = Toplevel(self.root)
        preview_window.title(file_name)
        
        progress = Progressbar(preview_window, mode='indeterminate')
        progress.pack(pady=20)
        progress.start()

        def load_preview():
            pptx_file = self.convert_ppt_to_pptx(selected_file)
            if not pptx_file:
                progress.stop()
                progress.pack_forget()
                return

            ppt = Presentation(pptx_file)
            slide = ppt.slides[0]  # Preview the first slide

            img = self.get_slide_image(slide)
            if img:
                img_tk = ImageTk.PhotoImage(img)
                canvas = Canvas(preview_window, width=img.width, height=img.height)
                canvas.create_image(0, 0, anchor='nw', image=img_tk)
                canvas.img = img_tk
                h_scroll = Scrollbar(preview_window, orient='horizontal', command=canvas.xview)
                v_scroll = Scrollbar(preview_window, orient='vertical', command=canvas.yview)
                canvas.config(scrollregion=(0, 0, img.width, img.height), xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)
                
                canvas.pack(side='left', fill='both', expand=True)
                h_scroll.pack(side='bottom', fill='x')
                v_scroll.pack(side='right', fill='y')

                canvas.bind("<MouseWheel>", lambda event: self.on_mouse_wheel(event, canvas))
                canvas.bind("<Shift-MouseWheel>", lambda event: self.on_shift_mouse_wheel(event, canvas))
            else:
                messagebox.showerror("Error", "No image found in the first slide.")
            
            progress.stop()
            progress.pack_forget()

        threading.Thread(target=load_preview).start()

    def on_mouse_wheel(self, event, canvas):
        if event.delta:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        elif event.num == 5:
            canvas.yview_scroll(1, "units")
        elif event.num == 4:
            canvas.yview_scroll(-1, "units")

    def on_shift_mouse_wheel(self, event, canvas):
        if event.delta:
            canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
        elif event.num == 5:
            canvas.xview_scroll(1, "units")
        elif event.num == 4:
            canvas.xview_scroll(-1, "units")

    def get_slide_image(self, slide, max_size=(1200, 900)):
        for shape in slide.shapes:
            if shape.shape_type == 13:  # Picture
                image_stream = shape.image.blob
                image_file = BytesIO(image_stream)
                image = Image.open(image_file)
                
                image.thumbnail(max_size, Image.LANCZOS)  # Scale down the image if it's too large
                
                return image
        return None

    def open_options_dialog(self):
        OptionsPage(self.root, self.config_manager)

if __name__ == "__main__":
    config_manager = ConfigManager()
    root = tk.Tk()
    app = PPTCombinerApp(root, config_manager)
    root.mainloop()

