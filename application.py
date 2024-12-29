import os
from threading import Thread

import ttkbootstrap as ttk
from ttkbootstrap import utility
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs.dialogs import Messagebox

from tkinter.filedialog import askopenfilename

from openpyxl.utils.exceptions import InvalidFileException

from MusicCue import MusicCue, SheetExistsException


class MusicCueGui(ttk.Frame):
    """
    Provides GUI for the Music Cue APP using two tabs
    """
    def __init__(self, master, **kwargs):
        super().__init__(master, padding=10, **kwargs)
        self.pack(fill=BOTH, expand=YES)

        self.widgets = {}

        self.music_cue = MusicCue()

        self.my_notebook = ttk.Notebook(self, width=1400, height=750)
        self.my_notebook.pack()
        self.my_notebook.bind('<<NotebookTabChanged>>', self.on_tab_change)

        self.create_general_tab()
        self.create_episode_tab()

    def on_tab_change(self, event):
        """
        Populates list of Episode name for 'Episode' combobox when tab button is pressed.
        """
        tab = event.widget.tab('current')['text']
        if tab == 'Episode':
            try:
                self.widgets['episode_cb']['values'] = self.music_cue.get_episode_names()
            except TypeError:
                Messagebox.ok('Error: First select Excel file', alert=True)

    def create_general_tab(self):
        """
        Create frames within this tab. The main purpose of this tab is to provide
        general functionalities.
        """
        tab = ttk.Frame(self.my_notebook)
        self.my_notebook.add(tab, text='General')

        # Label frame
        self.widgets['excel_select_lf_header'] = self.create_label_frame_header(tab, frame_text='File select')
        browse_btn = ttk.Button(
            master=self.widgets['excel_select_lf_header'],
            text="Browse",
            command=self.select_excel_file,
            width=10,
        )
        browse_btn.pack(side=LEFT, padx=5)
        hdr_txt = 'Please select Excel file'
        hdr = ttk.Label(master=self.widgets['excel_select_lf_header'], text=hdr_txt, width=50)
        hdr.pack(side=LEFT, pady=5, padx=10)

        # Label frame
        lf_header = self.create_label_frame_header(tab, frame_text='Additional sheets')
        create_sheets_btn = ttk.Button(
            master=lf_header,
            text="Create",
            command=self.create_additional_sheets,
            width=10,
        )
        create_sheets_btn.pack(side=LEFT, padx=5)
        hdr_txt = 'Create additional sheets'
        hdr = ttk.Label(master=lf_header, text=hdr_txt, width=50)
        hdr.pack(side=LEFT, pady=5, padx=10)

        update_sheets_btn = ttk.Button(
            master=lf_header,
            text="Update",
            command=self.update_additional_sheets,
            width=10,
        )
        update_sheets_btn.pack(side=LEFT, padx=5)
        hdr_txt = 'Update additional sheets'
        hdr = ttk.Label(master=lf_header, text=hdr_txt, width=50)
        hdr.pack(side=LEFT, pady=5, padx=10)

        # Label frame
        lf_header = self.create_label_frame_header(tab, frame_text='Directory structure')
        dir_btn_create = ttk.Button(
            master=lf_header,
            text="Create or Update",
            command=self.create_or_update_directory_structure,
            width=20,
        )
        dir_btn_create.pack(side=LEFT, padx=5)
        hdr_txt = 'Create or Update directory structure'
        hdr = ttk.Label(master=lf_header, text=hdr_txt, width=50)
        hdr.pack(side=LEFT, pady=5, padx=10)

    def create_episode_tab(self):
        """
        This tab provides functionalities to update information and to split Episode music files.
        """
        tab = ttk.Frame(self.my_notebook)
        self.my_notebook.add(tab, text='Episode')

        # Label frame
        lf_header = self.create_label_frame_header(tab, frame_text='Select Episode')
        self.widgets['episode_cb'] = ttk.Combobox(
            master=lf_header,
            values=[],
            width=25,
        )
        self.widgets['episode_cb'].pack(side=LEFT, padx=5)
        hdr_txt = 'Please select the episode'
        hdr = ttk.Label(master=lf_header, text=hdr_txt, width=50)
        hdr.pack(side=LEFT, pady=5, padx=10)

        # Label frame
        lf_header = self.create_label_frame_header(tab, frame_text='Get Episode data')
        episode_data_btn = ttk.Button(
            master=lf_header,
            text="Get",
            command=self.get_episode_data,
            width=10,
        )
        episode_data_btn.pack(side=LEFT, padx=5)
        hdr_txt = 'Get episode data'
        hdr = ttk.Label(master=lf_header, text=hdr_txt, width=50)
        hdr.pack(side=LEFT, pady=5, padx=10)

        # Label frame
        lf_header = self.create_label_frame_header(tab, frame_text='Episode data')
        columns = [0, 1, 2, 3, 4, 5]
        self.widgets['episode_result_view'] = ttk.Treeview(
            master=lf_header,
            columns=columns,
            show=HEADINGS
        )
        self.widgets['episode_result_view'].pack(fill=BOTH, expand=YES, pady=10)

        # setup columns
        self.widgets['episode_result_view'].heading(0, text='Event', anchor=W)
        self.widgets['episode_result_view'].heading(1, text='Clip name', anchor=W)
        self.widgets['episode_result_view'].heading(2, text='Start time', anchor=W)
        self.widgets['episode_result_view'].heading(3, text='End time', anchor=W)
        self.widgets['episode_result_view'].heading(4, text='Artist', anchor=W)
        self.widgets['episode_result_view'].heading(5, text='Title', anchor=W)
        for column in columns:
            self.widgets['episode_result_view'].column(
                column=column,
                anchor=W,
                width=utility.scale_size(self, 30),
                stretch=True
            )

        # Label frame
        lf_header = self.create_label_frame_header(tab, frame_text='Select Event')
        self.widgets['event_cb'] = ttk.Combobox(
            master=lf_header,
            values=[],
            width=25,
        )
        self.widgets['event_cb'].pack(side=LEFT, padx=5)
        hdr_txt = 'Please select the Event'
        hdr = ttk.Label(master=lf_header, text=hdr_txt, width=50)
        hdr.pack(side=LEFT, pady=5, padx=10)

        # Label frame
        lf_header = self.create_label_frame_header(tab, frame_text='Event management')
        edit_btn = ttk.Button(
            master=lf_header,
            text="Edit",
            command=self.get_artist_title_with_popup,
            width=10,
        )
        edit_btn.pack(side=LEFT, padx=5)
        hdr_txt = 'Edit Event data'
        hdr = ttk.Label(master=lf_header, text=hdr_txt, width=20)
        hdr.pack(side=LEFT, pady=5, padx=10)

        split_btn = ttk.Button(
            master=lf_header,
            text="Split",
            command=self.split_audio_file,
            width=10,
        )
        split_btn.pack(side=LEFT, padx=5)
        hdr_txt = 'Split audio file'
        hdr = ttk.Label(master=lf_header, text=hdr_txt, width=20)
        hdr.pack(side=LEFT, pady=5, padx=10)

    def get_episode_data(self):
        """
        Get Episode data from dB sheet and present in Tree view format.
        """
        self.clear_treeview(self.widgets['episode_result_view'])

        episode_name = self.widgets['episode_cb'].get()
        # Populate event combobox list.
        self.widgets['event_cb']['values'] = self.music_cue.get_events(self.widgets['episode_cb'].get())
        for row in self.music_cue.read_db_sheet():
            if row.episode_name != episode_name:
                continue
            artist = row.artist if row.artist else ''
            title = row.title if row.title else ''
            iid = self.widgets['episode_result_view'].insert(
                parent='',
                index=END,
                values=(row.event, row.clip_name, row.start, row.end, artist, title,),
            )
            self.widgets['episode_result_view'].selection_set(iid)
            self.widgets['episode_result_view'].see(iid)

    def create_or_update_directory_structure(self):
        """
        Generates subdirectories to store music files
        """
        try:
            os.chdir(self.music_cue.project_root_dir)
            for episode_name in self.music_cue.get_episode_names():
                os.makedirs(episode_name, exist_ok=True)
        except TypeError:
            Messagebox.ok('Error: First select Excel file', alert=True)
        finally:
            os.chdir(self.music_cue.ROOT_DIR)

    def split_audio_file(self):
        """
        Split music files. A Thread is used to prevent the GUI from getting frozen.
        """
        episode_name = self.widgets['episode_cb'].get()
        Thread(
            target=self.music_cue.split_mp4_file,
            args=(episode_name,),
            daemon=True,
        ).start()

    def get_artist_title_with_popup(self):
        """
        Input widget to store artist and title information.
        """
        self.widgets['artist_title_popup'] = ttk.Toplevel(size=(300, 200),)
        self.widgets['artist_title_popup'].title("Enter Artist and Title info")

        label = ttk.Label(self.widgets['artist_title_popup'], text="Enter Artist:")
        label.pack(pady=5)
        self.widgets['artist_entry'] = ttk.Entry(self.widgets['artist_title_popup'], width=50)
        self.widgets['artist_entry'].pack(pady=5)
        label = ttk.Label(self.widgets['artist_title_popup'], text="Enter Title:")
        label.pack(pady=5)
        self.widgets['title_entry'] = ttk.Entry(self.widgets['artist_title_popup'], width=50)
        self.widgets['title_entry'].pack(pady=5)

        submit_button = ttk.Button(
            self.widgets['artist_title_popup'], text="Submit", command=self.update_artist_title_info
        )
        submit_button.pack(pady=10, padx=10)

    def update_artist_title_info(self):
        """
        Updates dB sheet with information entered via input widget
        """
        clip_name = self.music_cue.get_clip_name_from_event(self.widgets['event_cb'].get())
        self.music_cue.update_artist_title_info(
            clip_name, self.widgets['artist_entry'].get(), self.widgets['title_entry'].get()
        )
        self.widgets['artist_title_popup'].destroy()

    def select_excel_file(self):
        """
        Callback after opening Excel file.
        """
        self.music_cue.excel_filename = askopenfilename(title="Select Excel file")
        try:
            self.music_cue.open_excel_document(self.music_cue.excel_filename)
            self.music_cue.project_root_dir = self.music_cue.excel_filename.rsplit('/', 1)[0]
            Messagebox.ok('File has been successfully selected')
            hdr_txt = f'Excel file {self.music_cue.excel_filename} has been selected'
            hdr = ttk.Label(master=self.widgets['excel_select_lf_header'], text=hdr_txt, width=150)
            hdr.pack(side=LEFT, pady=5, padx=10)
        except InvalidFileException:
            Messagebox.ok('Error: Invalid File type or None selected!', alert=True)
        except TypeError:
            Messagebox.ok('Error: First select Excel file', alert=True)

    def create_additional_sheets(self):
        """
        Additional sheets are created when project is initialized.
        """
        try:
            self.music_cue.create_or_update_db_sheet()
            self.music_cue.create_clip_usage_per_episode_sheet()
            self.music_cue.create_episode_tabs()
            Messagebox.ok('Additional sheets have been successfully created')
        except TypeError:
            Messagebox.ok(
                'Error: First Select Excel file',
                alert=True
            )
        except SheetExistsException:
            Messagebox.ok(
                'Error: Sheet already exists. Remove current dB sheet if you want to create a new one',
                alert=True
            )

    def update_additional_sheets(self):
        """
        dG sheet is updated when new Episodes have been added to the source sheet
        """
        try:
            self.music_cue.create_or_update_db_sheet(update=True)
            self.music_cue.create_clip_usage_per_episode_sheet()
            self.music_cue.create_episode_tabs()
            Messagebox.ok('Sheets have been successfully updated')
        except TypeError:
            Messagebox.ok(
                'Error: First Select Excel file',
                alert=True
            )
        except SheetExistsException:
            Messagebox.ok(
                'Error: Sheet already exists. Remove current dB sheet if you want to create a new one',
                alert=True
            )

    @staticmethod
    def create_label_frame_header(root, frame_text):
        """
        Helper function to create a ttk Frame
        """
        lf = ttk.Labelframe(root, text=frame_text, padding=5)
        lf.pack(fill=X, pady=15, padx=10)

        lf_header = ttk.Frame(lf, padding=10)
        lf_header.pack(fill=X)

        return lf_header

    @staticmethod
    def clear_treeview(tree):
        """
        Clears Tree view data
        """
        for i in tree.get_children():
            tree.delete(i)
