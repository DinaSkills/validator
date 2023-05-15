import tkinter as tk
import tkinter.filedialog as fd
import tkinter.ttk as ttk
import os
from sheetValidation import *


"""
ExcelValidationApp is a class that creates a GUI application for validating Excel files.
The class uses the tkinter module for creating a graphical user interface (GUI).

@version 1.0

@author Dina Ahmetspahic

"""


class Main:

    """
    A GUI application for validating Excel files.

    Attributes:
    window (Tk): The main application window.
    selected_file (str): The path of the selected file.
    frame_header (Frame): The header frame of the application.
    unlock_visible (bool): A flag indicating whether the unlock button is visible.
    unlock_button (bool): A flag indicating whether the unlock button is enabled.
    alert_label (Label): The label for displaying alerts.
    btn_validate_file (Button): The button for validating the selected file.
    """

    def __init__(self):
        """
        Initializes the ExcelValidationApp class.
        """
        self.window = tk.Tk()
        self.window.title("Excel Validator")
        self.window.geometry("700x600")
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        self.window.rowconfigure(1, weight=8)
        self.window.rowconfigure(2, weight=1)

        self.selected_file = None
        self.frame_header = None
        self.selected_file = None
        self.unlock_visible = False
        self.unlock_button = False
        self.alert_label = None
        self.btn_validate_file = None

        self.create_header()

    def create_header(self):
        """
        Creates the header of the application.

        Args:
        self: The instance of the class.

        Returns:
        None
        """
        self.frame_header = tk.Frame(self.window)
        self.frame_header.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        self.frame_header.columnconfigure(0, weight=0)
        self.frame_header.columnconfigure(1, weight=1)
        self.frame_header.rowconfigure(0, weight=0)

        header_upload_label = tk.Label(
            self.frame_header,
            text="Upload Excel file:",
            anchor="w",
            font=("Open sans", 12),
        )
        header_upload_label.grid(column=0, row=0, sticky="ew")

        btn_style = ttk.Style()
        btn_style.configure(
            "TButton",
            relief="flat",
            padding=4,
            foreground="black",
            font=("Open sans", 12),
        )

        btn_choose_file = ttk.Button(
            self.frame_header,
            text="Choose file",
            command=self.open_file,
            style="TButton",
        )
        btn_choose_file.grid(row=0, column=1, sticky="ew")

    def open_file(self):
        """
        Opens a file explorer window to select an Excel file to validate.
        If the selected path is valid, sets the "selected_file" attribute and shows the "Validate File" button.

        Args:
        self: The instance of the class.

        Returns:
        None

        Raises:
        Exception -- throws error working with file
        """
        try:
            file_path = fd.askopenfile(
                mode="r", filetypes=[("Excel files", ".xlsx .xls")]
            )
            if not file_path:
                raise ValueError("Path is empty.")
            elif not os.path.exists(file_path.name):
                raise FileNotFoundError("Path does not exist.")
            elif not os.path.isfile(file_path.name):
                self.alert(text=(f"{file_path.name} is not a file."))
            elif os.path.isdir(file_path.name):
                self.alert(text=(f"{file_path.name} is a directory."))
            elif os.path.isfile(file_path.name):
                self.selected_file = file_path.name
                self.show_validation_button(self.unlock_button)

        except (ValueError, FileNotFoundError) as e:
            self.alert(text=f"Error: {e}")

    def show_validation_button(self, unlock_button):
        """
        Enables the "Validate File" button and sets the
        "unlock_button" attribute to True.

        Args:
        self: The instance of the class,

        Returns:
        button: The "Validate File" button widget

        """

        if not unlock_button:
            unlock_button = True
            file_name = self.selected_file.split("/")[-1]
            self.btn_validate_file = ttk.Button(
                self.frame_header,
                text=f"Validate {file_name}",
                style="TButton",
                command=lambda: self.create_scrollable_listbox(self.unlock_visible),
            )
            self.btn_validate_file.grid(row=1, column=0, columnspan=2, sticky="nsew")

    def create_scrollable_listbox(self, unlock_visible):
        """
        Creates the main frame of the application window where the validation
        results will be displayed.Creates a scrollable listbox widget
        to display the validation results.

        Args:
        self: The instance of the class.

        Returns:
        The created listbox widget.
        """
        if not unlock_visible:
            unlock_visible = True
            self.btn_validate_file.config(state=tk.DISABLED)
            self.validate = sheetValidation(self.selected_file)
            self.validate.run()
            main_frame = tk.Frame(self.window)
            main_frame.grid(row=1, column=0, sticky="nsew")
            scrollbar = ttk.Scrollbar(main_frame)
            scrollbar.grid(row=1, column=1, sticky="nsew")

            listbox = tk.Listbox(
                main_frame,
                bg="white",
                font=("Open sans", 14),
                yscrollcommand=scrollbar.set,
            )
            listbox.grid(row=1, column=0, ipady=100, sticky="nsew")
            listbox.configure(
                borderwidth=1,
                relief="solid",
                highlightthickness=0,
                bd=0,
                highlightbackground="red",
            )
            scrollbar.config(command=listbox.yview)
            main_frame.grid_columnconfigure(0, weight=5)
            main_frame.grid_rowconfigure(0, weight=5)

            validate_results = (
                self.validate.error_date
                + self.validate.error_default
                + self.validate.error_len
            )
            if validate_results:
                for i in validate_results:
                    listbox.insert(tk.END, i)
                self.frame_footer()
            else:
                listbox.insert(tk.END, "Excel file has no errors!")

            self.alert(text=self.validate.check_sheet_path())

    def frame_footer(self):
        """
        Creates the footer of the GUI, which includes two buttons
        for downloading the validated data as a txt file or an Excel file.

        Args:
        self: The instance of the class.

        Returns:
        None
        """
        main_footer = tk.Frame(self.window)
        main_footer.grid(row=2, column=0, columnspan=2, padx=10, pady=10)
        main_footer.grid_columnconfigure(0, weight=1)
        main_footer.grid_columnconfigure(1, weight=5)
        main_footer.grid_columnconfigure(2, weight=1)
        tk.Button(
            main_footer,
            text="Download file",
            font=("Open sans", 12),
            foreground="Black",
            command=lambda: (
                self.validate.write_to_file(),
                self.alert(text="Your file is downloded"),
            ),
        ).grid(row=2, column=1, sticky="e")
        tk.Button(
            main_footer,
            text="Download excel",
            font=("Open sans", 12),
            foreground="black",
            command=lambda: (
                self.validate.write_to_excel(),
                self.alert(text="Your file is downloded"),
            ),
        ).grid(row=2, column=3, padx=20, sticky="we")

    def alert(self, text):
        """
        The created listbox widget.

        Args:
        text (str): The text to display in the alert label.
        self: The instance of the class.

        Returns:
         Label: Text message.

        """
        if self.alert_label:
            self.alert_label.destroy()
        self.alert_label = tk.Label(
            self.frame_header, text=text, font=("Open sans", 12), foreground="green"
        )
        self.alert_label.grid(row=2, column=0)

    def run(self):
        """
        Runs the main event loop of the application.

        Args:
        self: The instance of the class.
        """
        self.window.mainloop()


if __name__ == "__main__":
    app = Main()
    app.run()
