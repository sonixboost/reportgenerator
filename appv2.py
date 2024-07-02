import customtkinter
from customtkinter import filedialog
import os
import mainv2

#################### SETTINGS ########################
customtkinter.set_appearance_mode("dark") # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("dark-blue") # Themes: "blue" (standard), "green", "dark-blue"
########################################################

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # configure window
        self.title("Nikki's report")
        self.geometry(f"{1100}x{580}")
        self.current_dir = os.getcwd()
        self.save_path = os.getcwd()
        self.custom_name = ""

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(5, weight=1)
        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="Easy report generator!", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 0))
        self.label_subtitle = customtkinter.CTkLabel(self.sidebar_frame, text="- by zi xu, for Nikki", font=customtkinter.CTkFont(size=12))
        self.label_subtitle.grid(row=1, column=0, padx=20, pady=(0,10))
        self.sidebar_button_1 = customtkinter.CTkButton(self.sidebar_frame, text="Open File", command=self.upload_action)
        self.sidebar_button_1.grid(row=2, column=0, padx=20, pady=10)
        self.sidebar_button_2 = customtkinter.CTkButton(self.sidebar_frame, text="Save To", command=self.save_to_action)
        self.sidebar_button_2.grid(row=3, column=0, padx=20, pady=10)
        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=6, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark", "System"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=7, column=0, padx=20, pady=(10, 10))
        self.scaling_label = customtkinter.CTkLabel(self.sidebar_frame, text="UI Scaling:", anchor="w")
        self.scaling_label.grid(row=8, column=0, padx=20, pady=(10, 0))
        self.scaling_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["80%", "90%", "100%", "110%", "120%"],
                                                               command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=9, column=0, padx=20, pady=(10, 20))

        # create main entry and button
        self.entry = customtkinter.CTkEntry(self, placeholder_text="Input custom file name...")
        self.entry.grid(row=3, column=1, columnspan=1, padx=(20, 20), pady=(20, 20), sticky="nsew")

        self.generate_button = customtkinter.CTkButton(self, text="Generate!", command=self.generate)
        self.generate_button.grid(row=3, column=2, columnspan=1, padx=(0, 20), pady=(20, 20), sticky="nsew")
        
        # create textbox and report
        self.textbox = customtkinter.CTkTextbox(self, width=250)
        self.textbox.grid(row=0, column=1, columnspan=2, padx=(20, 20), pady=(20, 0), sticky="nsew")
        self.result =  customtkinter.CTkLabel(self, text = "-", font=customtkinter.CTkFont(size=12), justify="left")
        self.result.grid(row=1, column=1, columnspan=2, padx=(20,20), pady=(10, 10), sticky="nw")

        # set default values
        self.generate_button.configure(state="disabled")
        self.textbox.insert("0.0", "File selected:\n" + mainv2.EXCEL_SHEET)
        self.textbox.insert("end", "\n\n\nSave to:\n" + self.save_path)
        self.textbox.configure(state="disabled")


    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)
    
    def update_textbox_word(self):
        self.textbox.configure(state="normal")
        self.textbox.delete("0.0", "end")
        self.textbox.insert("0.0", "File selected:\n" + mainv2.EXCEL_SHEET)
        self.textbox.insert("end", "\n\n\nSave to:\n" + mainv2.save_path)
        self.textbox.configure(state="disabled")

    def upload_action(self):
        self.filename = filedialog.askopenfilename(initialdir=self.current_dir)
        if not self.filename:
            return
        
        mainv2.EXCEL_SHEET = self.filename
        print('Selected:', mainv2.EXCEL_SHEET)
        self.update_textbox_word()
        self.sidebar_button_1.configure(text='Choose again')
        self.generate_button.configure(state="normal")

    def save_to_action(self):
        self.save_path = filedialog.askdirectory()
        mainv2.save_path = self.save_path
        self.update_textbox_word()

    def generate(self):
        if self.entry.get():
            self.custom_name = self.entry.get()

        mainv2.generate_report(self.custom_name)
        report = mainv2.report
        self.result.configure(text=report)
        

if __name__ == "__main__":
    app = App()
    app.mainloop()