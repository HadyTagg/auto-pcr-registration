import tkinter as tk
import tkinter.ttk
import tkinter.messagebox
import tkinter.simpledialog
import sqlite3
import pyautogui
import webbrowser
import time
import datetime
from openpyxl import load_workbook


class DatabaseManager:
    def __init__(self):
        self.connection = sqlite3.connect('auto_pcr_registration.db')
        self.cursor = self.connection.cursor()

    # CREATE TABLES
    def create_tables(self):
        self.cursor.execute("""CREATE TABLE IF NOT EXISTS person_details (
                                id integer PRIMARY KEY,
                                first_name text,
                                last_name text,
                                dob text,
                                gender text,
                                mobile text,
                                postcode text,
                                first_line_address text
                            )""")

        self.connection.commit()

    # ADD PERSON TO DATABASE
    def add_person_to_database(self, first_name, last_name, dob, gender, mobile, postcode, first_line_address,):
        self.cursor.execute("INSERT INTO person_details (first_name, last_name, dob, gender, mobile,"
                            " postcode, first_line_address) VALUES (?, ?, ?, ?, ?, ?, ?)",
                            (first_name,
                                last_name,
                                dob,
                                gender,
                                mobile,
                                postcode,
                             first_line_address))
        self.connection.commit()

    # COLLECT PERSONS
    def collect_persons(self):
        self.cursor.execute("SELECT first_name, last_name FROM person_details")
        self.connection.commit()
        return self.cursor.fetchall()

    # COLLECT SPREADSHEET INFO
    def collect_spreadsheet_info(self, first_name, last_name):
        self.cursor.execute("SELECT * FROM person_details WHERE first_name == (?) and last_name == (?);", (first_name,
                                                                                                           last_name))
        self.connection.commit()
        return self.cursor.fetchall()


class WindowManager:
    def __init__(self, master, title, geometry, previous_window):
        self.master = master
        self.window = tk.Toplevel(self.master)
        self.window.geometry(geometry)
        self.window.iconbitmap(r"icon.ico")
        self.window.title(title)
        self.window.focus_force()
        self.window.resizable(False, False)
        self.window.protocol('WM_DELETE_WINDOW', self.on_exit)
        self.previous_window = previous_window

    # ON WINDOW EXIT
    def on_exit(self):
        self.window.destroy()
        if self.previous_window.title() == 'tk':
            self.previous_window.destroy()
        else:
            self.previous_window.deiconify()

    # CENTERS THE WINDOW
    def center_window(self, x, y):
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x_coordinate = int((screen_width / 2) - (x / 2))
        y_coordinate = int((screen_height / 2) - (y / 2))
        self.window.geometry("{}x{}+{}+{}".format(x, y, x_coordinate, y_coordinate))

    # CREATE A BUTTON
    def make_button(self, text, command, state, pad_x, pad_y, side):
        button = tk.Button(master=self.window, text=text, command=command, state=state)
        button.pack(padx=pad_x, pady=pad_y, side=side)
        return button

    # CREATE A LISTBOX
    def make_listbox(self, height, width, pad_x, pad_y, side):
        listbox = tk.Listbox(master=self.window, width=width, height=height, exportselection=False,
                             highlightcolor='blue', highlightthickness=2, bd=4)
        listbox.pack(padx=pad_x, pady=pad_y, side=side)
        return listbox

    # CREATE A LABEL
    def make_label(self, text, pad_x, pad_y):
        label = tk.Label(master=self.window, text=text)
        label.pack(padx=pad_x, pady=pad_y)
        return label

    # CREATE AN ENTRY FIELD
    def make_entry_field(self):
        entry_field = tk.Entry(master=self.window)
        entry_field.pack()
        return entry_field

    # CREATE A SCALE
    def make_scale(self):
        scale = tk.Scale(master=self.window, orient='horizontal', activebackground='Blue',
                         to=1000, width=15, length=250, troughcolor='White', sliderlength=30, resolution=0.5)
        scale.pack()
        return scale

    # CREATE A COMBO BOX
    def make_combo_box(self, list_items):
        combo_box = tkinter.ttk.Combobox(master=self.window, state='readonly', values=list_items)
        combo_box.pack()
        return combo_box

    # CREATE A TEXT BOX
    def make_text_box(self, width, height):
        text_box = tk.Text(master=self.window, width=width, height=height)
        text_box.pack()
        return text_box

    # CREATE A MESSAGE BOX
    @staticmethod
    def make_message_box(title, message, icon):
        tk.messagebox.showinfo(title=title, message=message, icon=icon)


class AutoPCRRegistration(WindowManager):
    def __init__(self, master, title, geometry, previous_window):
        super().__init__(master, title, geometry, previous_window)

        self.collected_registrations = []
        self.collected_registrant_barcode_numbers = []
        self.collected_swab_time = []
        self.collected_am_pm = []
        self.swab_time_prompt = ''
        self.am_or_pm_prompt = ''

        self.email_address_for_results = ''
        self.approved_email_addresses = ['owennolan@marthatrust.org.uk', 'hadytagg@marthatrust.org.uk',
                                         'helenchantler@marthatrust.org.uk', 'kiristammers@marthatrust.org.uk',
                                         'jasonelliott@marthatrust.org.uk', 'emmabulloch@marthatrust.org.uk',
                                         'clairedoe@marthatrust.org.uk']

        self.url_for_registration = \
            'https://organisations.test-for-coronavirus.service.gov.uk/register-organisation-tests/consent'

# WIDGETS
# WINDOW, WINDOW CONFIGS, LABEL, COMBOBOX, AND BUTTONS
        WindowManager.make_label(self, text='Person Selection', pad_x=0, pad_y=5)
        self.person_selection_combobox = WindowManager.make_combo_box(
            self, list_items=[''] + sorted(DatabaseManager.collect_persons(DatabaseManager())))

        self.run_button = WindowManager.make_button(self, text='Run', command=self.run,
                                                    state='active', side=tk.BOTTOM, pad_x=5, pad_y=5)

        self.clear_all_registrations_button = WindowManager.make_button(
            self, text='Clear All Registrations', command=self.clear_all_registrations, state='active',
            side=tk.BOTTOM, pad_x=5, pad_y=5)

        self.clear_last_registration_button = WindowManager.make_button(
            self, text='Clear Last Registration', command=self.clear_last_registration, state='active',
            side=tk.BOTTOM, pad_x=5, pad_y=5)

        self.add_registration_button = WindowManager.make_button(
            self, text='Add PCR Registration', command=self.add_pcr_registration, state='active', side=tk.BOTTOM,
            pad_x=5, pad_y=5)

        self.currently_added_registrations_window = tk.Toplevel(self.window)
        self.currently_added_registrations_window.geometry('350x800')
        self.currently_added_registrations_window.iconbitmap('icon.ico')
        self.currently_added_registrations_window.title('Currently Added Registrants')
        self.currently_added_registrations_window.protocol('WM_DELETE_WINDOW',
                                                           self.on_attempted_show_registrations_window_exit)

        self.window.update()
        self.currently_added_registrations_window.geometry(
            "+%d+%d" % (self.window.winfo_rootx() + 1000, self.window.winfo_rooty() + 0))

        self.currently_added_text_box = tk.Text(self.currently_added_registrations_window, width='350', height='700',
                                                state=tk.DISABLED)
        self.currently_added_text_box.pack()

# BUTTON COMMANDS
    def clear_last_registration(self):
        confirmation = tk.messagebox.askyesno("Confirmation", "Are you sure you want to clear the last registration?")
        if confirmation:
            self.collected_registrant_barcode_numbers.pop()
            self.collected_registrations.pop()
            self.collected_swab_time.pop()
            self.collected_am_pm.pop()
            self.populate_show_registrations_text_box()

    def clear_all_registrations(self):
        confirmation = tk.messagebox.askyesno("Confirmation", "Are you sure you want to clear ALL registrations?")
        if confirmation:
            self.collected_registrant_barcode_numbers = []
            self.collected_registrations = []
            self.collected_swab_time = []
            self.collected_am_pm = []
            self.populate_show_registrations_text_box()

    def add_pcr_registration(self):
        self.disable_buttons_combobox_main_window()

        selected_person_name = self.person_selection_combobox.get()

        if selected_person_name and selected_person_name not in self.collected_registrations:
            registrant_barcode = tk.simpledialog.askstring(
                'Add Registration', f'Scan barcode for {selected_person_name}', parent=self.window)
            barcode_verification = tk.simpledialog.askstring(
                'Add Registration', f'Please verify barcode for {selected_person_name}', parent=self.window)

            if registrant_barcode is None or barcode_verification is None:
                WindowManager.make_message_box(
                    title='Error', message='You did not enter a valid barcode. No registration was added.',
                    icon='error')
                self.enable_buttons_combobox_main_window()

            if registrant_barcode == barcode_verification and registrant_barcode.isalnum():
                self.swab_time_prompt = tk.simpledialog.askstring(
                    'Add Registration', f'What hour was {selected_person_name} swabbed?', parent=self.window)
                if self.swab_time_prompt not in ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12']:
                    WindowManager.make_message_box(
                        title='Error', message='You did not enter an hour between 1-12. No registration was added.',
                        icon='error')
                    self.enable_buttons_combobox_main_window()

                if self.swab_time_prompt in ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12']:
                    self.am_or_pm_prompt = tk.simpledialog.askstring('Add Registration', 'AM or PM?',
                                                                     parent=self.window)
                    if self.am_or_pm_prompt not in ['am', 'pm', 'AM', 'PM']:
                        WindowManager.make_message_box(
                            title='Error', message='You must input either AM or PM. No registration was added.',
                            icon='error')
                        self.enable_buttons_combobox_main_window()
                    else:
                        self.collected_registrant_barcode_numbers.append(registrant_barcode)
                        self.collected_registrations.append(self.person_selection_combobox.get())
                        self.collected_swab_time.append(self.swab_time_prompt)
                        self.collected_am_pm.append(self.am_or_pm_prompt.lower())
                        self.enable_buttons_combobox_main_window()
                        self.populate_show_registrations_text_box()
                        WindowManager.make_message_box(
                            title='Success', message=f'Registration added for {self.person_selection_combobox.get()}.',
                            icon='info')
            else:
                self.enable_buttons_combobox_main_window()
                WindowManager.make_message_box(
                    title='Error', message='You did not enter a valid barcode. No registration was added.',
                    icon='error')
        else:
            self.enable_buttons_combobox_main_window()
            WindowManager.make_message_box(title='Error', message='Select a valid person.',
                                           icon='error')

    def run(self):
        self.disable_buttons_combobox_main_window()

        confirmation = tk.messagebox.askyesno("Confirmation", "Are you sure you want to run Auto PCR automation?")

        if confirmation and len(self.collected_registrations) > 0:
            if self.populate_spreadsheet():
                self.auto_gui()
        else:
            if confirmation is False:
                WindowManager.make_message_box(
                    title='Error', message='Auto PCR Registration automation was cancelled.', icon='error')

            elif len(self.collected_registrations) == 0:
                WindowManager.make_message_box(
                    title='Error', message='A registrant must be added before running '
                                           'Auto PCR Registration automation.', icon='error')
            else:
                WindowManager.make_message_box(
                    title='Error', message='An error occurred. Automation cancelled.', icon='error')
            self.enable_buttons_combobox_main_window()

# METHODS
    def disable_buttons_combobox_main_window(self):
        self.run_button.config(state='disabled')
        self.add_registration_button.config(state='disabled')
        self.person_selection_combobox.config(state='disabled')
        self.clear_last_registration_button.config(state='disabled')
        self.clear_all_registrations_button.config(state='disabled')

    def enable_buttons_combobox_main_window(self):
        self.run_button.config(state='normal')
        self.add_registration_button.config(state='normal')
        self.person_selection_combobox.config(state='readonly')
        self.clear_last_registration_button.config(state='normal')
        self.clear_all_registrations_button.config(state='normal')

    def on_attempted_show_registrations_window_exit(self):
        print(self)
        WindowManager.make_message_box(
            title='Error', message='This window can\'t be closed.', icon='error')

    def populate_show_registrations_text_box(self):
        currently_added_registrants_barcodes_list_numbers = list(
            zip([*range(1, len(self.collected_registrations) + 1)], self.collected_registrations,
                self.collected_registrant_barcode_numbers, self.collected_swab_time, self.collected_am_pm))
        self.currently_added_text_box.config(state=tk.NORMAL)
        self.currently_added_text_box.delete(1.0, tk.END)
        for entry in currently_added_registrants_barcodes_list_numbers:
            self.currently_added_text_box.insert(tk.END, entry)
            self.currently_added_text_box.insert(tk.END, '\n')
        self.currently_added_text_box.config(state=tk.DISABLED)

    def populate_spreadsheet(self):
        self.email_address_for_results = tk.simpledialog.askstring(
            'Email Address', 'Enter an approved email address. This is where the results will be emailed.',
            parent=self.window)

        if self.email_address_for_results and self.email_address_for_results in self.approved_email_addresses:

            workbook = load_workbook(filename='spreadsheet_template.xlsx')
            sheet = workbook.get_sheet_by_name('Enter details here')
            row = '4'
            counter = 0

            for entry in self.collected_registrations:
                # IDENTIFIERS
                first_name = entry.split()[0]
                last_name = entry.split()[1]

                db_first_name = DatabaseManager.collect_spreadsheet_info(DatabaseManager(), first_name, last_name)[0][1]
                db_last_name = DatabaseManager.collect_spreadsheet_info(DatabaseManager(), first_name, last_name)[0][2]
                db_dob = DatabaseManager.collect_spreadsheet_info(DatabaseManager(), first_name, last_name)[0][3]
                db_gender = DatabaseManager.collect_spreadsheet_info(DatabaseManager(), first_name, last_name)[0][4]
                db_mobile = DatabaseManager.collect_spreadsheet_info(DatabaseManager(), first_name, last_name)[0][5]
                db_postcode = DatabaseManager.collect_spreadsheet_info(DatabaseManager(), first_name, last_name)[0][6]
                db_first_line_address = DatabaseManager.collect_spreadsheet_info(DatabaseManager(), first_name,
                                                                                 last_name)[0][7]

                # DEFAULT ENTRIES
                sheet['A' + row] = 'PCR test'
                sheet["E" + row] = 'No'
                sheet["L" + row] = 'Prefer not to say'
                sheet["N" + row] = 'England'
                sheet["Q" + row] = 'Prefer not to say'
                sheet["W" + row] = 'No'
                sheet["U" + row] = self.email_address_for_results

                # PERSONALISED ENTRIES
                sheet["B" + row] = self.collected_registrant_barcode_numbers[counter]
                sheet["C" + row] = datetime.datetime.now().date()
                sheet["D" + row] = f'{self.collected_swab_time[counter]} {self.collected_am_pm[counter]}'
                sheet["H" + row] = db_first_name
                sheet["I" + row] = db_last_name
                sheet["J" + row] = db_dob
                sheet["K" + row] = db_gender
                sheet["O" + row] = db_postcode
                sheet["P" + row] = db_first_line_address
                sheet["V" + row] = db_mobile

                row = str(int(row) + 1)
                counter += 1
            workbook.save(filename='populated_spreadsheet.xlsx')
            self.enable_buttons_combobox_main_window()
            WindowManager.make_message_box(
                title='Success', message='Automation will now proceed. Do not interrupt the process!', icon='info')
            return True
        else:
            WindowManager.make_message_box(
                title='Error', message='Please enter an approved email address.', icon='error')
            self.enable_buttons_combobox_main_window()

    def auto_gui(self):
        initial_load_time = 10
        between_page_time = 0.5
        regular_interval_time = 0.2

        # LOAD WEBPAGE
        webbrowser.open(self.url_for_registration)
        time.sleep(initial_load_time)

        # DEAL WITH COOKIES
        accept_cookies_found = pyautogui.locateCenterOnScreen(image='images/accept_cookies.png', confidence=0.8)

        if accept_cookies_found:
            pyautogui.click(accept_cookies_found)
            time.sleep(regular_interval_time)
            hide_button_found = pyautogui.locateCenterOnScreen(image='images/hide.png', confidence=0.8)
            pyautogui.click(hide_button_found)

        # AUTHORISATION SCREEN
        confirm_authorisation_found = pyautogui.locateCenterOnScreen(image='images/authorisation_confirmation.png',
                                                                     confidence=0.8)
        time.sleep(regular_interval_time)
        pyautogui.click(confirm_authorisation_found)
        time.sleep(regular_interval_time)
        pyautogui.press('tab', presses=1)
        time.sleep(regular_interval_time)
        pyautogui.press('space', presses=1)
        time.sleep(regular_interval_time)
        pyautogui.press('tab', presses=1)
        time.sleep(regular_interval_time)
        pyautogui.press('space', presses=1)
        time.sleep(between_page_time)

        # ORGANISATION NUMBER WINDOW
        pyautogui.press('tab', presses=6)
        time.sleep(regular_interval_time)
        pyautogui.write('10112076')
        time.sleep(regular_interval_time)
        pyautogui.press('tab', presses=1)
        time.sleep(regular_interval_time)
        pyautogui.press('space', presses=1)
        time.sleep(0.1)

        # CHECK YOUR ORGANISATION'S DETAILS WINDOW
        pyautogui.press('tab', presses=5)
        time.sleep(regular_interval_time)
        pyautogui.press('space', presses=1)
        time.sleep(regular_interval_time)
        pyautogui.press('tab', presses=1)
        time.sleep(regular_interval_time)
        pyautogui.press('space', presses=1)
        time.sleep(0.1)

        # SPREADSHEET SELECTION WINDOW
        pyautogui.press('tab', presses=5)
        time.sleep(regular_interval_time)
        pyautogui.press('space', presses=1)
        time.sleep(regular_interval_time)
        pyautogui.press('tab', presses=1)
        time.sleep(regular_interval_time)
        pyautogui.press('space', presses=1)
        time.sleep(between_page_time)

        pyautogui.press('tab', presses=5)
        time.sleep(regular_interval_time)
        pyautogui.press('space', presses=1)
        time.sleep(regular_interval_time)
        pyautogui.press('tab', presses=1)
        time.sleep(regular_interval_time)
        pyautogui.press('space', presses=1)
        time.sleep(between_page_time)

        pyautogui.press('tab', presses=6)
        time.sleep(regular_interval_time)
        pyautogui.press('space', presses=1)
        time.sleep(regular_interval_time)

        populated_spread_sheet_found = None
        while populated_spread_sheet_found is None:
            populated_spread_sheet_in_dir_found = pyautogui.locateCenterOnScreen(
                image='images/populated_spread_sheet_in_dir.png', confidence=0.8)

            if populated_spread_sheet_in_dir_found:
                pyautogui.doubleClick(populated_spread_sheet_in_dir_found)
            populated_spread_sheet_found = pyautogui.locateOnScreen(image='images/populated_spread_sheet.png',
                                                                    confidence=0.8)
        time.sleep(3)

        pyautogui.press('tab', presses=1)
        time.sleep(regular_interval_time)
        pyautogui.press('space', presses=1)
        time.sleep(between_page_time)

        # BAR CODE PAGE
        pyautogui.press('tab', presses=5, interval=regular_interval_time)
        time.sleep(regular_interval_time)

        pyautogui.hotkey('ctrl', '+')
        pyautogui.hotkey('ctrl', '+')
        pyautogui.hotkey('ctrl', '+')
        pyautogui.hotkey('ctrl', '+')
        pyautogui.hotkey('ctrl', '+')
        pyautogui.hotkey('ctrl', '+')

        counter = 0
        for entry in self.collected_registrant_barcode_numbers:
            print(entry)
            confirm_button_found = pyautogui.locateCenterOnScreen(image='images/confirm.png', confidence=0.8)

            if confirm_button_found:
                break

            time.sleep(regular_interval_time)
            pyautogui.press('tab', presses=6, interval=regular_interval_time)
            time.sleep(regular_interval_time)

            pyautogui.press('tab', presses=1)
            time.sleep(regular_interval_time)

            yellow_next_page_found = pyautogui.locateCenterOnScreen(image='images/yellow_next_page_250.png',
                                                                    confidence=0.8)
            next_page_found = pyautogui.locateCenterOnScreen(image='images/next_page_250.png', confidence=0.8)

            if yellow_next_page_found:
                time.sleep(regular_interval_time)
                confirm_all_details_found = pyautogui.locateCenterOnScreen(image='images/confirm_all_details_250.png',
                                                                           confidence=0.8)
                time.sleep(regular_interval_time)
                pyautogui.click(confirm_all_details_found)
                time.sleep(regular_interval_time)
                pyautogui.click(yellow_next_page_found)
                time.sleep(between_page_time)

                pyautogui.hotkey('ctrl', '-')
                pyautogui.hotkey('ctrl', '-')
                pyautogui.hotkey('ctrl', '-')
                pyautogui.hotkey('ctrl', '-')
                pyautogui.hotkey('ctrl', '-')
                pyautogui.hotkey('ctrl', '-')
                time.sleep(regular_interval_time)
                pyautogui.press('pgup', presses=1)
                time.sleep(regular_interval_time)
                pyautogui.press('pgup', presses=1)
                time.sleep(regular_interval_time)
                pyautogui.press('pgup', presses=1)
                time.sleep(regular_interval_time)
                pyautogui.press('pgup', presses=1)
                time.sleep(regular_interval_time)

                check_button_found = pyautogui.locateCenterOnScreen(image='images/check.png', confidence=0.8)
                pyautogui.click(check_button_found)
                time.sleep(regular_interval_time)
                pyautogui.hotkey('ctrl', '+')
                pyautogui.hotkey('ctrl', '+')
                pyautogui.hotkey('ctrl', '+')
                pyautogui.hotkey('ctrl', '+')
                pyautogui.hotkey('ctrl', '+')
                pyautogui.hotkey('ctrl', '+')
                pyautogui.press('tab', presses=1)
                counter += 1
                continue

            elif next_page_found:
                time.sleep(regular_interval_time)
                confirm_all_details_found = pyautogui.locateCenterOnScreen(image='images/confirm_all_details_250.png',
                                                                           confidence=0.8)
                time.sleep(regular_interval_time)
                pyautogui.click(confirm_all_details_found)
                time.sleep(regular_interval_time)
                pyautogui.click(next_page_found)
                time.sleep(between_page_time)

                pyautogui.hotkey('ctrl', '-')
                pyautogui.hotkey('ctrl', '-')
                pyautogui.hotkey('ctrl', '-')
                pyautogui.hotkey('ctrl', '-')
                pyautogui.hotkey('ctrl', '-')
                pyautogui.hotkey('ctrl', '-')
                time.sleep(regular_interval_time)
                pyautogui.press('pgup', presses=1)
                time.sleep(regular_interval_time)
                pyautogui.press('pgup', presses=1)
                time.sleep(regular_interval_time)
                pyautogui.press('pgup', presses=1)
                time.sleep(regular_interval_time)
                pyautogui.press('pgup', presses=1)
                time.sleep(regular_interval_time)

                check_button_found = pyautogui.locateCenterOnScreen(image='images/check.png', confidence=0.8)
                pyautogui.click(check_button_found)
                time.sleep(regular_interval_time)

                pyautogui.hotkey('ctrl', '+')
                pyautogui.hotkey('ctrl', '+')
                pyautogui.hotkey('ctrl', '+')
                pyautogui.hotkey('ctrl', '+')
                pyautogui.hotkey('ctrl', '+')
                pyautogui.hotkey('ctrl', '+')
                pyautogui.press('tab', presses=1)
                counter += 1
                continue

            else:
                confirm_all_details_found = pyautogui.locateCenterOnScreen(image='images/confirm_all_details_250.png',
                                                                           confidence=0.8)
                if confirm_all_details_found:
                    pyautogui.click(confirm_all_details_found)
                    time.sleep(regular_interval_time)
                    pyautogui.press('tab', presses=1)
                    time.sleep(regular_interval_time)
                    pyautogui.press('space', presses=1)
                    time.sleep(between_page_time)

                    pyautogui.hotkey('ctrl', '-')
                    pyautogui.hotkey('ctrl', '-')
                    pyautogui.hotkey('ctrl', '-')
                    pyautogui.hotkey('ctrl', '-')
                    pyautogui.hotkey('ctrl', '-')
                    pyautogui.hotkey('ctrl', '-')

                else:
                    pyautogui.press('tab', presses=1)
                    time.sleep(regular_interval_time)
                    counter += 1

        WindowManager.make_message_box(
            title='Success', message='Automation complete.', icon='info')


def main():
    database = DatabaseManager()
    database.create_tables()

    root = tk.Tk()
    root.iconbitmap('icon.ico')
    root.withdraw()

    auto_pcr_registration = AutoPCRRegistration(master=root, geometry='300x300', title='Auto PCR Registration',
                                                previous_window=root)
    auto_pcr_registration.center_window(300, 300)

    root.mainloop()

    database.connection.close()


if __name__ == '__main__':
    main()
