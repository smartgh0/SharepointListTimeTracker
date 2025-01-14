import kivy
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.popup import Popup
from kivy.core.window import Window
import csv
import time
import os
import socket
import requests
from msal import PublicClientApplication
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from datetime import datetime, timedelta
import json
import threading

class WorkTimeTracker(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        Window.size = (250, 500)
        Window.resize = False

        # Settings button to open SharePoint settings
        self.settings_btn = Button(text="Settings", on_press=self.open_settings)
        self.add_widget(self.settings_btn)

        # User ID input field
        self.user_id_input = TextInput(hint_text="Enter your ID", multiline=False)
        self.add_widget(self.user_id_input)

        # Department input field
        self.department_input = TextInput(hint_text="Enter your Department", multiline=False)
        self.add_widget(self.department_input)
        
        # Task input field
        self.task_input = TextInput(hint_text="Enter task description", multiline=False)
        self.add_widget(self.task_input)

        # Clock in and out buttons
        self.clock_in_btn = Button(text="Clock In", on_press=self.clock_in)
        self.add_widget(self.clock_in_btn)

        self.clock_out_btn = Button(text="Clock Out", on_press=self.clock_out, disabled=True)
        self.add_widget(self.clock_out_btn)

        self.generate_pdf_btn = Button(text="Generate Weekly Summary (PDF)", on_press=self.generate_pdf)
        self.add_widget(self.generate_pdf_btn)

        self.start_time = None
        self.end_time = None
        self.task_description = None

        # Load SharePoint configuration from file
        self.load_settings()

    def load_settings(self):
        try:
            with open('settings.json', 'r') as file:
                settings = json.load(file)
                self.client_id = settings.get("client_id", "")
                self.authority = settings.get("authority", "")
                self.scopes = settings.get("scopes", ["https://graph.microsoft.com/Sites.ReadWrite.All"])
                self.site_url = settings.get("site_url", "")
                self.list_name = settings.get("list_name", "")
                self.token = None  # Initialize token here, will be acquired during clock out
                self.site_id = None
                self.list_id = None
                self.token_expiry = None  # Token expiry time
        except FileNotFoundError:
            print("Settings file not found. Please enter your SharePoint settings.")
            self.client_id = ""
            self.authority = ""
            self.scopes = ["https://graph.microsoft.com/Sites.ReadWrite.All"]
            self.site_url = ""
            self.list_name = ""
            self.token = None
            self.site_id = None
            self.list_id = None
            self.token_expiry = None

    def open_settings(self, instance):
        self.settings_popup = Popup(
            title="SharePoint Settings",
            content=self.create_settings_form(),
            size_hint=(0.8, 0.6)
        )
        self.settings_popup.open()

    def create_settings_form(self):
        layout = BoxLayout(orientation='vertical', padding=10)
        
        # Create input fields for settings
        self.client_id_input = TextInput(hint_text="Enter your Client ID", text=self.client_id, multiline=False)
        self.authority_input = TextInput(hint_text="Enter your Authority URL", text=self.authority, multiline=False)
        self.site_url_input = TextInput(hint_text="Enter SharePoint Site URL", text=self.site_url, multiline=False)
        self.list_name_input = TextInput(hint_text="Enter List Name", text=self.list_name, multiline=False)

        save_button = Button(text="Save Settings", on_press=self.save_settings)
        cancel_button = Button(text="Cancel", on_press=self.close_settings_popup)

        # Add widgets to layout
        layout.add_widget(self.client_id_input)
        layout.add_widget(self.authority_input)
        layout.add_widget(self.site_url_input)
        layout.add_widget(self.list_name_input)
        layout.add_widget(save_button)
        layout.add_widget(cancel_button)

        return layout

    def save_settings(self, instance):
        self.client_id = self.client_id_input.text.strip()
        self.authority = self.authority_input.text.strip()
        self.site_url = self.site_url_input.text.strip()
        self.list_name = self.list_name_input.text.strip()

        settings = {
            "client_id": self.client_id,
            "authority": self.authority,
            "scopes": ["https://graph.microsoft.com/Sites.ReadWrite.All"],
            "site_url": self.site_url,
            "list_name": self.list_name
        }

        with open('settings.json', 'w') as file:
            json.dump(settings, file, indent=4)

        self.site_id = self.get_site_id()
        self.list_id = self.get_list_id()

        self.settings_popup.dismiss()
        self.show_popup("Settings Saved", "Your SharePoint settings have been saved successfully.")

    def close_settings_popup(self, instance):
        self.settings_popup.dismiss()

    def get_site_id(self):
        if not self.site_url or not self.token:
            return None
        
        site_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_url.replace('https://', '').replace('/', ':')}"
        headers = {"Authorization": f"Bearer {self.token}", "Accept": "application/json"}
        response = requests.get(site_url, headers=headers)

        if response.status_code == 200:
            return response.json().get('id')
        else:
            return None

    def get_list_id(self):
        if not self.site_id or not self.token:
            return None
        
        list_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/lists"
        headers = {"Authorization": f"Bearer {self.token}", "Accept": "application/json"}
        response = requests.get(list_url, headers=headers)

        if response.status_code == 200:
            for l in response.json().get('value', []):
                if l['name'] == self.list_name:
                    return l['id']
        return None

    def acquire_token(self):
        """
        Handles token acquisition and automatic renewal when expired.
        """
        if not self.client_id or not self.authority:
            self.show_popup("Error", "Client ID or Authority is not set.")
            return False
        
        app = PublicClientApplication(client_id=self.client_id, authority=self.authority)
        accounts = app.get_accounts()
        
        if accounts:
            token_response = app.acquire_token_silent(scopes=self.scopes, account=accounts[0])
            if token_response and "access_token" in token_response:
                self.token = token_response["access_token"]
                self.token_expiry = datetime.now() + timedelta(seconds=int(token_response["expires_in"]))
                return True

        # Interactive acquisition if no valid token is available
        token_response = app.acquire_token_interactive(scopes=self.scopes)
        if "access_token" in token_response:
            self.token = token_response["access_token"]
            self.token_expiry = datetime.now() + timedelta(seconds=int(token_response["expires_in"]))
            return True
        
        self.show_popup("Error", "Failed to acquire token.")
        return False

    def clock_in(self, instance):
        if not self.task_input.text.strip():
            self.show_popup("Error", "Please enter a task description before clocking in.")
            return

        self.start_time = time.time()
        self.task_description = self.task_input.text.strip()
        self.clock_in_btn.disabled = True
        self.clock_out_btn.disabled = False

        self.show_popup("Clock In", f"Clocked in at: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(self.start_time))}")

    def clock_out(self, instance):
        if not self.token or (self.token_expiry and datetime.now() >= self.token_expiry):
            # Acquire or renew token on clock-out
            if not self.acquire_token():
                return

        self.end_time = time.time()
        duration = round((self.end_time - self.start_time) / 3600, 2)  # Duration in hours
        current_date = time.strftime("%Y-%m-%d", time.localtime(self.start_time))

        # Log time to CSV
        with open('time_log.csv', mode='a', newline='') as file:
            writer = csv.writer(file)
            writer.writerow([current_date, self.task_description, self.start_time, self.end_time, duration])

        # Start the background thread for sending details to SharePoint
        threading.Thread(target=self.send_to_sharepoint, args=(current_date, self.task_description, self.start_time, self.end_time, duration)).start()

        self.show_popup("Clock Out", f"Clocked out at: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(self.end_time))}")
        self.task_input.text = ""
        self.clock_in_btn.disabled = False
        self.clock_out_btn.disabled = True

    def send_to_sharepoint(self, date, task, start, end, duration):
        add_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/lists/{self.list_id}/items"
        data = {
            "fields": {
                "Title": socket.gethostname(),
                "TaskDescription": task,
                "Clock_x002d_inTime": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(float(start))),
                "Clock_x002d_out": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(float(end))),
                "Department":  str(self.department_input.text.strip()),
                "WorkerID": str(self.user_id_input.text.strip()),
                "Date": str(datetime.today().date()),
            }
        }
        headers = {"Authorization": f"Bearer {self.token}", "Content-Type": "application/json"}
        response = requests.post(add_url, headers=headers, json=data)

        if response.status_code == 201:
            print("Task successfully added to SharePoint.")
        else:
            print(f"Failed to add task to SharePoint: {response.status_code} - {response.text}")

    def generate_pdf(self, instance):
        save_path = os.path.join(os.getcwd(), "weekly_summary.pdf")
        document = SimpleDocTemplate(save_path, pagesize=letter)
        elements = []

        data = [["Date", "Task Description", "Clock-in", "Clock-out", "Hours"]]
        total_hours = 0

        try:
            with open('time_log.csv', mode='r') as file:
                reader = csv.reader(file)
                for row in reader:
                    date, task, start, end, duration = row
                    clock_in_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(float(start)))
                    clock_out_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(float(end)))
                    total_hours += float(duration)
                    data.append([date, task, clock_in_time, clock_out_time, duration])

        except FileNotFoundError:
            self.show_popup("Error", "No time log found to generate summary.")

        table = Table(data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ]))

        elements.append(table)
        elements.append(Label(text=f"Total Hours Worked: {total_hours}"))

        document.build(elements)

        self.show_popup("PDF Generated", f"Weekly summary saved as: {save_path}")

    def show_popup(self, title, message):
        content = BoxLayout(orientation='vertical')
        popup_label = Label(text=message)
        content.add_widget(popup_label)
        close_btn = Button(text="Close", on_press=self.close_popup)
        content.add_widget(close_btn)
        
        self.popup = Popup(title=title, content=content, size_hint=(0.6, 0.4))
        self.popup.open()

    def close_popup(self, instance):
        self.popup.dismiss()

class WorkTimeApp(App):
    def build(self):
        return WorkTimeTracker()

if __name__ == "__main__":
    WorkTimeApp().run()
