#!/usr/bin/env python3
"""
Schedule HTML to ICS Converter
Converts Berufsschule HTML schedule to ICS calendar format
"""

import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from bs4 import BeautifulSoup
import datetime
from icalendar import Calendar, Event
import pytz
from pathlib import Path

class ScheduleConverter:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Berufsschule Schedule to ICS Converter")
        self.root.geometry("800x700")
        
        self.schedule_data = {}
        self.html_file = None
        self.week_vars = {}
        self.reminder_vars = {}
        
        # Time slots for Monday-Thursday
        self.weekday_times = {
            1: ("08:00", "09:30"),  # 1-2 UE
            3: ("09:45", "11:15"),  # 3-4 UE
            5: ("11:25", "12:10"),  # 5 UE
            6: ("12:55", "14:25"),  # 6-7 UE
            8: ("14:35", "15:20"),  # 8 UE
            9: ("15:30", "17:00"),  # 9-10 UE
        }
        
        # Time slots for Friday
        self.friday_times = {
            1: ("08:00", "09:30"),  # 1-2 UE
            3: ("09:45", "11:15"),  # 3-4 UE
            5: ("11:30", "13:00"),  # 5-6 UE
            7: ("13:00", "14:30"),  # 7-8 UE (ILIAS-Lernauftrag - skip this)
        }
        
        # Subjects to ignore
        self.ignore_subjects = {"#NV", "Frei", "Betrieb", "Feiertag", "", " "}
        
        self.setup_ui()
    
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # File selection
        file_frame = ttk.Frame(main_frame)
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(file_frame, text="HTML Schedule File:").grid(row=0, column=0, sticky=tk.W)
        self.file_label = ttk.Label(file_frame, text="No file selected", foreground="red")
        self.file_label.grid(row=1, column=0, sticky=tk.W)
        
        ttk.Button(file_frame, text="Browse", command=self.browse_file).grid(row=0, column=1, padx=(10, 0))
        ttk.Button(file_frame, text="Load Schedule", command=self.load_schedule).grid(row=1, column=1, padx=(10, 0))
        
        # Warning label
        warning_frame = ttk.Frame(main_frame)
        warning_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        warning_label = ttk.Label(warning_frame, 
                                 text="⚠️ WARNING: Subjects and times may change. Please verify with official schedule!",
                                 foreground="orange", font=("TkDefaultFont", 9, "bold"))
        warning_label.grid(row=0, column=0, sticky=tk.W)
        
        # Week selection frame
        self.week_frame = ttk.LabelFrame(main_frame, text="Select Weeks to Include", padding="5")
        self.week_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Scrollable frame for weeks
        canvas = tk.Canvas(self.week_frame, height=200)
        scrollbar = ttk.Scrollbar(self.week_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Reminder settings frame
        reminder_frame = ttk.LabelFrame(main_frame, text="Reminder Settings", padding="5")
        reminder_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Reminder checkboxes
        self.reminder_vars['before_first'] = tk.BooleanVar(value=True)
        self.reminder_vars['end_previous'] = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(reminder_frame, 
                       text="15 minutes before first lesson of the day",
                       variable=self.reminder_vars['before_first']).grid(row=0, column=0, sticky=tk.W, pady=2)
        
        ttk.Checkbutton(reminder_frame, 
                       text="At the end of the lesson before (for next lesson preparation)",
                       variable=self.reminder_vars['end_previous']).grid(row=1, column=0, sticky=tk.W, pady=2)
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        ttk.Button(button_frame, text="Select All", command=self.select_all_weeks).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(button_frame, text="Deselect All", command=self.deselect_all_weeks).grid(row=0, column=1, padx=5)
        ttk.Button(button_frame, text="Generate ICS", command=self.generate_ics).grid(row=0, column=2, padx=(20, 0))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        self.week_frame.columnconfigure(0, weight=1)
        self.week_frame.rowconfigure(0, weight=1)
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select HTML Schedule File",
            filetypes=[("HTML files", "*.html *.htm"), ("All files", "*.*")]
        )
        if file_path:
            self.html_file = file_path
            self.file_label.config(text=Path(file_path).name, foreground="green")
    
    def load_schedule(self):
        if not self.html_file:
            messagebox.showerror("Error", "Please select an HTML file first!")
            return
        
        try:
            with open(self.html_file, 'r', encoding='utf-8') as file:
                content = file.read()
            
            soup = BeautifulSoup(content, 'html.parser')
            self.parse_schedule(soup)
            self.display_weeks()
            messagebox.showinfo("Success", "Schedule loaded successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load schedule: {str(e)}")
    
    def parse_schedule(self, soup):
        table = soup.find('table')
        if not table:
            raise Exception("No table found in HTML file")
        
        rows = table.find_all('tr')
        
        # Find the data rows (skip header rows)
        data_rows = []
        for i, row in enumerate(rows):
            cells = row.find_all(['td', 'th'])
            if len(cells) >= 12 and cells[0].get_text().strip():
                # Check if first cell contains a date
                date_text = cells[0].get_text().strip()
                if re.match(r'\d{1,2}\.\s*\w+', date_text):
                    data_rows.append(row)
        
        # Parse each data row
        for row in data_rows:
            cells = row.find_all(['td', 'th'])
            if len(cells) < 12:
                continue
                
            date_text = cells[0].get_text().strip()
            day_text = cells[1].get_text().strip()
            
            # Parse date
            date_match = re.match(r'(\d{1,2})\.\s*(\w+)', date_text)
            if not date_match:
                continue
                
            day_num = int(date_match.group(1))
            month_name = date_match.group(2).lower()
            
            # Convert German month names to numbers
            month_map = {
                'jan': 1, 'feb': 2, 'mär': 3, 'apr': 4, 'mai': 5, 'jun': 6,
                'jul': 7, 'aug': 8, 'sep': 9, 'okt': 10, 'nov': 11, 'dez': 12
            }
            
            month_num = month_map.get(month_name[:3], 1)
            
            # Determine year (assume current academic year 2025/2026)
            year = 2025 if month_num >= 8 else 2026
            
            try:
                date_obj = datetime.date(year, month_num, day_num)
            except ValueError:
                continue
            
            # Parse subjects for each time slot (columns 2-11 represent periods 1-10)
            day_schedule = {}
            for i in range(2, 12):  # Columns 2-11
                if i < len(cells):
                    subject_text = cells[i].get_text().strip()
                    # Clean up subject text
                    subject_lines = [line.strip() for line in subject_text.split('\n') if line.strip()]
                    if subject_lines and subject_lines[0] not in self.ignore_subjects:
                        period = i - 1  # Convert to period number (1-10)
                        day_schedule[period] = subject_lines[0]
            
            if day_schedule:  # Only add days that have subjects
                self.schedule_data[date_obj] = {
                    'day': day_text,
                    'subjects': day_schedule
                }
    
    def display_weeks(self):
        # Clear existing widgets
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        self.week_vars.clear()
        
        # Group dates by week
        weeks = {}
        for date_obj in sorted(self.schedule_data.keys()):
            # Get Monday of the week
            monday = date_obj - datetime.timedelta(days=date_obj.weekday())
            if monday not in weeks:
                weeks[monday] = []
            weeks[monday].append(date_obj)
        
        # Create checkboxes for each week
        row = 0
        for monday in sorted(weeks.keys()):
            week_dates = sorted(weeks[monday])
            
            # Create week label
            week_start = week_dates[0]
            week_end = week_dates[-1]
            week_label = f"Week {week_start.strftime('%d.%m')} - {week_end.strftime('%d.%m.%Y')}"
            
            # Get subjects for this week
            week_subjects = set()
            for date_obj in week_dates:
                week_subjects.update(self.schedule_data[date_obj]['subjects'].values())
            
            subjects_text = ", ".join(sorted(week_subjects)) if week_subjects else "No subjects"
            
            # Create checkbox
            var = tk.BooleanVar(value=True)
            self.week_vars[monday] = var
            
            checkbox = ttk.Checkbutton(self.scrollable_frame, variable=var, text=week_label)
            checkbox.grid(row=row, column=0, sticky=tk.W, pady=2)
            
            # Create subjects label
            subjects_label = ttk.Label(self.scrollable_frame, text=f"Subjects: {subjects_text}", 
                                     foreground="gray", font=("TkDefaultFont", 8))
            subjects_label.grid(row=row+1, column=0, sticky=tk.W, padx=(20, 0))
            
            row += 2
    
    def select_all_weeks(self):
        for var in self.week_vars.values():
            var.set(True)
    
    def deselect_all_weeks(self):
        for var in self.week_vars.values():
            var.set(False)
    
    def generate_ics(self):
        if not self.schedule_data:
            messagebox.showerror("Error", "No schedule data loaded!")
            return
        
        # Get selected weeks
        selected_weeks = [monday for monday, var in self.week_vars.items() if var.get()]
        
        if not selected_weeks:
            messagebox.showerror("Error", "No weeks selected!")
            return
        
        # Create calendar
        cal = Calendar()
        cal.add('prodid', '-//Berufsschule Schedule Converter//mxm.dk//')
        cal.add('version', '2.0')
        cal.add('calscale', 'GREGORIAN')
        
        # Set timezone
        timezone = pytz.timezone('Europe/Berlin')
        
        # Add events for selected weeks
        for monday in selected_weeks:
            week_dates = []
            for date_obj in self.schedule_data.keys():
                if date_obj >= monday and date_obj < monday + datetime.timedelta(days=7):
                    week_dates.append(date_obj)
            
            for date_obj in week_dates:
                day_data = self.schedule_data[date_obj]
                subjects = day_data['subjects']
                
                # Determine which time schedule to use
                is_friday = date_obj.weekday() == 4
                time_schedule = self.friday_times if is_friday else self.weekday_times
                
                # Sort periods to find first lesson of the day
                sorted_periods = sorted(subjects.keys())
                first_period = sorted_periods[0] if sorted_periods else None
                
                # Find previous lesson for each period (for end-of-previous-lesson reminders)
                previous_lessons = {}
                for i, period in enumerate(sorted_periods):
                    if i > 0:
                        previous_lessons[period] = sorted_periods[i-1]
                
                for period, subject in subjects.items():
                    if subject in self.ignore_subjects:
                        continue
                    
                    # Get time for this period
                    if period in time_schedule:
                        start_time_str, end_time_str = time_schedule[period]
                        
                        # Parse times
                        start_time = datetime.datetime.strptime(start_time_str, "%H:%M").time()
                        end_time = datetime.datetime.strptime(end_time_str, "%H:%M").time()
                        
                        # Create datetime objects
                        start_datetime = timezone.localize(
                            datetime.datetime.combine(date_obj, start_time)
                        )
                        end_datetime = timezone.localize(
                            datetime.datetime.combine(date_obj, end_time)
                        )
                        
                        # Create event
                        event = Event()
                        event.add('summary', subject)
                        event.add('dtstart', start_datetime)
                        event.add('dtend', end_datetime)
                        event.add('description', f'Berufsschule - Period {period}')
                        
                        # Add reminders
                        alarms = []
                        
                        # 15 minutes before first lesson of the day
                        if self.reminder_vars['before_first'].get() and period == first_period:
                            from icalendar import Alarm
                            alarm = Alarm()
                            alarm.add('action', 'DISPLAY')
                            alarm.add('description', f'First lesson starting soon: {subject}')
                            alarm.add('trigger', datetime.timedelta(minutes=-15))
                            alarms.append(alarm)
                        
                        # At the end of the previous lesson
                        if self.reminder_vars['end_previous'].get() and period in previous_lessons:
                            previous_period = previous_lessons[period]
                            if previous_period in time_schedule:
                                prev_end_time_str = time_schedule[previous_period][1]
                                prev_end_time = datetime.datetime.strptime(prev_end_time_str, "%H:%M").time()
                                prev_end_datetime = timezone.localize(
                                    datetime.datetime.combine(date_obj, prev_end_time)
                                )
                                
                                # Calculate minutes from previous lesson end to current lesson start
                                time_diff = start_datetime - prev_end_datetime
                                minutes_diff = int(time_diff.total_seconds() / 60)
                                
                                from icalendar import Alarm
                                alarm = Alarm()
                                alarm.add('action', 'DISPLAY')
                                alarm.add('description', f'Next lesson preparation: {subject}')
                                alarm.add('trigger', datetime.timedelta(minutes=-minutes_diff))
                                alarms.append(alarm)
                        
                        # Add alarms to event
                        for alarm in alarms:
                            event.add_component(alarm)
                        
                        cal.add_component(event)
        
        # Save ICS file
        file_path = filedialog.asksaveasfilename(
            title="Save ICS Calendar File",
            defaultextension=".ics",
            filetypes=[("ICS files", "*.ics"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'wb') as f:
                    f.write(cal.to_ical())
                messagebox.showinfo("Success", f"Calendar saved successfully to {Path(file_path).name}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save calendar: {str(e)}")
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    # Check for required packages
    try:
        import tkinter as tk
        from tkinter import ttk, filedialog, messagebox
        from bs4 import BeautifulSoup
        from icalendar import Calendar, Event
        import pytz
    except ImportError as e:
        print(f"Missing required package: {e}")
        print("Please install required packages:")
        print("pip install beautifulsoup4 icalendar pytz")
        exit(1)
    
    app = ScheduleConverter()
    app.run()