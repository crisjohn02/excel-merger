import sqlalchemy as db
from sqlalchemy import create_engine, and_
from datetime import datetime, timedelta
import sqlalchemy.sql.default_comparator
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from sqlalchemy import desc
import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter.filedialog import asksaveasfilename
import pandas as pd
from ttkwidgets.autocomplete import AutocompleteEntry

engine = create_engine('sqlite:///timesheet.db')
connection = engine.connect()
Base = declarative_base()


class Time(Base):
    __tablename__ = 'timesheets'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True, index=True)
    name = db.Column(db.String, nullable=False)
    project = db.Column(db.Text, nullable=False)
    task = db.Column(db.Text, nullable=False)
    link = db.Column(db.Text, nullable=True)
    start = db.Column(db.DateTime, nullable=True)
    end = db.Column(db.DateTime, nullable=True)


# Create database if not present and create table
Base.metadata.create_all(bind=engine)

Session = sessionmaker()
Session.configure(bind=engine)
session = Session()


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.geometry("260x200")
        self.title("Timesheet")

        # configure the grid
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=3)

        # inputs
        self.name = None
        self.project = None
        self.description = None
        self.link = None
        self.cmd_load_last = None
        self.cmd_start = None
        self.cmd_stop = None
        self.cmd_export = None
        self.lbl_time = None

        # db rows
        self.row = None
        self.projects = []

        self.init_data()
        self.create_widgets()

    def create_widgets(self):
        ttk.Label(self, text="Name").grid(column=0, row=1, sticky=tk.W, padx=5, pady=5)
        self.name = ttk.Entry(self)
        self.name.grid(column=1, row=1, sticky=tk.EW, padx=5, columnspan=2)

        ttk.Label(self, text="Project").grid(column=0, row=2, sticky=tk.W, padx=5, pady=5)
        self.project = AutocompleteEntry(self, completevalues=self.projects)
        self.project.grid(column=1, row=2, sticky=tk.EW, padx=5, columnspan=2)

        ttk.Label(self, text="Description").grid(column=0, row=4, sticky=tk.W, padx=5, pady=5)
        self.description = Text(self, height=2, width=20)
        self.description.grid(column=1, row=4, sticky=tk.EW, padx=5, columnspan=2)

        ttk.Label(self, text="Drive/Github").grid(column=0, row=6, sticky=tk.W, padx=5, pady=5)
        self.link = Text(self, height=2, width=20)
        self.link.grid(column=1, row=6, sticky=tk.EW, padx=5, columnspan=2)

        self.cmd_export = ttk.Button(self, text="EXPORT", command=self.export)
        self.cmd_export.grid(column=0, row=8, pady=5, padx=5)

        self.cmd_start = tk.Button(self, text="START", command=self.start, fg="black", bg="#00FF00")
        self.cmd_start.grid(column=1, row=8, sticky=tk.EW, padx=5, pady=5)

        self.cmd_stop = tk.Button(self, text="STOP", command=self.stop, fg="white", bg="#ff0000")
        self.cmd_stop.grid(column=2, row=8, sticky=tk.EW, padx=5, pady=5)

        ttk.Label(self, text="Datetime").grid(column=0, row=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        self.lbl_time = ttk.Label(self, text=datetime.now().strftime('%a %b %d, %Y %I:%M:%S %p'), background='purple',
                                  foreground='white')
        self.lbl_time.grid(column=1, row=0, columnspan=2)
        self.time()
        self.init_state()

    def time(self):
        string = datetime.now().strftime('%a %b %d, %Y %I:%M:%S %p')
        self.lbl_time.config(text=string)
        self.lbl_time.after(1000, self.time)

    def init_data(self):
        p = session.query(Time)
        for _row in p.all():
            self.projects.append(_row.project)

        self.row = session.query(Time).order_by(desc(Time.id)).first()

    def init_state(self):
        if self.row:
            self.name.delete(0, END)
            self.name.insert(0, self.row.name)

            self.project.delete(0, END)
            self.project.insert(0, self.row.project)

            self.description.delete('1.0', END)
            self.description.insert(END, self.row.task)

            self.link.delete('1.0', END)
            self.link.insert(END, self.row.link)

            if self.row.end is None:
                self.cmd_stop['state'] = NORMAL
                self.cmd_start['state'] = DISABLED
            else:
                self.cmd_stop['state'] = DISABLED
                self.cmd_start['state'] = NORMAL
        else:
            self.cmd_stop['state'] = DISABLED

    def start(self):
        r = Time(
            name=self.name.get(),
            project=self.project.get(),
            task=self.description.get("1.0", "end-1c"),
            link=self.link.get("1.0", "end-1c"),
            start=datetime.now()
        )
        session.add(r)
        session.commit()
        self.row = r
        self.cmd_stop['state'] = NORMAL
        self.cmd_start['state'] = DISABLED

    def stop(self):
        self.row.end = datetime.now()
        session.commit()
        self.init_state()

    def export(self):
        # dt = datetime.now().strptime('%d/%b/%Y')
        _start = datetime.now() - timedelta(days=datetime.now().weekday())
        _end = _start + timedelta(days=6)
        start = _start.strftime('%Y-%m-%d')
        end = _end.strftime('%Y-%m-%d')

        query = session.query(Time).filter(
            and_(Time.start.between(start, end), Time.end != None)
        ).order_by(Time.start)
        d_row = {
            'Date': [],
            'Name': [],
            'Project': [],
            'Task/Description': [],
            'Drive/Github': [],
            'Start': [],
            'End': [],
            'Total Hours': []
        }
        for _row in query.all():
            # __start = datetime.strptime(_row.start, '%d/%m/%Y %H:%M:%S.%f')
            # __end = datetime.strptime(_row.end, '%d/%m/%Y %H:%M:%S.%f')
            diff = _row.end - _row.start
            d_row['Date'].append(_row.start.strftime('%m/%d/%Y'))
            d_row['Name'].append(_row.name)
            d_row['Project'].append(_row.project)
            d_row['Task/Description'].append(_row.task)
            d_row['Drive/Github'].append(_row.link)
            d_row['Start'].append(_row.start)
            d_row['End'].append(_row.end)
            d_row['Total Hours'].append(round(diff.total_seconds() / 3600, 2))
        df = pd.DataFrame(d_row)
        filename = asksaveasfilename(filetypes=(("Excel", ('*.xls', '*.xlsx')), ("All files", '*.*')),
                                     defaultextension=".xlsx")
        df.to_excel(filename, index=False)


if __name__ == "__main__":
    app = App()
    app.mainloop()
