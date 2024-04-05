from tkinter import *
from tkinter import ttk
from tkcalendar import *
import json
import win32com.client
import tkinter.messagebox as messagebox


def send_email():

    AR = entry_AR.get()
    if not AR:        
        messagebox.showerror("Error", "AR is required")
        return

    project_manager = entry_project_manager.get()
    network_number = entry_network_number.get()
    TS = combo_TS.get()
    if not TS:
        messagebox.showerror("Error", "Transformer Station is required")
        return
    
    team_lead = label_team_lead_value.cget('text')
    DS = entry_DS.get()
    date = entry_date.get()
    gate = combo_gate.get()

    links_value = links.get('1.0', 'end-1c')
    comments_value = comments.get('1.0', 'end-1c')

    resource_requested = []
    if outage_planner.get():
        resource_requested.append('Outage Planner')
    if outage_coordinator.get():
        resource_requested.append('Outage Coordinator')
    if contractor_training.get():
        resource_requested.append('Contractor Training')
    if outage_planner_for_new_DG.get():
        resource_requested.append('Outage Planner for New DG')

    if not resource_requested:
        messagebox.showerror("Error", "At least one resource is required")
        return
    
    email_string = ""
    recipient = ""
    email_HTMLbody = "<p>Hi {recipient},</p><p>There is a request for OP:</p><p><b>AR:</b> {AR}</p><p><b>Project Manager:</b> {project_manager}</p><p><b>Network Number:</b> {network_number}</p><p><b>Transformer Station:</b> {TS}</p><p><b>Distribution Station:</b> {DS}</p><p><b>Kick-Off Date:</b> {date}</p><p><b>Stage Gate:</b> {gate}</p><p><b>Links to Supporting Documentation:</b> {links_value}</p><p><b>Additional Comments:</b> {comments_value}</p>"

    with open('TS.json', 'r') as f:
        data = json.load(f)

    planner = [item['Team Lead'] for item in data if item['Transformer Station'] == TS][0]

    planner_email = [item['Email'] for item in data if item['Transformer Station'] == TS][0]

    if 'Outage Planner' in resource_requested and not 'Outage Coordinator' in resource_requested and not 'Contractor Training' in resource_requested:
        email_string = planner_email
        recipient = planner.split()[0]
    elif 'Outage Coordinator' in resource_requested and not 'Outage Planner' in resource_requested and not 'Contractor Training' in resource_requested:
        email_string = "Darrel.Davies@HydroOne.com;"
        recipient = "Darrel"
    elif 'Contractor Training' in resource_requested and not 'Outage Coordinator' in resource_requested:
        email_string = planner_email + "; Shane.Suppa@HydroOne.com"
        recipient = "Shane & " + planner.split()[0]
    elif 'Outage Planner for New DG' in resource_requested:
        email_string = "Patrick.OGrady@HydroOne.com"
        recipient = "Patrick"
    elif 'Outage Planner' in resource_requested and 'Outage Coordinator' in resource_requested and not 'Contractor Training' in resource_requested:
        email_string = planner_email + "; Darrel.Davies@HydroOne.com"
        recipient = planner.split()[0] + " & Darrel"
    elif 'Outage Coordinator' in resource_requested and 'Contractor Training' in resource_requested:
        email_string = planner_email + "; Darrel.Davies@HydroOne.com; Shane.Suppa@HydroOne.com"
        recipient = planner.split()[0] + ", Darrel & Shane"

    try:
        outlook_app = win32com.client.Dispatch('outlook.application')
        mail = outlook_app.CreateItem(0)

        mail.Subject = 'Work Execution & UWPC Services Request Form'
        mail.To = email_string
        mail.HTMLBody = email_HTMLbody.format(recipient=recipient, AR=AR, project_manager=project_manager, network_number=network_number, TS=TS, team_lead=team_lead, DS=DS, date=date, gate=gate, links_value=links_value, comments_value=comments_value)

        mail.display()

    except Exception as e:
        print(e)
        return

def pick_date(e):
    global cal, date_window

    date_window = Toplevel()
    date_window.grab_set()
    date_window.title('Choose Kick-Off Date')
    date_window.geometry('250x220+590+370')
    date_window.resizable(False, False)
    cal = Calendar(date_window, selectmode="day", date_pattern="mm/dd/y")
    cal.place(x=0, y=0)

    submit_btn = Button(date_window, text="Submit", command=grab_date)
    submit_btn.place(x=100, y=180)

def grab_date():
    entry_date.delete(0, END)
    entry_date.insert(0, cal.get_date())
    date_window.destroy()

def TS_selected(e):
    
    TS = combo_TS.get()
    for item in data:
        if item['Transformer Station'] == TS:
            region = item['Region']
            team_lead = item['Team Lead']

    label_region_value.config(text=region)
    label_team_lead_value.config(text=team_lead)

def check_boxes_status():
    root.after(800, lambda: disactivate())

def disactivate():
    print(outage_planner.get(), outage_coordinator.get(), contractor_training.get(), outage_planner_for_new_DG.get())
    if outage_planner.get() or outage_coordinator.get() or contractor_training.get():
        check_four.config(state=DISABLED)
        print('one')
    elif outage_planner_for_new_DG.get():
        check_one.config(state=DISABLED)
        check_two.config(state=DISABLED)
        check_three.config(state=DISABLED)
        print('two')
    elif not outage_planner.get() and not outage_coordinator.get() and not contractor_training.get() and not outage_planner_for_new_DG.get():
        check_one.config(state=NORMAL)
        check_two.config(state=NORMAL)
        check_three.config(state=NORMAL)
        check_four.config(state=NORMAL)
        print('three')

root = Tk()

root.title("Work Execution & UWPC Services Request Form")
root.resizable(False, False)

###########

labelframe_project_details = ttk.Labelframe(root, text='Project Details')
labelframe_project_details.grid(row=0, column=0, padx=10, pady=10, sticky=NW)

###########

label_AR = ttk.Label(labelframe_project_details, text='AR')
label_AR.grid(row=0, column=0, padx=5, pady=5, sticky=W)

AR = StringVar()
entry_AR = ttk.Entry(labelframe_project_details, textvariable=AR, width=15)
entry_AR.grid(row=0, column=1, padx=5, pady=5, sticky=W)

###########

label_project_manager = ttk.Label(labelframe_project_details, text='Project Manager')
label_project_manager.grid(row=1, column=0, padx=5, pady=5, sticky=W)

project_manager = StringVar()
entry_project_manager = ttk.Entry(labelframe_project_details, textvariable=project_manager, width=15)
entry_project_manager.grid(row=1, column=1, padx=5, pady=5, sticky=W)

###########

label_network_number = ttk.Label(labelframe_project_details, text='Network Number')
label_network_number.grid(row=2, column=0, padx=5, pady=5, sticky=W)

network_number = StringVar()
entry_network_number = ttk.Entry(labelframe_project_details, textvariable=network_number, width=15)
entry_network_number.grid(row=2, column=1, padx=5, pady=5, sticky=W)

###########

label_TS = ttk.Label(labelframe_project_details, text='Transformer Station')
label_TS.grid(row=3, column=0, padx=5, pady=5, sticky=W)

with open('TS.json', 'r') as f:
    data = json.load(f)

list_TS = [item['Transformer Station'] for item in data]

TS = StringVar()

combo_TS = ttk.Combobox(labelframe_project_details, textvariable=TS, values=list_TS, width=15)
combo_TS.grid(row=3, column=1, padx=5, pady=5, sticky=W)

combo_TS.bind("<<ComboboxSelected>>", TS_selected)

###########

label_region = ttk.Label(labelframe_project_details, text='Region')
label_region.grid(row=4, column=0, padx=5, pady=5, sticky=W)

label_region_value = ttk.Label(labelframe_project_details, text='',foreground='blue')
label_region_value.grid(row=4, column=1, padx=5, pady=5, sticky=W)

###########

label_team_lead = ttk.Label(labelframe_project_details, text='Team Lead')
label_team_lead.grid(row=5, column=0, padx=5, pady=5, sticky=W)

label_team_lead_value = ttk.Label(labelframe_project_details, text='', foreground='blue')
label_team_lead_value.grid(row=5, column=1, padx=5, pady=5, sticky=W)

###########

label_DS = ttk.Label(labelframe_project_details, text='Distribution Station')
label_DS.grid(row=6, column=0, padx=5, pady=5, sticky=W)

DS = StringVar()
entry_DS = ttk.Entry(labelframe_project_details, textvariable=DS, width=15)
entry_DS.grid(row=6, column=1, padx=5, pady=5, sticky=W)

###########

label_date = ttk.Label(labelframe_project_details, text='Kick-Off Date')
label_date.grid(row=7, column=0, padx=5, pady=5, sticky=W)

date = StringVar()
entry_date = ttk.Entry(labelframe_project_details, textvariable=date, width=15)
entry_date.grid(row=7, column=1, padx=5, pady=5, sticky=W)
entry_date.bind("<1>", pick_date)

###########

label_gate = ttk.Label(labelframe_project_details, text='Stage Gate')
label_gate.grid(row=8, column=0, padx=5, pady=5, sticky=W)

gate = StringVar()
combo_gate = ttk.Combobox(labelframe_project_details, textvariable=gate, values=['INIT', 'BEST', 'DETL', 'EMPP'], width=15)
combo_gate.grid(row=8, column=1, padx=5, pady=5, sticky=W)

###########

labelframe_resource = ttk.Labelframe(root, text='Resource Requested')
labelframe_resource.grid(row=0, column=1, pady=10, sticky=NW)

# Add checkboxes inside the LabelFrame

outage_planner = IntVar()
outage_coordinator = IntVar()
contractor_training = IntVar()

check_one = ttk.Checkbutton(labelframe_resource, text="Outage Planner", variable=outage_planner)
check_one.grid(row=0, column=0, padx=5, pady=5, sticky=W)

check_two = ttk.Checkbutton(labelframe_resource, text="Outage Coodinator", variable=outage_coordinator)
check_two.grid(row=1, column=0, padx=5, pady=5, sticky=W)

check_three = ttk.Checkbutton(labelframe_resource, text="Contractor Training", variable=contractor_training)
check_three.grid(row=2, column=0, padx=5, pady=5, sticky=W)

s = ttk.Separator(labelframe_resource, orient=HORIZONTAL)
s.grid(row=3, column=0, columnspan=2, sticky=EW)

outage_planner_for_new_DG = IntVar()

check_four = ttk.Checkbutton(labelframe_resource, text="Outage Planner for New DG", variable=outage_planner_for_new_DG)
check_four.grid(row=4, column=0, padx=5, pady=5, sticky=W)

check_one.bind("<1>", lambda e: check_boxes_status())
check_one.bind("<space>", lambda e: check_boxes_status())
check_two.bind("<1>", lambda e: check_boxes_status())
check_two.bind("<space>", lambda e: check_boxes_status())
check_three.bind("<1>", lambda e: check_boxes_status())
check_three.bind("<space>", lambda e: check_boxes_status())
check_four.bind("<1>", lambda e: check_boxes_status())
check_four.bind("<space>", lambda e: check_boxes_status())

###########

labelframe_links = ttk.Labelframe(root, text='Links to Supporting Documentation (SOW, Planning Specs, etc.)')
labelframe_links.grid(row=1, column=0, padx=10, pady=10, sticky=NW, columnspan=2)

links = Text(labelframe_links, height=5)
links.grid(row=0, column=0, padx=5, pady=5, sticky=NW)

scrollbar = Scrollbar(labelframe_links, command=links.yview)
scrollbar.grid(row=0, column=1, sticky='nsew')
links['yscrollcommand'] = scrollbar.set


###########

labelframe_comments = ttk.Labelframe(root, text='Additional Comments')
labelframe_comments.grid(row=2, column=0, padx=10, pady=10, sticky=NW, columnspan=2)

comments = Text(labelframe_comments, height=5)
comments.grid(row=0, column=0, padx=5, pady=5, sticky=NW)

scrollbar = Scrollbar(labelframe_comments, command=comments.yview)
scrollbar.grid(row=0, column=1, sticky='nsew')
comments['yscrollcommand'] = scrollbar.set

###########

btn_submit = ttk.Button(root, text='Submit', command=send_email)
btn_submit.grid(row=3, column=0, pady=10, sticky=NE)



root.mainloop()