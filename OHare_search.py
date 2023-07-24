import customtkinter as tk
from customtkinter import filedialog
import pandas as pd
from pandastable import Table

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    entry_file_path.delete(0, tk.END)
    entry_file_path.insert(tk.END, file_path)

def search_excel():
    file_path = entry_file_path.get()
    tr_num = entry_tr_num.get()
    ess_num = entry_ess_num.get()
    feeder_num = entry_feeder_num.get()
    station_name_num = entry_station_name_num.get()
    try:
        df = pd.read_excel(file_path)
    except FileNotFoundError:
        return

    search_columns = ['CE TR# or NC Node TED/CEGIS', 'CE Network Center/ESS #', 'CE Feeder', 'Station Name & #']
    search_values = [tr_num, ess_num, feeder_num, station_name_num]

    row_filter = pd.Series([True] * len(df))
    if search_values[0] != '':
        row_filter &= (df[search_columns[0]] == search_values[0])

    if search_values[1] != '':
        row_filter &= (df[search_columns[1]] == search_values[1])
    if search_values[2] != '':
        row_filter &= (df[search_columns[2]] == search_values[2])
    if search_values[3]!='':
        row_filter &= (df[search_columns[3]].astype(str).str.contains(search_values[3], case=False, na=False))
    filtered_df = df[row_filter]


    if filtered_df.empty:
        for widget in result_frame.winfo_children():
            widget.destroy()
        label = tk.CTkLabel(result_frame, text="No matching entries found.")
        label.grid()
    else:
        selected_columns = [column for column, include in zip(filtered_df.columns, checkbox_vars) if include.get()]
        subset_df = filtered_df[selected_columns]
        table = Table(result_frame, dataframe=subset_df, width=1100)
        table.show()

def toggle_select_all_checkboxes():
    state = select_all_var.get()
    for var in checkbox_vars:
        var.set(state)

# Create the main window
window = tk.CTk()
window.title("Excel Search")
window.geometry("800x600")

# Create and position the file path entry and label
label_file_path = tk.CTkLabel(window, text="File Path:")
label_file_path.pack()
entry_file_path = tk.CTkEntry(window, width=400)
entry_file_path.pack()

# Create the browse button
button_browse = tk.CTkButton(window, text="Browse", command=browse_file,border_spacing=2)
button_browse.pack()

# create the search fields frame
search_frame = tk.CTkFrame(window)
search_frame.pack(padx=10, pady=10)

# Create and position the tr_num entry and label
label_tr_num = tk.CTkLabel(search_frame, text="Transformer #:")
label_tr_num.grid(row=0,column=0)
entry_tr_num = tk.CTkEntry(search_frame)
entry_tr_num.grid(row=1,column=0)

# Create and position the ess_num entry and label
label_ess_num = tk.CTkLabel(search_frame, text="ESS:")
label_ess_num.grid(row=0,column=1)
entry_ess_num = tk.CTkEntry(search_frame)
entry_ess_num.grid(row=1,column=1)

# Create and position the feeder_num entry and label
label_feeder_num = tk.CTkLabel(search_frame, text="Feeder #:")
label_feeder_num.grid(row=0,column=2)
entry_feeder_num = tk.CTkEntry(search_frame)
entry_feeder_num.grid(row=1,column=2)

# Create and position the station_name_num entry and label
label_station_name_num = tk.CTkLabel(search_frame, text="Station Name/#:")
label_station_name_num.grid(row=0,column=3)
entry_station_name_num = tk.CTkEntry(search_frame)
entry_station_name_num.grid(row=1,column=3)

# Create and position the checkboxes
checkbox_frame = tk.CTkFrame(window)
checkbox_frame.pack(padx=10, pady=10)

checkbox_vars = []
checkbox_names = ['Building','CDA ID','Meter ID', 'Account',
                  'Account Name','operation','Concession Space',
                  'Equipment_Served', 'Location', 'Changes to Document',
                  'Amperage', 'V6', 'Electrical Comments', 'Unnamed',
                  'CE Customer Name','CE Account Number','CE TR#','ESS #','CE Feeder',
                  'CE Alternate Feeder','CE ATO','Customer ATO',
                  'Station Name & #','Field Verifcation Needed?',
                  '# of Meters Review', 'Blank']

for i in range(len(checkbox_names)):
    # checkbox_var = tk.BooleanVar()
    checkbox_vars.append(tk.BooleanVar())
    checkbox = tk.CTkCheckBox(checkbox_frame, text=f"{checkbox_names[i]}", variable=checkbox_vars[i])
    checkbox.grid(row=(i//4), column=(i%4),sticky="w", padx=5)

# Create the "Select All" checkbox
select_all_var = tk.BooleanVar()
select_all_checkbox = tk.CTkCheckBox(checkbox_frame, text="Select All", variable=select_all_var, command=toggle_select_all_checkboxes)
select_all_checkbox.grid(row=0, columnspan=4, sticky="nsew", padx=5)

# Create the search button
button_search = tk.CTkButton(window, text="Search", command=search_excel)
button_search.pack()

# Create and position the result text with a horizontal scroll-bar
scrollbar = tk.CTkScrollbar(window, orientation="horizontal")
scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

# result_text = tk.CTkTextbox(window, wrap=tk.NONE, height=150, xscrollcommand=scrollbar.set,width=700)
# result_text.pack()
result_frame = tk.CTkFrame(window, width=700)
result_frame.pack(padx=10)

# scrollbar.configure(command=result_frame.xview)
# # Create and position the result label
# result_text = tk.StringVar()
# label_result = tk.Label(window, textvariable=result_text, wraplength=400)
# label_result.pack()

# Start the main event loop
window.mainloop()