import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import xml.etree.ElementTree as ET
import datetime
import csv

global tester_name
global order_name
global num_orders

order_data = []
XML_data = []

def import_country_to_carrier_service_from_csv(filename):
    with open(filename, 'r', encoding='ISO-8859-1') as f:
        reader = csv.reader(f, delimiter=';')  # Use ';' as the delimiter
        next(reader)  # Skip the header row
        dict_obj = {}
        for row in reader:
            country, carrier, service = row
            if country not in dict_obj:
                dict_obj[country] = {}
            if carrier not in dict_obj[country]:
                dict_obj[country][carrier] = []
            dict_obj[country][carrier].append(service)
    return dict_obj

def import_service_to_data_from_csv(filename):
    with open(filename, 'r', encoding='ISO-8859-1') as f:
        reader = csv.reader(f, delimiter=';')  # Use ';' as the delimiter
        next(reader)  # Skip the header row
        dict_obj = {}
        for row in reader:
            service, i_carrier, c_delivery, c_shipto = row
            dict_obj[service] = {'i_carrier': i_carrier, 'c_delivery': c_delivery, 'c_shipto': c_shipto}
    return dict_obj

def import_country_address_dict_from_csv(filename):
    with open(filename, 'r', encoding='ISO-8859-1') as f:
        reader = csv.reader(f, delimiter=';')  # Use ';' as the delimiter
        next(reader)  # Skip the header row
        dict_obj = {}
        for row in reader:
            country, code, city, address, zipcode = row
            dict_obj[country] = {'code': code, 'city': city, 'address': address, 'zipcode': zipcode}
    return dict_obj

def import_brands_from_csv(filename):
    with open(filename, 'r', encoding='ISO-8859-1') as f:
        reader = csv.reader(f, delimiter=';')  # Use ';' as the delimiter
        next(reader)  # Skip the header row
        dict_obj = {}
        for row in reader:
            brand, s_platform, s_brand, s_value = row
            dict_obj[brand] = {'s_platform': s_platform, 's_brand': s_brand, 's_value': s_value}
    return dict_obj

country_to_carrier_service = import_country_to_carrier_service_from_csv('country_to_carrier_service.csv')
service_to_data = import_service_to_data_from_csv('service_to_data.csv')
country_address_dict = import_country_address_dict_from_csv('country_address_dict.csv')
brands = import_brands_from_csv('brands.csv')

default_row = [
    #####################<p_customer>#####################
    '40',  # i_company
    '',  # s_title
    'Tester',  # s_name
    'TesterName',  # s_firstname
    'Waalwijk',  # s_city
    'Sluisweg 38',  # s_address1
    '',  # s_address2
    '',  # s_address3
    '0',  # i_zipcode
    '123456789',  # s_phoneno1
    'ILoveCeva@gmail.com',  # s_email
    '1502088263',  # s_extern_custno
    '528',  # i_country
    '5932 TT',  # s_zipcode
    '4005',  # i_cust_company
    #####################<p_shipto>#####################
    '40',  # i_company
    'Tester',  # s_name
    'TesterName',  # s_firstname
    'Waalwijk',  # s_city
    'Sluisweg 38',  # s_address1
    '',  # s_address2
    '',  # s_address3
    '0',  # i_zipcode
    '653938014',  # s_phoneno1
    '528',  # i_country
    '5932 TV',  # s_zipcode
    #####################<e_ordsum>#####################
    '40',  # i_company
    '258.93',  # f_orderval
    'Y',  # c_pandp
    '0.0',  # f_pandpval
    'E',  # c_ordertype
    'R',  # c_shipto (ordsum)
    '0',  # c_delivery
    'Y',  # c_pandptax
    '19',  # i_pandp_taxcode
    'OL1502088263',  # s_extern_orderno
    '23',  # i_carrier
    'OL1502088263',  # s_extern_orderno_long
    '2023-05-19 08:26:28',  # d_order
    '',  # s_platform
    'ON',  # s_brand
    #####################<e_orddet>#####################
    '40',  # i_company
    '5715097595869',  # s_extern_itemno
    '39.99',  # f_retailprice
    '1',  # i_orderqty
    '0.0',  # f_discountval
    '19',  # i_taxcode
    'N',  # c_itemtype
    'INVOICE',  # s_paymode
    '3',  # i_attrib
    'ON'  # s_value
]

def submit_data():
    global tester_name
    global order_name
    global num_orders

    tester_name = tester_name_entry.get()
    order_name = order_name_entry.get()
    num_orders = int(num_orders_entry.get())

    # Create the list of orders for the dropdown based on the number entered
    orders = [order_name + str(i) for i in range(1, num_orders + 1)]
    order_var.set("")  # Reset the selected order
    order_dropdown['values'] = orders

    # Initialize order_data with empty data for each order
    global order_data
    order_data = [[order, "", "", "", "", "", "", ""] for order in orders]

    # Print the values to verify they are captured correctly
    print("Tester's Name:", tester_name)
    print("Order Name:", order_name)
    print("Number of Orders:", num_orders)

    # Open the second window for order selection and data display
    order_window.deiconify()
    root.withdraw()

root = tk.Tk()
root.configure(background="#f0f0f0")
root.title("Order Data Entry")

# Create labels and entry fields for data input
tester_name_label = ttk.Label(root, text="Tester's Name:")
tester_name_label.grid(row=0, column=0, padx=10, pady=10)
tester_name_entry = ttk.Entry(root)
tester_name_entry.grid(row=0, column=1, padx=10, pady=10)

order_name_label = ttk.Label(root, text="Order Name:")
order_name_label.grid(row=1, column=0, padx=10, pady=10)
order_name_entry = ttk.Entry(root)
order_name_entry.grid(row=1, column=1, padx=10, pady=10)

num_orders_label = ttk.Label(root, text="Number of Orders:")
num_orders_label.grid(row=2, column=0, padx=10, pady=10)
num_orders_entry = ttk.Entry(root)
num_orders_entry.grid(row=2, column=1, padx=10, pady=10)

submit_button = ttk.Button(root, text="Submit", command=submit_data)
submit_button.grid(row=3, columnspan=2, padx=10, pady=10)

def open_order_window():
    selected_order = order_var.get()
    order_window = tk.Toplevel(root)
    order_window.title("Order Window")
    order_window.configure(background="#f0f0f0")

    def save_order_data():
        global order_data
        country = country_var.get()
        carrier = carrier_var.get()
        carrier_service = carrier_service_var.get()
        address = address_entry.get()
        brand_code = brand_code_var.get()
        date = date_entry.get()
        num_items = num_items_entry.get()

        # Check if the order already exists in order_data
        for order in order_data:
            if order[0] == selected_order:
                # Update the existing order's data
                order[1:8] = [country, carrier, carrier_service, address, brand_code, date, num_items]
                break
        else:
            # If the order doesn't exist, append a new entry to order_data
            order_data.append([selected_order, country, carrier, carrier_service, address, brand_code, date, num_items])

        # Print the order data to verify it is captured correctly
        print("Order Data:", order_data)

        # Find the index of selected_order in order_data
        selected_order_index = next(i for i, order in enumerate(order_data) if order[0] == selected_order)

        # Open the "Item Window"
        open_item_window(order_data[selected_order_index])
        order_window.destroy()

    # Create labels and entry fields for order data input
    country_label = ttk.Label(order_window, text="Country:")
    country_label.grid(row=0, column=0, padx=10, pady=10)
    country_var = tk.StringVar()
    country_dropdown = ttk.Combobox(order_window, textvariable=country_var,
                                    values=list(country_to_carrier_service.keys()))
    country_dropdown.set("Country 1")
    country_dropdown.grid(row=0, column=1, padx=10, pady=10)

    carrier_label = ttk.Label(order_window, text="Carrier:")
    carrier_label.grid(row=1, column=0, padx=10, pady=10)
    carrier_var = tk.StringVar()
    carrier_dropdown = ttk.Combobox(order_window, textvariable=carrier_var)
    carrier_dropdown.set("Carrier 1")
    carrier_dropdown.grid(row=1, column=1, padx=10, pady=10)

    carrier_service_label = ttk.Label(order_window, text="Carrier Service:")
    carrier_service_label.grid(row=2, column=0, padx=10, pady=10)
    carrier_service_var = tk.StringVar()
    carrier_service_dropdown = ttk.Combobox(order_window, textvariable=carrier_service_var, width=40)
    carrier_service_dropdown.set("Service 1")
    carrier_service_dropdown.grid(row=2, column=1, padx=10, pady=10)

    # Add trace to country and carrier variables
    country_var.trace('w', lambda *args: update_carriers(country_var, carrier_var, carrier_dropdown))
    carrier_var.trace('w', lambda *args: update_services(country_var, carrier_var, carrier_service_dropdown))

    address_label = ttk.Label(order_window, text="Address 3:")
    address_label.grid(row=3, column=0, padx=10, pady=10)
    address_entry = ttk.Entry(order_window)
    address_entry.grid(row=3, column=1, padx=10, pady=10)

    brand_code_label = ttk.Label(order_window, text="Brand Code:")
    brand_code_label.grid(row=4, column=0, padx=10, pady=10)
    brand_code_var = tk.StringVar()
    brand_code_dropdown = ttk.Combobox(order_window, textvariable=brand_code_var,
                                       values=list(brands.keys()))
    brand_code_dropdown.set("Code 1")
    brand_code_dropdown.grid(row=4, column=1, padx=10, pady=10)

    date_label = ttk.Label(order_window, text="Date:")
    date_label.grid(row=5, column=0, padx=10, pady=10)
    current_datetime = datetime.datetime.now()
    formatted_datetime = current_datetime.strftime("%Y-%m-%d %H:%M:%S")
    date_entry = ttk.Entry(order_window)
    date_entry.insert(0, f"{formatted_datetime}")
    date_entry.grid(row=5, column=1, padx=10, pady=10)

    num_items_label = ttk.Label(order_window, text="Number of Items:")
    num_items_label.grid(row=6, column=0, padx=10, pady=10)
    num_items_entry = ttk.Entry(order_window)
    num_items_entry.insert(0, "1")
    num_items_entry.grid(row=6, column=1, padx=10, pady=10)

    save_button = ttk.Button(order_window, text="Save", command=save_order_data)
    save_button.grid(row=7, columnspan=2, padx=10, pady=10)

    def update_carriers(country_var, carrier_var, carrier_dropdown):
        country = country_var.get()
        carrier_var.set('')
        carrier_dropdown['values'] = list(country_to_carrier_service.get(country, {}).keys())

    def update_services(country_var, carrier_var, carrier_service_dropdown):
        country = country_var.get()
        carrier = carrier_var.get()
        carrier_service_var.set('')
        carrier_service_dropdown['values'] = country_to_carrier_service.get(country, {}).get(carrier, [])

def save_item_data(item_window):
    # Retrieve data from the item fields
    item_fields = []
    for entry in item_entries:
        item_value = entry.get()
        if not item_value:
            # If any item field is empty, show an error message and return
            messagebox.showerror("Error", "Item fields cannot be empty.")
            return
        item_fields.append(item_value)

    # Retrieve the selected order from order_var
    selected_order = order_var.get()

    # Find the selected order in order_data
    for order in order_data:
        if order[0] == selected_order:
            # Replace the item data for the order
            order[8:] = item_fields  # Update the item data starting from index 8
            break

    # Close the item window and return to the order selection window
    print("Order Data:", order_data)
    item_window.destroy()
    order_window.deiconify()

def open_item_window(order_data):
    order_window.withdraw()

    # Create the item window
    item_window = tk.Toplevel(root)
    item_window.attributes("-topmost", True)
    item_window.title("Item Window")
    item_window.geometry("530x300")
    item_window.configure(background="#f0f0f0")

    num_items = int(order_data[7])  # Get the number of items from order_data

    # Create a scrollable frame to hold the item fields and "Save" button
    scroll_frame = tk.Frame(item_window, bg="#f0f0f0")
    scroll_frame.pack(fill=tk.BOTH, expand=True)

    # Create a canvas and a scrollbar
    canvas = tk.Canvas(scroll_frame, bg="#f0f0f0")
    scrollbar = tk.Scrollbar(scroll_frame, orient=tk.VERTICAL, command=canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    canvas.configure(yscrollcommand=scrollbar.set)

    # Create a frame inside the canvas to hold the item fields
    item_frame = tk.Frame(canvas, bg="#f0f0f0")
    canvas.create_window((0, 0), window=item_frame, anchor=tk.NW)

    # Create data fields for each item
    global item_entries
    item_entries = []
    for i in range(num_items):
        item_number_label = ttk.Label(item_frame, text=f"Item Number {i+1}:")
        item_number_label.grid(row=i, column=0, padx=10, pady=10)
        item_number_entry = ttk.Entry(item_frame)
        item_number_entry.grid(row=i, column=1, padx=10, pady=10)
        item_entries.append(item_number_entry)

        order_quantity_label = ttk.Label(item_frame, text="Order Quantity:")
        order_quantity_label.grid(row=i, column=2, padx=10, pady=10)
        order_quantity_entry = ttk.Entry(item_frame)
        order_quantity_entry.insert(0, "1")
        order_quantity_entry.grid(row=i, column=3, padx=10, pady=10)
        item_entries.append(order_quantity_entry)

    # Create the "Save" button inside the item_frame
    save_button = ttk.Button(item_frame, text="Save", command=lambda: save_item_data(item_window))
    save_button.grid(row=num_items, columnspan=4, padx=10, pady=10)

    # Update the scroll region of the canvas
    item_frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))

def update_item_data_display():
    global item_data

    # Retrieve the item data for the selected order
    selected_order = order_var.get()
    selected_order_data = [data for data in order_data if data[0] == selected_order]

    if selected_order_data:
        # Retrieve the item data for the selected order
        item_data = selected_order_data[0][8:]

        # Create the item data display string with the desired format
        item_display_string = "\tItems Data:\n\n"
        for i in range(0, len(item_data), 2):
            item_display_string += f"Item {i // 2 + 1}:\n"
            item_display_string += f"-{item_data[i]}\n"
            if i + 1 < len(item_data):
                item_display_string += f"-{item_data[i + 1]}\n"

        # Set the item data display label text
        item_data_label.config(text=item_display_string)
    else:
        # Clear the item data display label if no order data is found
        item_data_label.config(text="")

def update_data_display(*args):
    global order_data

    selected_order = order_var.get()

    # Find the selected order data in order_data
    selected_order_data = [data for data in order_data if data[0] == selected_order]

    if selected_order_data:
        # Convert the elements to strings and display the order data in the data display label
        order_display_string = (
            f"\t{selected_order_data[0][0]}\n\n"
            f"Country: {selected_order_data[0][1]}\n"
            f"Carrier: {selected_order_data[0][2]}\n"
            f"Carrier Service: {selected_order_data[0][3]}\n"
            f"Address 3: {selected_order_data[0][4]}\n"
            f"Brand code: {selected_order_data[0][5]}\n"
            f"Date: {selected_order_data[0][6]}\n"
            f"Number of items: {selected_order_data[0][7]}"
        )
        data_display.config(text=order_display_string)

        # Update the item data display
        update_item_data_display()
    else:
        # Clear the data display label if no order data is found
        data_display.config(text="")

        # Update the item data display
        update_item_data_display()



# Create the second window for order selection and data display
order_window = tk.Toplevel(root)
order_window.title("Order Selection")
order_window.withdraw()  # Hide the second window initially
order_window.configure(background="#f0f0f0")

# Create the order selection dropdown box
order_label = ttk.Label(order_window, text="Select Order:")
order_label.grid(row=0, column=0, padx=10, pady=10)
order_var = tk.StringVar()
order_var.trace('w', update_data_display)  # Call update_data_display when the selected order changes
order_dropdown = ttk.Combobox(order_window, textvariable=order_var, state="readonly")
order_dropdown.grid(row=0, column=1, padx=10, pady=10)

# Create the data display for the selected order
data_display_label = ttk.Label(order_window, text="Data Display:")
data_display_label.grid(row=1, column=0, padx=10, pady=10)
data_display = ttk.Label(order_window, text="", background="white", relief="solid", anchor="w")
data_display.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

# Create the label for item data display
item_label = ttk.Label(order_window, text="Items Data:")
item_label.grid(row=3, column=0, padx=10, pady=10, sticky="w")
item_data_label = ttk.Label(order_window, text="", background="white", relief="solid", anchor="w")
item_data_label.grid(row=3, column=1, padx=10, pady=10, sticky="ew")

# Call the refresh_data_display function when the order selection window is opened
order_window.bind("<Visibility>", lambda event: update_data_display())

def apply_to_all_orders():
    global order_data

    selected_order = order_var.get()

    # Find the selected order data in order_data
    selected_order_data = [data for data in order_data if data[0] == selected_order]

    if selected_order_data:
        # Get the index of the selected order
        selected_order_index = order_data.index(selected_order_data[0])
        # Copy the data from the selected order to all other orders
        for i, order in enumerate(order_data):
            if i != selected_order_index:
                order[1:] = selected_order_data[0][1:]  # Copy all data fields including items data

        # Update the data display for all orders
        update_data_display()

        # Update the item data display for all orders
        update_item_data_display()

def apply_one_to_all_window():
    global dropdown_field
    global chosen_value
    selected_order = order_var.get()

    field_names = ["Country", "Carrier", "Carrier Service", "Address 3", "Brand Code", "Date"]

    # Create the apply_one_to_all window
    apply_one_to_all_window = tk.Toplevel(root)
    apply_one_to_all_window.title("Apply One to All")
    apply_one_to_all_window.configure(background="#f0f0f0")

    field_label = ttk.Label(apply_one_to_all_window, text="Choose an option:")
    field_label.grid(row=0, column=0, padx=10, pady=10)
    field_var = tk.StringVar()
    field_dropdown = ttk.Combobox(apply_one_to_all_window, textvariable=field_var, values=field_names)
    field_dropdown.set("Country")
    field_dropdown.grid(row=0, column=1, padx=10, pady=10)

    new_value_label = ttk.Label(apply_one_to_all_window, text="New Value:")
    new_value_label.grid(row=1, column=0, padx=10, pady=10)

    new_value_frame = tk.Frame(apply_one_to_all_window, bg="#f0f0f0")
    new_value_frame.grid(row=1, column=1, padx=10, pady=10)

    def update_new_value_widget(*args):
        global dropdown_field
        global chosen_value

        selected_field = field_var.get()
        selected_order_index = next((index for index, order in enumerate(order_data) if order[0] == selected_order),
                                    None)

        dropdown_field = 1  # Default to dropdown

        # Clear the existing widget
        for widget in new_value_frame.winfo_children():
            widget.destroy()

        if selected_field == "Country":
            choices = list(country_to_carrier_service.keys())
        elif selected_field == "Carrier":
            selected_country = order_data[selected_order_index][1]
            choices = list(country_to_carrier_service.get(selected_country, {}).keys())
        elif selected_field == "Carrier Service":
            selected_country = order_data[selected_order_index][1]
            selected_carrier = order_data[selected_order_index][2]
            choices = country_to_carrier_service.get(selected_country, {}).get(selected_carrier, [])
        elif selected_field == "Brand Code":
            choices = list(brands.keys())
        elif selected_field == "Address 3" or selected_field == "Date":
            dropdown_field = 0

        if dropdown_field == 1:
            new_value_var = tk.StringVar()
            new_value_dropdown = ttk.Combobox(new_value_frame, textvariable=new_value_var, values=choices)
            new_value_dropdown.grid(row=0, column=0, padx=10, pady=10)
            chosen_value = new_value_dropdown
        elif dropdown_field == 0:
            new_value_var = tk.StringVar()
            new_value_entry = ttk.Entry(new_value_frame, textvariable=new_value_var)
            new_value_entry.grid(row=0, column=0, padx=10, pady=10)
            chosen_value = new_value_entry

    # Bind the trace method to the field_var
    field_var.trace('w', update_new_value_widget)

    def apply_change():

        new_value = chosen_value.get()

        # Update the selected field in all orders
        selected_field = field_var.get()
        for order in order_data:
            field_index = field_names.index(selected_field)
            order[field_index+1] = new_value

        # Update the data display for all orders
        update_data_display()

        apply_one_to_all_window.destroy()

    apply_button = ttk.Button(apply_one_to_all_window, text="Apply", command=apply_change)
    apply_button.grid(row=2, columnspan=2, padx=10, pady=10)

# Create the buttons for applying data
select_button = ttk.Button(order_window, text="Select", command=open_order_window)
select_button.grid(row=2, column=0, padx=10, pady=10)

apply_all_button = ttk.Button(order_window, text="Apply to All", command=apply_to_all_orders)
apply_all_button.grid(row=2, column=1, padx=10, pady=10)

apply_one_button = ttk.Button(order_window, text="Apply One to All", command=apply_one_to_all_window)
apply_one_button.grid(row=2, column=2, padx=10, pady=10)

def export_country_to_carrier_service_to_csv(dict_obj, filename):
    with open(filename, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f, delimiter=';', lineterminator='\r\n')  # Use ';' as the delimiter
        writer.writerow(['Country - code', 'Carrier', 'Service'])  # Write the header row
        for country, carrier_services in dict_obj.items():
            for carrier, services in carrier_services.items():
                for service in services:
                    writer.writerow([country, carrier, service])

def export_service_to_data_to_csv(dict_obj, filename):
    with open(filename, 'w', newline='', encoding='ISO-8859-1') as f:
        writer = csv.writer(f, delimiter=';', lineterminator='\r\n')  # Use ';' as the delimiter
        writer.writerow(['Service', 'i_carrier', 'c_delivery', 'c_shipto'])  # Write the header row
        for service, values in dict_obj.items():
            i_carrier = values.get('i_carrier', '')
            c_delivery = values.get('c_delivery', '')
            c_shipto = values.get('c_shipto', '')
            writer.writerow([service, i_carrier, c_delivery, c_shipto])

def export_country_address_dict_to_csv(dict_obj, filename):
    with open(filename, 'w', newline='', encoding='ISO-8859-1') as f:
        writer = csv.writer(f, delimiter=';', lineterminator='\r\n')  # Use ';' as the delimiter
        writer.writerow(['Country', 'Code', 'City', 'Address', 'Zipcode'])  # Write the header row
        for country, values in dict_obj.items():
            code = values.get('code', '')
            city = values.get('city', '')
            address = values.get('address', '')
            zipcode = values.get('zipcode', '')
            writer.writerow([country, code, city, address, zipcode])

def export_brands_to_csv(dict_obj, filename):
    with open(filename, 'w', newline='', encoding='ISO-8859-1') as f:
        writer = csv.writer(f, delimiter=';', lineterminator='\r\n')  # Use ';' as the delimiter
        writer.writerow(['Brand', 's_platform', 's_brand', 's_value'])  # Write the header row
        for brand, values in dict_obj.items():
            s_platform = values.get('s_platform', '')
            s_brand = values.get('s_brand', '')
            s_value = values.get('s_value', '')
            writer.writerow([brand, s_platform, s_brand, s_value])

#CSV_button = ttk.Button(order_window, text="Delete", command=)
#CSV_button.grid(row=5, column=2, padx=10, pady=10)

def data_to_xml():
    global XML_data
    XML_data = []  # Clear the existing data
    for i in range(num_orders):
        # Start with a copy of the default row
        row = default_row.copy()

        key_country = order_data[i][1]
        values = country_address_dict.get(key_country)
        if values is not None:
            code = values.get("code")
            city = values.get("city")
            address = values.get("address")
            zipcode = values.get("zipcode")
        else:
            code = city = address = zipcode = None  # or some default values

        key_brand = order_data[i][5]
        values = brands.get(key_brand)

        if values is not None:
            s_platform = values.get("s_platform")
            s_brand = values.get("s_brand")
            s_value = values.get("s_value")
        else:
            s_platform = s_brand = s_value = None

        key_service = order_data[i][3]  # assuming the 9th column in order_data corresponds to the service
        values = service_to_data.get(key_service)
        if values is not None:
            i_carrier = values.get("i_carrier")
            c_delivery = values.get("c_delivery")
            c_shipto = values.get("c_shipto")
        else:
            i_carrier = c_delivery = c_shipto = None  # or some default values

        row[3] = tester_name # s_firstname
        row[17] = tester_name # s_firstname
        row[35] = order_data[i][0] # s_extern_orderno_long
        row[37] = order_data[i][0] # s_extern_orderno_long
        row[12] = code
        row[24] = code
        row[4] = city
        row[18] = city
        row[5] = address
        row[19] = address
        row[13] = zipcode
        row[25] = zipcode
        row[32] = c_delivery
        row[31] = c_shipto
        row[36] = i_carrier
        row[38] = order_data[i][6]
        row[39] = s_platform
        row[40] = s_brand
        row[50] = s_value

        if order_data[i][5] == 'EMP':
            row[35] = f"8{order_data[i][0]}"
            row[37] = f"8{order_data[i][0]}"

        # Append the updated row to XML_data
        XML_data.append(row)


def indent(elem, level=0):
    i = "\n" + level*"       "  # Four spaces for each level of indentation
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "       "
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        for elem in elem:
            indent(elem, level+1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i

def replace_first_line(filename, new_line):
    with open(filename, 'r') as file:
        lines = file.readlines()

    lines[0] = new_line + '\n'

    with open(filename, 'w') as file:
        file.writelines(lines)

def generate_xml():
    #export_country_to_carrier_service_to_csv(country_to_carrier_service, 'country_to_carrier_service.csv')
    #export_service_to_data_to_csv(service_to_data, 'service_to_data.csv')
    #export_country_address_dict_to_csv(country_address_dict, 'country_address_dict.csv')
    #export_brands_to_csv(brands, 'brands.csv')
    data_to_xml()

    # Create the root element
    root = ET.Element('amos_orderlist', attrib={
        'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
        'xsi:noNamespaceSchemaLocation': 'orderlist_import.xsd'
    })

    for i in range(num_orders):
        # Create an order
        order = ET.SubElement(root, 'amos_order')

        # Create customer details
        customer = ET.SubElement(order, 'p_customer')
        ET.SubElement(customer, 'i_company').text = XML_data[i][0]
        ET.SubElement(customer, 's_title').text = XML_data[i][1]
        ET.SubElement(customer, 's_name').text = XML_data[i][2]
        ET.SubElement(customer, 's_firstname').text = XML_data[i][3]
        ET.SubElement(customer, 's_city').text = XML_data[i][4]
        ET.SubElement(customer, 's_address1').text = XML_data[i][5]
        ET.SubElement(customer, 's_address2').text = XML_data[i][6]
        ET.SubElement(customer, 's_address3').text = XML_data[i][7]
        ET.SubElement(customer, 'i_zipcode').text = XML_data[i][8]
        ET.SubElement(customer, 's_phoneno1').text = XML_data[i][9]
        ET.SubElement(customer, 's_email').text = XML_data[i][10]
        ET.SubElement(customer, 's_extern_custno').text = XML_data[i][11]
        ET.SubElement(customer, 'i_country').text = XML_data[i][12]
        ET.SubElement(customer, 's_zipcode').text = XML_data[i][13]
        ET.SubElement(customer, 'i_cust_company').text = XML_data[i][14]

        # Create shipping details
        shipto = ET.SubElement(order, 'p_shipto')
        ET.SubElement(shipto, 'i_company').text = XML_data[i][15]
        ET.SubElement(shipto, 's_name').text = XML_data[i][16]
        ET.SubElement(shipto, 's_firstname').text = XML_data[i][17]
        ET.SubElement(shipto, 's_city').text = XML_data[i][18]
        ET.SubElement(shipto, 's_address1').text = XML_data[i][19]
        ET.SubElement(shipto, 's_address2').text = XML_data[i][20]
        ET.SubElement(shipto, 's_address3').text = XML_data[i][21]
        ET.SubElement(shipto, 'i_zipcode').text = XML_data[i][22]
        ET.SubElement(shipto, 's_phoneno1').text = XML_data[i][23]
        ET.SubElement(shipto, 'i_country').text = XML_data[i][24]
        ET.SubElement(shipto, 's_zipcode').text = XML_data[i][25]

        # Create order summary
        ordsum = ET.SubElement(order, 'e_ordsum')
        ET.SubElement(ordsum, 'i_company').text = XML_data[i][26]
        ET.SubElement(ordsum, 'f_orderval').text = XML_data[i][27]
        ET.SubElement(ordsum, 'c_pandp').text = XML_data[i][28]
        ET.SubElement(ordsum, 'f_pandpval').text = XML_data[i][29]
        ET.SubElement(ordsum, 'c_ordertype').text = XML_data[i][30]
        ET.SubElement(ordsum, 'c_shipto').text = XML_data[i][31]
        ET.SubElement(ordsum, 'c_delivery').text = XML_data[i][32]
        ET.SubElement(ordsum, 'c_pandptax').text = XML_data[i][33]
        ET.SubElement(ordsum, 'i_pandp_taxcode').text = XML_data[i][34]
        ET.SubElement(ordsum, 's_extern_orderno').text = XML_data[i][35]
        ET.SubElement(ordsum, 'i_carrier').text = XML_data[i][36]
        ET.SubElement(ordsum, 's_extern_orderno_long').text = XML_data[i][37]
        ET.SubElement(ordsum, 'd_order').text = XML_data[i][38]
        ET.SubElement(ordsum, 's_platform').text = XML_data[i][39]
        ET.SubElement(ordsum, 's_brand').text = XML_data[i][40]

        zmienna = 8
        zakres = int(order_data[i][7])
        for j in range(zakres):
            XML_data[i][42] = order_data[i][zmienna]
            XML_data[i][44] = order_data[i][zmienna+1]
            # Create order details
            orddet = ET.SubElement(order, 'e_orddet')
            ET.SubElement(orddet, 'i_company').text = XML_data[i][41]
            ET.SubElement(orddet, 'i_detail').text = str(j + 1)
            ET.SubElement(orddet, 's_extern_itemno').text = XML_data[i][42]
            ET.SubElement(orddet, 'f_retailprice').text = XML_data[i][43]
            ET.SubElement(orddet, 'i_orderqty').text = XML_data[i][44]
            ET.SubElement(orddet, 'f_discountval').text = XML_data[i][45]
            ET.SubElement(orddet, 'i_taxcode').text = XML_data[i][46]
            ET.SubElement(orddet, 'c_itemtype').text = XML_data[i][47]
            zmienna = zmienna + 2

        # Create order summary for print
        ordsum_4print = ET.SubElement(order, 'p_ordsum_4print')
        ET.SubElement(ordsum_4print, 's_paymode').text = XML_data[i][48]

        for j in range(zakres):
            # Create order details for print
            orddet_4print = ET.SubElement(order, 'p_orddet_4print')
            ET.SubElement(orddet_4print, 'i_detail').text = str(j + 1)

        # Create order summary attributes
        ordsum_attrib = ET.SubElement(order, 'e_ordsum_attrib')
        ET.SubElement(ordsum_attrib, 'i_attrib').text = XML_data[i][49]
        ET.SubElement(ordsum_attrib, 's_value').text = XML_data[i][50]

    # Get the file path using a file dialog
    file_path = filedialog.asksaveasfilename(defaultextension='.xml', filetypes=(('XML files', '*.xml'), ('All files', '*.*')))

    indent(root)

    # Write the XML to the specified file path
    tree = ET.ElementTree(root)
    tree.write(file_path, encoding='UTF-8', xml_declaration=True)

    replace_first_line(file_path, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')

generate_button = ttk.Button(order_window, text="Generate", command=generate_xml)
generate_button.grid(row=4, column=1, padx=10, pady=10)

def close_application():
    order_window.quit()
    order_window.destroy()

close_button = ttk.Button(order_window, text="Exit", command=close_application)
close_button.grid(row=5, column=1, padx=10, pady=10)

root.mainloop()