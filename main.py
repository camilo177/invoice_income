import pandas as pd
import tkinter
from tkinter import messagebox, simpledialog, ttk
import sqlite3
from docxtpl import DocxTemplate
import datetime
import openpyxl
import matplotlib.pyplot as plt

from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image
from flask import Flask, render_template, request

app = Flask(__name__)

conn = sqlite3.connect("C:/Users/camil/database/invoices.db")
cursor = conn.cursor()

# Create the invoices table if it doesn't exist
cursor.execute("""
    CREATE TABLE IF NOT EXISTS invoices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        total REAL,
        income REAL DEFAULT (total)
    )
""")
conn.commit()

# ...


#Function to delete spaces on the treeview
def clear_item():
    spinbox_quantity.delete(0, tkinter.END)
    spinbox_quantity.insert(0, "1")
    entry_description.delete(0, tkinter.END)
    entry_total_amount.delete(0, tkinter.END)
    
    
#Function to add items to the treeview of the invoice
invoice_list = []
def add_item():
    qty = int(spinbox_quantity.get())
    desc = entry_description.get()
    price = float(entry_total_amount.get())
    total = qty*price
    invoice_item = [qty, desc, price, total]
    
    tree.insert('', 0, values=invoice_item)    
    clear_item()
    
    invoice_list.append(invoice_item)

 
#Function to start a new invoice
def new_invoice():
    entry_customer_first_name.delete(0, tkinter.END)
    entry_customer_surname.delete(0, tkinter.END)
    entry_phone.delete(0, tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())  
    
    invoice_list.clear()  
    
    
#Function to create the word invoice
def create_invoice():
    name = entry_customer_first_name.get() + entry_customer_surname.get()
    phone = entry_phone.get()
    subtotal = sum(item[3] for item in invoice_list)
    salestax = 0.19  # Modify according to the country
    total = subtotal * (1 + salestax)
    description = entry_description.get()  

    try:
        with sqlite3.connect("C:/Users/camil/database/invoices.db") as conn:
            cursor = conn.cursor()

            # Insert the invoice information into the database
            cursor.execute("""
                INSERT INTO invoices (name, total)
                VALUES (?, ?)
            """, (name, total))
            conn.commit()

            # Calculate the total income
            cursor.execute("SELECT SUM(total) FROM invoices")
            total_income = cursor.fetchone()[0]

            # Update the income column in all rows
            cursor.execute("UPDATE invoices SET income = ?", (total_income,))
            conn.commit()

            messagebox.showinfo("Invoice Complete", "Invoice data saved successfully.")

    except sqlite3.Error as e:
        messagebox.showerror("Error", str(e))
    finally:
        cursor.close()
        conn.close()

    # Prepare the invoice data for rendering
    invoice_data = {
        "name": name,
        "phone": phone,
        "invoice_list": invoice_list,
        "salestax": salestax,
        "total": total,
        "description": description,
    }

    # Word document generation code
    doc = DocxTemplate("./invoice_template.docx")
    doc.render(invoice_data) 

    doc_name = "new_invoice" + name + datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S") + ".docx"
    doc.save(doc_name)
    messagebox.showinfo("Invoice Complete", "Invoice Word document saved successfully.")
    new_invoice()

     
def calculate_total_income():
    try:
        # Retrieve the invoice data from the database
        cursor.execute("SELECT * FROM invoices")
        data = cursor.fetchall()

        # Calculate the total income from all invoices
        total_income = sum(item[3] for item in data)

        # Prompt the user to enter the expenses
        expenses = simpledialog.askfloat("Calculate Income", "Enter the expenses:")
        if expenses is None:
            return

        # Calculate the profit by subtracting expenses from total income
        profit = total_income - expenses

        # Prepare data for the chart
        chart_data = {"Category": ["Income", "Expenses", "Profit"],
                      "Amount": [total_income, expenses, profit]}

        # Create a DataFrame from the chart data
        df = pd.DataFrame(chart_data)

        # Create the plot using Matplotlib
        plt.bar(df["Category"], df["Amount"])
        plt.xlabel("Category")
        plt.ylabel("Amount")
        plt.title("Income and Expenses")

        # Display the plot
        plt.show()

        # Export the DataFrame to an Excel file
        file_name = f"income_report_{datetime.date.today()}.xlsx"
        with pd.ExcelWriter(file_name, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Income Data")

        messagebox.showinfo("Total Income", f"The total income of the company is: {total_income}\nExpenses: {expenses}\nProfit: {profit}\nIncome data and chart exported to {file_name}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while calculating the total income:\n{str(e)}")
              
#Interface creation
window = tkinter.Tk()
window.title("Facturación QUEENS")
window.geometry("400x400")
#Second container
frame = tkinter.Frame(window)
frame.pack(padx=20, pady=10)


#Creation of the spaces for the form
label_customer_first_name = tkinter.Label(frame, text="First Name:")
label_customer_first_name.grid(row=0, column=0)
entry_customer_first_name = tkinter.Entry(frame)
entry_customer_first_name.grid(row=1, column=0)

label_customer_surname = tkinter.Label(frame, text="Last Name:")
label_customer_surname.grid(row=0, column=1)
entry_customer_surname = tkinter.Entry(frame)
entry_customer_surname.grid(row=1, column=1)

label_phone = tkinter.Label(frame, text="Phone:")
label_phone.grid(row=0, column=2)
entry_phone = tkinter.Entry(frame)
entry_phone.grid(row=1, column=2)

label_quantity = tkinter.Label(frame, text="Quantity:")
label_quantity.grid(row=2, column=0)
spinbox_quantity = tkinter.Spinbox(frame, from_=1, to=50)
spinbox_quantity.grid(row=3, column=0)

label_description = tkinter.Label(frame, text="Description:")
label_description.grid(row=2, column=1)
entry_description = tkinter.Entry(frame)
entry_description.grid(row=3, column=1)

label_total_amount = tkinter.Label(frame, text="Unit Price:")
label_total_amount.grid(row=2, column=2)
entry_total_amount = tkinter.Entry(frame)
entry_total_amount.grid(row=3, column=2)

#Buttons and table 
button_add_item=tkinter.Button(frame, text="Add Item", command= add_item)
button_add_item.grid(row=4, column=1, columnspan=2)

columns=("Quantity", "Price", "Total")
tree=ttk.Treeview(frame, columns=columns, show="headings")
tree.heading("Quantity", text="Quantity")
tree.heading("Price", text="Price")
tree.heading("Total", text="Total")
tree.grid(row = 6, column=0, columnspan = 3, padx=20, pady=10)


button_create = tkinter.Button(frame, text="Create Invoice", command=create_invoice)
button_create.grid(row=7, column=0, columnspan=3, sticky="news", padx=20, pady=5)

button_new_invoice = tkinter.Button(frame, text="New Invoice", command=new_invoice)
button_new_invoice.grid(row=8, column=0, columnspan=3, sticky="news", padx=20, pady=5)

button_calculate_income = tkinter.Button(frame, text="Calculate Total Income", command=calculate_total_income)
button_calculate_income.grid(row=9, column=0, columnspan=3, sticky="news", padx=20, pady=5)


frame.mainloop()

conn.close()

# Route for the index page
@app.route('/')
def index():
    return render_template('index.html')

if __name__ == '__main__':
    app.run()