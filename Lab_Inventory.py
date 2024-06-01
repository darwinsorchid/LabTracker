
# --------------------------------------------------------| DEPENDENCIES |------------------------------------------------------------------------------

import openpyxl


#-----------------------------------------------------------| CLASSES |----------------------------------------------------------------------------------
class Product():
    """Product information passed in by the user and written in a file. 
    """
    def __init__(self):
        self.name = input("Product title: ")
        self.company = input("Company title: ")
        self.code = input("Catalogue Number: ")
        self.date = input("Expiration Date: ")
        self.quantity = input("Quantity: ")
    
    def product_info(self):
        pr_info = [self.name, self.company, self.code, self.date, self.quantity]
        return pr_info
    


class Cryovial():
    """Cryovial information passed in by the user and written in a file.
    """

    def __init__(self):
        self.cell_line = input("Cell line: ")
        self.passage = input("Passage: ")
        self.freezing_date = input("Date of freezing: ")
        self.initials = input("Freezing done by (initials only): ")
        self.quantity = input("Number of cryovials: ")
    
    def cryovial_info(self):
        pr_info = [self.cell_line, self.passage, self.freezing_date, self.initials, self.quantity]
        return pr_info


#------------------------------------------------------------/ FUNCTIONS /-------------------------------------------------------------------------------

def take_pr_info():
    print("Welcome to LabTracker Product Inventory!")
    num = int(input("How many products do you want to add today? "))
    products = []
    for n in range(num):
        product = Product()
        products.append(product.product_info())

# Write product info to xl file and save: 
    wb = openpyxl.load_workbook("Lab_Inventory.xlsx")
    sheet = wb["Products"]
    for row in products:
        sheet.append(row)
    wb.save("Lab_Inventory.xlsx")
    print("Product Inventory Updated")


def take_cr_info():
    print("Welcome to LabTracker Cryovial Inventory!")
    num = int(input("How many cryovials do you want to add today? "))
    cryovials = []
    for n in range(num):
        cryovial = Cryovial()
        cryovials.append(cryovial.cryovial_info())

# Write cryovial info to xl file and save:
    wb = openpyxl.load_workbook("Lab_Inventory.xlsx")
    sheet = wb["Cryovials"]
    for row in cryovials:
        sheet.append(row)
    wb.save("Lab_Inventory.xlsx")
    print("Cryovial Inventory Updated")

#--------------------------------------------------------/ LOGIC /-----------------------------------------------------------------


choice = input("Which inventory do you want to add to today? Choose Products/Cryovials. ")
if choice.lower()[0] == "p":
    take_pr_info()
elif choice.lower()[0] == "c":
    take_cr_info()
