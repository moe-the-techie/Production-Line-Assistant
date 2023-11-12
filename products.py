# This file will have the functions needed to calculate the costs for each product/product type
# I'm thinking of a method for each product where it returns a formatted string that shoinvoice_sheet all the costs
# We will also place the constant values in here and ensure we can't and won't use them elsewhere
# We'll also need functions to modify the certain attributes of each function and update their values

import openpyxl

PRODUCT_CODES = {
    "Linear Slot Diffuser at 45Deg Angle": "LSD45^"
}
N_CORNERS = 4
POWDER_PRICE_PER_KG = 15
SHEET_SIZE = 2440 * 1220
global discount


def add_product(code, discount_inputted, product_count=1, current_workbooks=None):
    """
    Adds a product to the current workbook by picking the right method for the code entered.
    :param discount_inputted: a float of the number part of the discount percentage to be applied to all products
    :param code: string representing product code
    :param product_count: an integer representing the number of products entered so far
    :param current_workbooks: a tuple containing two openpyxl workbooks each representing a report file to be saved
    :return:
    """

    global discount
    discount = discount_inputted

    try:
        if code == "LSD45^":

            n_slots = int(input("Enter number of slots: "))
            gap_size = float(input("Enter gap size in mm: "))
            lsd_length = float(input("Enter length of linear slot diffuser in mm: "))

            # Ensure all values are positive before calculating
            if n_slots <= 0 or gap_size <= 0 or lsd_length <= 0:
                raise ValueError

            linear_slot_diffuser(n_slots, gap_size, lsd_length, product_count,current_workbooks)

    except ValueError:

        if current_workbooks is None:
            print("Invalid value entered\nProcess terminated.")

        else:
            current_workbooks[0].save("Invoice.xlsx")
            current_workbooks[1].save("Internal Report.xlsx")
            print("Invalid value entered\nOutput files saved. Process terminated.")
        exit(1)


def linear_slot_diffuser(n_slots, gap_size, lsd_length, product_count, current_workbooks):
    """
    Calculates the costs of production for the Linear Slot Diffuser and exports it into an Excel file.
    :param n_slots: integer representing the number of slots
    :param gap_size: float representing the size of the gap in millimeters
    :param lsd_length: float representing the length of the linear slot diffuser in millimeters
    :param current_workbooks: a tuple containing two openpyxl workbooks each representing a report file to be saved
    :param product_count: an integer representing the number of products entered so far
    :return:
    """

    # Prices are specific for LSD per 6 meters and the change every 3 months approximately
    outer_frame_price = 7
    inner_frame_price = 5.5
    louver_price = 1.5
    pipe_price = 1.2

    # Prices per unit
    sheet_price = 220
    hanging_clamp_price = 0.5
    corner_price = 0.5

    # Calculations for: outer & inner frames, space bar, and pipe

    outer_frame_thickness = 4.4
    inner_frame_thickness = 1.2
    n_inner_frames = n_slots - 1
    space_bar_size = gap_size + 16

    pipe_length = gap_size + 16 * n_slots + inner_frame_thickness * n_inner_frames + 8.8

    n_pipes = int(round((lsd_length - 200)) / 300) + 1
    n_space_bars = n_slots * n_pipes

    pipe_size = pipe_length * n_pipes

    end_cap_size = 2 * pipe_size

    outer_frame_size = lsd_length * 2 + 10 + 380 + end_cap_size

    inner_frame_size = lsd_length * n_inner_frames

    n_louvers = n_dampers = n_hanging_clamps = n_slots * 2

    louver_size = n_louvers * lsd_length

    # Calculate the percentage of aluminum sheet used for the dampers
    sheet_percentage_used = (lsd_length + 10) * space_bar_size * n_dampers / SHEET_SIZE

    powder_weight = 0.0001333333333 * lsd_length

    # Cost calculations

    material_cost = (outer_frame_size * outer_frame_price + inner_frame_size * inner_frame_price +
                     louver_size * louver_price + pipe_size * pipe_price + sheet_percentage_used *
                     sheet_price + corner_price * N_CORNERS + n_hanging_clamps * hanging_clamp_price + powder_weight
                     * POWDER_PRICE_PER_KG) / 1000

    labor_cost = material_cost * 0.3375

    overhead_cost = labor_cost * 1.5

    total_cost = labor_cost + material_cost + overhead_cost

    unit_price = material_cost * 4
    unit_price *= 0.3

    if current_workbooks is None:
        invoice = openpyxl.load_workbook("customer template.xlsx")
        report = openpyxl.load_workbook("material template.xlsx")

    else:
        invoice = current_workbooks[0]
        report = current_workbooks[1]

    report_sheet = report.active
    invoice_sheet = invoice.active

    # Add discount percentage to the invoice
    invoice_sheet.cell(row=31, column=16).value = str(discount) + "%"

    invoice_index = product_count + 18
    report_index = product_count + 1

    if product_count > 13:
        print("Sheet full, can't add product, saving file & terminating program.")
        report.save("Internal Report.xlsx")
        invoice.save("Invoice.xlsx")
        exit(0)

    # Entering the data into the right row, notice that the 1's represent the quantity and should be updated
    # to whatever value that ends up taking on
    invoice_sheet.cell(row=invoice_index, column=2).value = PRODUCT_CODES["Linear Slot Diffuser at 45Deg Angle"]
    invoice_sheet.cell(row=invoice_index, column=3).value = "Linear Slot Diffuser at 45Deg Angle"
    invoice_sheet.cell(row=invoice_index, column=7).value = lsd_length
    invoice_sheet.cell(row=invoice_index, column=8).value = 1
    invoice_sheet.cell(row=invoice_index, column=12).value = unit_price
    invoice_sheet.cell(row=invoice_index, column=13).value = unit_price - (unit_price * discount / 100)
    invoice_sheet.cell(row=invoice_index, column=14).value = unit_price * 1

    report_sheet.cell(row=report_index, column=1).value = "LSD_45^"
    report_sheet.cell(row=report_index, column=2).value = str(round(total_cost, 2)) + "SAR"
    report_sheet.cell(row=report_index, column=3).value = str(round(material_cost, 2)) + "SAR"
    report_sheet.cell(row=report_index, column=4).value = str(round(labor_cost, 2)) + "SAR"
    report_sheet.cell(row=report_index, column=5).value = str(overhead_cost) + "SAR"
    report_sheet.cell(row=report_index, column=6).value = str(outer_frame_size) + "mm"
    report_sheet.cell(row=report_index, column=7).value = str(inner_frame_size) + "mm"
    report_sheet.cell(row=report_index, column=8).value = str(louver_size) + "mm"
    report_sheet.cell(row=report_index, column=9).value = str(n_pipes) + " pipe"
    report_sheet.cell(row=report_index, column=10).value = str(pipe_size) + "mm"
    report_sheet.cell(row=report_index, column=11).value = str(n_space_bars) + " bars"
    report_sheet.cell(row=report_index, column=12).value = str(space_bar_size) + "mm"
    report_sheet.cell(row=report_index, column=13).value = str(end_cap_size) + "mm"
    report_sheet.cell(row=report_index, column=14).value = str(round(sheet_percentage_used, 2)) + "%"
    report_sheet.cell(row=report_index, column=15).value = str(round(powder_weight, 2)) + "kg"

    try:
        if (input("Product Added.\nDo you wish to add another product? (y: yes, anything else: exit): ")
                .lower() in ["y", "yes"]):
            print("PRODUCT --> PRODUCT_CODE:\n")

            for product_name in PRODUCT_CODES:
                print(product_name + " --> " + PRODUCT_CODES[product_name])

            code = input("\nEnter Product Code: ")

            if code not in PRODUCT_CODES.values():
                raise ValueError

            add_product(code, discount, product_count=product_count+1, current_workbooks=(invoice, report))

        else:

            print("Saving output file and terminating process")
            report.save("Internal Report.xlsx")
            invoice.save("Invoice.xlsx")
            exit(0)

    except ValueError:
        report.save("Internal Report.xlsx")
        invoice.save("Invoice.xlsx")
        print("Invalid value entered\nOutput file saved. Process terminated.")
        exit(1)

    # Calculates the primary cost and returns it alongside the product description.
    # prim_cost = direct_material + direct_labor + overhead
    # Calculates the customer cost which is for now prim_cost + 0.25 * prim_cost.
    # Ask for discount percentage to apply if any.
