import openpyxl

PRODUCT_CODES = {
    "Linear Slot Diffuser at 45Deg Angle": "LSD45^"
}
N_CORNERS = 4
POWDER_PRICE_PER_KG = 20
SHEET_SIZE = 2440 * 1220
global discount


def add_product(code, discount_inputted, quantity=1,  product_count=1, current_workbooks=None):
    """
    Adds a product to the current workbook by picking the right method for the code entered.
    :param discount_inputted: a float of the number part of the discount percentage to be applied to all products
    :param code: string representing product code
    :param quantity: integer representing the quantity of LSDs required
    :param product_count: an integer representing the number of products entered so far
    :param current_workbooks: a tuple containing two openpyxl workbooks each representing a report file to be saved
    :return:
    """

    global discount
    discount = discount_inputted

    # Call product function based on code entered
    try:
        if code == "LSD45^":

            n_slots = int(input("Enter number of slots: "))
            gap_size = float(input("Enter gap size in mm: "))
            lsd_length = float(input("Enter length of linear slot diffuser in mm: "))

            # Ensure all values are positive before calculating
            if n_slots <= 0 or gap_size <= 0 or lsd_length <= 0:
                raise ValueError

            linear_slot_diffuser(n_slots, gap_size, lsd_length, quantity, product_count, current_workbooks)

    except ValueError:

        if current_workbooks is None:
            print("Invalid value entered\nProcess terminated.")

        else:
            current_workbooks[0].save("Invoice.xlsx")
            current_workbooks[1].save("Internal Report.xlsx")
            print("Invalid value entered\nOutput files saved. Process terminated.")
        exit(1)


def linear_slot_diffuser(n_slots, gap_size, lsd_length, quantity, product_count, current_workbooks):
    """
    Calculates the costs of production for the Linear Slot Diffuser and exports it into an Excel file.
    :param n_slots: integer representing the number of slots
    :param gap_size: float representing the size of the gap in millimeters
    :param lsd_length: float representing the length of the linear slot diffuser in millimeters
    :param quantity: integer representing the quantity of LSDs required
    :param product_count: an integer representing the number of products entered so far
    :param current_workbooks: a tuple containing two openpyxl workbooks each representing a report file to be saved
    :return:
    """

    # Prices that are specific for LSD per 6 meters and the change around every 3 months
    outer_frame_price = 6.6
    inner_frame_price = 4.7
    louver_price = 4.8
    pipe_price = 1.6
    space_bar_price = 1.8

    # Calculations for: outer & inner frames, space bar, powder time, and pipe

    n_inner_frames = n_slots - 1
    inner_frame_thickness = 1.2
    outer_frame_thickness = 4.4

    end_cap_size = (20 + 16) * n_slots + n_inner_frames * inner_frame_thickness + 2 * outer_frame_thickness
    space_bar_size = (lsd_length / 350) * end_cap_size - 1

    pipe_size = (lsd_length / 350) * end_cap_size
    outer_frame_size = (lsd_length + end_cap_size) * 2 + 380
    inner_frame_size = n_inner_frames * lsd_length

    louver_size = n_slots * 2 * lsd_length

    powder_weight = 0.0001333333333 * lsd_length

    powder_time = ((90 / (6000 * 12 / lsd_length)) + (90 / (6000 * 36 / lsd_length)) + 1) * 2

    # Cost calculations

    material_cost = (((space_bar_size * space_bar_price) + (inner_frame_size * inner_frame_price) +
                     (outer_frame_size * outer_frame_price) + (louver_size * louver_price) +
                     (pipe_size * pipe_price)) / 1000) + (powder_weight * POWDER_PRICE_PER_KG)

    labor_cost = material_cost * 0.4 + powder_time

    overhead_cost = labor_cost * 0.65

    total_cost = labor_cost + material_cost + overhead_cost

    unit_price = material_cost * 4
    unit_price *= 0.3

    # Creates report openpyxl file
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
    report_index = product_count + 2

    if product_count > 13:
        print("Sheet full, can't add product, saving file & terminating program.")
        report.save("Internal Report.xlsx")
        invoice.save("Invoice.xlsx")
        exit(0)

    # Writes data to the Invoice and Report sheets
    invoice_sheet.cell(row=invoice_index, column=2).value = PRODUCT_CODES["Linear Slot Diffuser at 45Deg Angle"]
    invoice_sheet.cell(row=invoice_index, column=3).value = "Linear Slot Diffuser at 45Deg Angle"
    invoice_sheet.cell(row=invoice_index, column=7).value = lsd_length
    invoice_sheet.cell(row=invoice_index, column=8).value = quantity
    invoice_sheet.cell(row=invoice_index, column=12).value = unit_price
    invoice_sheet.cell(row=invoice_index, column=13).value = unit_price - (unit_price * discount / 100)
    invoice_sheet.cell(row=invoice_index, column=14).value = unit_price * quantity

    report_sheet.cell(row=report_index, column=1).value = "LSD_45^"
    report_sheet.cell(row=report_index, column=2).value = quantity
    report_sheet.cell(row=report_index, column=3).value = str(round(total_cost, 2) * quantity) + "SAR"
    report_sheet.cell(row=report_index, column=4).value = str(round(material_cost, 2)) + "SAR"
    report_sheet.cell(row=report_index, column=5).value = str(round(labor_cost, 2)) + "SAR"
    report_sheet.cell(row=report_index, column=6).value = str(round(overhead_cost)) + "SAR"
    report_sheet.cell(row=report_index, column=7).value = str(outer_frame_size) + "mm"
    report_sheet.cell(row=report_index, column=8).value = str(outer_frame_price * outer_frame_size / 1000) + "SAR"
    report_sheet.cell(row=report_index, column=9).value = str(inner_frame_size) + "mm"
    report_sheet.cell(row=report_index, column=10).value = str(inner_frame_price * inner_frame_size / 1000) + "SAR"
    report_sheet.cell(row=report_index, column=11).value = str(louver_size) + "mm"
    report_sheet.cell(row=report_index, column=12).value = str(louver_price * louver_size / 1000) + "SAR"
    report_sheet.cell(row=report_index, column=13).value = str(pipe_size) + "mm"
    report_sheet.cell(row=report_index, column=14).value = str(pipe_price * pipe_size / 1000) + "SAR"
    report_sheet.cell(row=report_index, column=15).value = str(space_bar_size) + "mm"
    report_sheet.cell(row=report_index, column=16).value = str(space_bar_price * space_bar_size / 1000) + "SAR"
    report_sheet.cell(row=report_index, column=17).value = str(round(powder_weight, 2)) + "kg"
    report_sheet.cell(row=report_index, column=18).value = str(round(powder_weight, 2) * POWDER_PRICE_PER_KG / 1000) + "SAR"

    # Ask for another product and act accordingly
    try:
        if (input("Product Added.\nDo you wish to add another product? (y: yes, anything else: exit): ")
                .lower() in ["y", "yes"]):
            print("PRODUCT --> PRODUCT_CODE:\n")

            for product_name in PRODUCT_CODES:
                print(product_name + " --> " + PRODUCT_CODES[product_name])

            code = input("\nEnter Product Code: ")
            quantity = int(input("Enter Quantity: "))

            if code not in PRODUCT_CODES.values():
                raise ValueError

            add_product(code, discount, quantity, product_count=product_count+1, current_workbooks=(invoice, report))

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
