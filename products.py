# This file will have the functions needed to calculate the costs for each product/product type
# I'm thinking of a method for each product where it returns a formatted string that shows all the costs
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


def add_product(code, discount_inputted, current_workbook=None):
    """
    Adds a product to the current workbook by picking the right method for the code entered.
    :param discount_inputted: a float of the number part of the discount percentage to be applied to all products
    :param code: string representing product code
    :param current_workbook: existing workbook to which we add the product
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

            linear_slot_diffuser(n_slots, gap_size, lsd_length, current_workbook)

    except ValueError:

        if current_workbook is None:
            print("Invalid value entered\nProcess terminated.")

        else:
            current_workbook.save("output.xlsx")
            print("Invalid value entered\nOutput file saved. Process terminated.")
        exit(1)


def linear_slot_diffuser(n_slots, gap_size, lsd_length, current_workbook=None):
    """
    Calculates the costs of production for the Linear Slot Diffuser and exports it into an Excel file.
    :param n_slots: integer representing the number of slots
    :param gap_size: float representing the size of the gap in millimeters
    :param lsd_length: float representing the length of the linear slot diffuser in millimeters
    :param current_workbook: an openpyxl reference to the current workbook we're appending products into
    :return:
    """

    # Prices are specific for LSD per 6 meters and the change every 3 months approximately
    outer_frame_price = 6.5
    inner_frame_price = 7
    louver_price = 3.5
    pipe_price = 2

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

    powder_weight_kg = 0.0001333333333 * lsd_length

    # Cost calculations

    material_cost = ((outer_frame_size / 6000) * outer_frame_price + (inner_frame_size / 6000) * inner_frame_price +
                     (louver_size / 6000) * louver_price + (pipe_size / 6000) * pipe_price + sheet_percentage_used *
                     sheet_price + corner_price * N_CORNERS + n_hanging_clamps * hanging_clamp_price + powder_weight_kg
                     * POWDER_PRICE_PER_KG)

    labor_cost = material_cost * 0.3375

    overhead_cost = labor_cost * 1.5

    total_cost = labor_cost + material_cost + overhead_cost

    unit_price = material_cost * 4

    if current_workbook is None:
        wb = openpyxl.load_workbook("template.xlsx")

    else:
        wb = current_workbook

    ws = wb.active

    ws.cell(row=31, column=16).value = str(discount) + "%"

    i = 19

    while i <= 29:
        if ws.cell(row=i, column=2).value is None:
            break

        i += 1

    if i == 30:
        print("Sheet full, can't add product, saving file & terminating program.")
        wb.save("output.xlsx")
        exit(0)

    # Entering the data into the right row, notice that the 1's represent the quantity and should be updated
    # to whatever value that ends up taking on
    ws.cell(row=i, column=2).value = PRODUCT_CODES["Linear Slot Diffuser at 45Deg Angle"]
    ws.cell(row=i, column=3).value = "Linear Slot Diffuser at 45Deg Angle"
    ws.cell(row=i, column=7).value = lsd_length
    ws.cell(row=i, column=8).value = 1
    ws.cell(row=i, column=12).value = unit_price
    ws.cell(row=i, column=13).value = unit_price - (unit_price * discount / 100)
    ws.cell(row=i, column=14).value = unit_price * 1

    print(("\nPRODUCT ADDED TO INVOICE:\n\nMATERIALS:\nOuter Frame 22242: {outer}mm\nInner Frame 22241: {"
           "inner}mm\nLouver 22245 ({louver_c}pcs): {louver}mm\nPipe ({pipe_c}pipes): {pipe}\nSpace bar "
           "({space_bar_c}pcs): {space_bar}\nEnd Cap: {end_cap}\nSheet: {sheet_percent}%\nPowder:"
           "{powder}kg\n\nCOSTS:\nMaterial: {material}\nLabor:{labor}\nOverhead: {overhead}\nTotal: "
           "{total}\n").format(outer=round(outer_frame_size, 2),
                               inner=round(inner_frame_size, 2),
                               louver_c=n_louvers,
                               louver=louver_size,
                               pipe_c=n_pipes,
                               pipe=round(pipe_size, 2),
                               space_bar_c=n_space_bars,
                               space_bar=round(space_bar_size, 2),
                               end_cap=round(end_cap_size, 2),
                               sheet_percent=sheet_percentage_used,
                               powder=round(powder_weight_kg, 2),
                               material=round(material_cost, 2),
                               labor=round(labor_cost, 2),
                               overhead=round(overhead_cost, 2),
                               total=round(total_cost, 2)))

    try:
        if input("Do you wish to add another product? (y: yes, anything else: exit): ").lower() in ["y", "yes"]:
            print("PRODUCT --> PRODUCT_CODE:\n")

            for product_name in PRODUCT_CODES:
                print(product_name + " --> " + PRODUCT_CODES[product_name])

            code = input("\nEnter Product Code: ")

            if code not in PRODUCT_CODES.values():
                raise ValueError

            add_product(code, discount, wb)

        else:

            print("Saving output file and terminating process")
            wb.save("output.xlsx")
            exit(0)

    except ValueError:

        wb.save("output.xlsx")
        print("Invalid value entered\nOutput file saved. Process terminated.")
        exit(1)

    # Calculates the primary cost and returns it alongside the product description.
    # prim_cost = direct_material + direct_labor + overhead
    # Calculates the customer cost which is for now prim_cost + 0.25 * prim_cost.
    # Ask for discount percentage to apply if any.
