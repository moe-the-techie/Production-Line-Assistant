# This file will have the functions needed to calculate the costs for each product/product type
# I'm thinking of a method for each product where it returns a formatted string that shows all the costs
# We will also place the constant values in here and ensure we can't and won't use them elsewhere
# We'll also need functions to modify the certain attributes of each function and update their values

N_CORNERS = 4
POWDER_PRICE_PER_KG = 15
SHEET_SIZE = 2440 * 1220


def linear_slot_diffuser(n_slots, gap_size, lsd_length):
    """
    Calculates the costs of production for the Linear Slot Diffuser.
    :param n_slots: integer representing the number of slots
    :param gap_size: float representing the size of the gap in millimeters
    :param lsd_length: float representing the length of the linear slot diffuser in millimeters
    :return: A formatted string highlighting the various costs of producing an LSD with the given properties
    """

    # Prices are specific for LSD per 6 meters and the change every 3 months approximately.
    outer_frame_price = 6.5
    inner_frame_price = 7
    louver_price = 3.5
    pipe_price = 2

    # Prices per unit
    sheet_price = 220
    hanging_clamp_price = 0.5
    corner_price = 0.5

    # Calculations for: outer & inner frames, space bar, pipe

    outer_frame_thickness = 4.4
    inner_frame_thickness = 1.2
    n_inner_frames = n_slots - 1
    space_bar_size = gap_size + 16

    pipe_size = 2 * outer_frame_thickness + inner_frame_thickness * n_inner_frames + space_bar_size * n_slots

    n_pipes = int(round((lsd_length - 200)) / 300) + 1
    n_space_bars = n_slots * n_pipes

    end_cap_size = 2 * pipe_size

    outer_frame_size = lsd_length * 2 + 10 + 380 + end_cap_size

    inner_frame_size = lsd_length * n_inner_frames

    n_louvers = n_dampers = n_hanging_clamps = n_slots * 2

    louver_size = n_louvers * lsd_length

    # Calculate the percentage of aluminum sheet used for the dampers.
    sheet_percentage_used = (lsd_length + 10) * space_bar_size * n_dampers / SHEET_SIZE

    powder_weight_kg = 0.0001333333333 * lsd_length

    # Cost calculations

    material_cost = ((outer_frame_size / 6000) * outer_frame_price + (inner_frame_size / 6000) * inner_frame_price +
                     (louver_size / 6000) * louver_price + (pipe_size / 6000) * pipe_price + sheet_percentage_used *
                     sheet_price + corner_price * N_CORNERS + n_hanging_clamps * hanging_clamp_price + powder_weight_kg
                     * POWDER_PRICE_PER_KG)

    labor_cost = (material_cost * 0.45) * 0.75

    overhead_cost = labor_cost * 1.5

    total_cost = labor_cost + material_cost + overhead_cost

    # Return full report

    return ("\n(PRODUCT DESC)\n\nMATERIALS:\nOuter Frame: {outer}mm\nInner Frame: {inner}mm\nLouver ({louver_c}pcs): "
            "{louver}mm\nPipe ({pipe_c}pipes): {pipe}\nSpace bar ({space_bar_c}pcs): {space_bar}\nEnd Cap: {end_cap}\n"
            "Sheet: {sheet_percent}%\nPowder: {powder}kg\n\nCOSTS:\nMaterial: {material}\nLabor: {labor}\nOverhead: "
            "{overhead}\nTotal: {total}").format(outer=round(outer_frame_size, 2), inner=round(inner_frame_size, 2),
                                                 louver_c=n_louvers, louver=round(louver_size, 2), pipe_c=n_pipes,
                                                 pipe=round(pipe_size, 2), space_bar_c=n_space_bars,
                                                 space_bar=round(space_bar_size, 2), end_cap=round(end_cap_size, 2),
                                                 sheet_percent=sheet_percentage_used, powder=round(powder_weight_kg, 2),
                                                 material=round(material_cost, 2), labor=round(labor_cost, 2),
                                                 overhead=round(overhead_cost, 2), total=round(total_cost, 2))

    # Calculates the primary cost and returns it alongside the product description.
    # prim_cost = direct_material + direct_labor + overhead
    # Calculates the customer cost which is for now prim_cost + 0.25 * prim_cost.
    # Ask for discount percentage to apply if any.
