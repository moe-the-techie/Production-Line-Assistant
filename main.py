import products

PRODUCT_CODES = {
    "Linear Slot Diffuser at 45Deg Angle": "LSD45^"
}


def main():

    try:
        print("PRODUCT --> PRODUCT_CODE:\n")

        for product_name in PRODUCT_CODES:
            print(product_name + " --> " + PRODUCT_CODES[product_name])

        product = input("Enter Product Code: ")

        if product not in PRODUCT_CODES.values():
            raise ValueError

        if product == "LSD45^":

            n_slots = int(input("Enter number of slots: "))
            gap_size = float(input("Enter gap size in mm: "))
            lsd_length = float(input("Enter length of linear slot diffuser in mm: "))

            # Ensure all values are positive before calculating
            if n_slots <= 0 or gap_size <= 0 or lsd_length <= 0:
                raise ValueError

            print(products.linear_slot_diffuser(n_slots, gap_size, lsd_length))

    except ValueError:
        print("Invalid value entered\nProcess terminated.")
        exit(1)


if __name__ == "__main__":
    main()
