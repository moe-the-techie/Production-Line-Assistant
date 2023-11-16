import products


def main():

    try:
        print("PRODUCT --> PRODUCT_CODE:\n")

        for product_name in products.PRODUCT_CODES:
            print(product_name + " --> " + products.PRODUCT_CODES[product_name])

        product = input("\nEnter Product Code: ")
        quantity = int(input("Enter quantity: "))
        discount = float(input("Enter discount percentage (number only e.g. 15): "))

        if product not in products.PRODUCT_CODES.values():
            raise ValueError

        products.add_product(product, discount, quantity)

    except ValueError:
        print("Invalid value entered\nProcess terminated.")
        exit(1)


if __name__ == "__main__":
    main()
