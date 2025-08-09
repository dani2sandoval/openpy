"""Command line interface for the billing system."""
from __future__ import annotations

import argparse

from .database import Database


def main() -> None:
    parser = argparse.ArgumentParser(description="Billing system CLI")
    sub = parser.add_subparsers(dest="command")

    cust_parser = sub.add_parser("add-customer", help="Add a new customer")
    cust_parser.add_argument("name")
    cust_parser.add_argument("email")

    prod_parser = sub.add_parser("add-product", help="Add a new product")
    prod_parser.add_argument("name")
    prod_parser.add_argument("price", type=float)

    inv_parser = sub.add_parser("create-invoice", help="Create a new invoice")
    inv_parser.add_argument("customer_id", type=int)

    item_parser = sub.add_parser("add-item", help="Add an item to an invoice")
    item_parser.add_argument("invoice_id", type=int)
    item_parser.add_argument("product_id", type=int)
    item_parser.add_argument("quantity", type=int)

    sub.add_parser("list", help="List all invoice items")

    args = parser.parse_args()
    db = Database()

    if args.command == "add-customer":
        cid = db.add_customer(args.name, args.email)
        print(f"Customer {cid} added")
    elif args.command == "add-product":
        pid = db.add_product(args.name, args.price)
        print(f"Product {pid} added")
    elif args.command == "create-invoice":
        iid = db.create_invoice(args.customer_id)
        print(f"Invoice {iid} created")
    elif args.command == "add-item":
        db.add_invoice_item(args.invoice_id, args.product_id, args.quantity)
        print("Item added to invoice")
    elif args.command == "list":
        for (iid, cname, cemail, pname, quantity, price) in db.list_invoices():
            total = quantity * price
            print(f"Invoice {iid}: {cname} <{cemail}> - {pname} x{quantity} = ${total:.2f}")
    else:
        parser.print_help()

    db.close()


if __name__ == "__main__":
    main()
